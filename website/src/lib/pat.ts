import { Kysely } from 'kysely';
import { D1Dialect } from 'kysely-d1';

const TOKEN_PREFIX = 'dxs_';

interface PatRecord {
  id: string;
  tenantId: string;
  name: string;
  tokenHash: string;
  tokenPrefix: string;
  createdAt: string;
  lastUsedAt: string | null;
  expiresAt: string | null;
}

export interface PatInfo {
  id: string;
  name: string;
  tokenPrefix: string;
  createdAt: string;
  lastUsedAt: string | null;
  expiresAt: string | null;
}

export interface CreatePatResult {
  pat: PatInfo;
  token: string; // Only returned on creation, never stored
}

function getKysely(db: D1Database) {
  return new Kysely<{ personal_access_token: PatRecord }>({
    dialect: new D1Dialect({ database: db }),
  });
}

async function hashToken(token: string): Promise<string> {
  const encoder = new TextEncoder();
  const data = encoder.encode(token);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
}

function generateToken(): string {
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);
  const randomPart = Array.from(bytes)
    .map((b) => b.toString(16).padStart(2, '0'))
    .join('');
  return `${TOKEN_PREFIX}${randomPart}`;
}

export async function createPat(
  db: D1Database,
  tenantId: string,
  name: string,
  expiresAt?: Date,
): Promise<CreatePatResult> {
  const kysely = getKysely(db);

  const token = generateToken();
  const tokenHash = await hashToken(token);
  const tokenPrefix = token.slice(0, 12); // "dxs_" + 8 chars

  const now = new Date().toISOString();
  const id = crypto.randomUUID();

  const record: PatRecord = {
    id,
    tenantId,
    name,
    tokenHash,
    tokenPrefix,
    createdAt: now,
    lastUsedAt: null,
    expiresAt: expiresAt ? expiresAt.toISOString() : null,
  };

  await kysely.insertInto('personal_access_token').values(record).execute();

  return {
    pat: {
      id,
      name,
      tokenPrefix,
      createdAt: now,
      lastUsedAt: null,
      expiresAt: record.expiresAt,
    },
    token, // Return the raw token only once
  };
}

export async function listPats(
  db: D1Database,
  tenantId: string,
): Promise<PatInfo[]> {
  const kysely = getKysely(db);

  const records = await kysely
    .selectFrom('personal_access_token')
    .select(['id', 'name', 'tokenPrefix', 'createdAt', 'lastUsedAt', 'expiresAt'])
    .where('tenantId', '=', tenantId)
    .orderBy('createdAt', 'desc')
    .execute();

  return records;
}

export async function deletePat(
  db: D1Database,
  tenantId: string,
  patId: string,
): Promise<boolean> {
  const kysely = getKysely(db);

  const result = await kysely
    .deleteFrom('personal_access_token')
    .where('id', '=', patId)
    .where('tenantId', '=', tenantId)
    .executeTakeFirst();

  return (result.numDeletedRows ?? 0) > 0;
}

export async function renamePat(
  db: D1Database,
  tenantId: string,
  patId: string,
  newName: string,
): Promise<boolean> {
  const kysely = getKysely(db);

  const result = await kysely
    .updateTable('personal_access_token')
    .set({ name: newName })
    .where('id', '=', patId)
    .where('tenantId', '=', tenantId)
    .executeTakeFirst();

  return (result.numUpdatedRows ?? 0) > 0;
}

export interface VerifyPatResult {
  valid: boolean;
  tenantId?: string;
  patId?: string;
  expired?: boolean;
}

export async function verifyPat(
  db: D1Database,
  token: string,
): Promise<VerifyPatResult> {
  if (!token.startsWith(TOKEN_PREFIX)) {
    return { valid: false };
  }

  const kysely = getKysely(db);
  const tokenHash = await hashToken(token);

  const record = await kysely
    .selectFrom('personal_access_token')
    .select(['id', 'tenantId', 'expiresAt'])
    .where('tokenHash', '=', tokenHash)
    .executeTakeFirst();

  if (!record) {
    return { valid: false };
  }

  // Check expiration
  if (record.expiresAt && new Date(record.expiresAt) < new Date()) {
    return { valid: false, expired: true };
  }

  // Update lastUsedAt
  await kysely
    .updateTable('personal_access_token')
    .set({ lastUsedAt: new Date().toISOString() })
    .where('id', '=', record.id)
    .execute();

  return {
    valid: true,
    tenantId: record.tenantId,
    patId: record.id,
  };
}
