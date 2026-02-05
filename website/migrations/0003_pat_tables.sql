CREATE TABLE IF NOT EXISTS "personal_access_token" (
    "id" TEXT PRIMARY KEY NOT NULL,
    "tenantId" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "tokenHash" TEXT NOT NULL UNIQUE,
    "tokenPrefix" TEXT NOT NULL,
    "createdAt" TEXT NOT NULL,
    "lastUsedAt" TEXT,
    "expiresAt" TEXT,
    FOREIGN KEY ("tenantId") REFERENCES "tenant"("id") ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS "idx_pat_tenantId" ON "personal_access_token"("tenantId");
CREATE INDEX IF NOT EXISTS "idx_pat_tokenHash" ON "personal_access_token"("tokenHash");
