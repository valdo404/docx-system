import type { APIRoute } from 'astro';
import { createPat, listPats } from '../../../lib/pat';

export const prerender = false;

// GET /api/pat - List all PATs for the current tenant
export const GET: APIRoute = async (context) => {
  const tenant = context.locals.tenant;
  if (!tenant) {
    return new Response(JSON.stringify({ error: 'Tenant not found' }), {
      status: 404,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  const { env } = await import('cloudflare:workers');
  const pats = await listPats((env as unknown as Env).DB, tenant.id);

  return new Response(JSON.stringify({ pats }), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
};

// POST /api/pat - Create a new PAT
export const POST: APIRoute = async (context) => {
  const tenant = context.locals.tenant;
  if (!tenant) {
    return new Response(JSON.stringify({ error: 'Tenant not found' }), {
      status: 404,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  let body: { name?: string; expiresAt?: string };
  try {
    body = await context.request.json();
  } catch {
    return new Response(JSON.stringify({ error: 'Invalid JSON' }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  const name = body.name?.trim();
  if (!name) {
    return new Response(JSON.stringify({ error: 'Name is required' }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  const expiresAt = body.expiresAt ? new Date(body.expiresAt) : undefined;

  const { env } = await import('cloudflare:workers');
  const result = await createPat((env as unknown as Env).DB, tenant.id, name, expiresAt);

  return new Response(JSON.stringify(result), {
    status: 201,
    headers: { 'Content-Type': 'application/json' },
  });
};
