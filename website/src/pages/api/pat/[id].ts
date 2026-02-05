import type { APIRoute } from 'astro';
import { deletePat, renamePat } from '../../../lib/pat';

export const prerender = false;

// DELETE /api/pat/:id - Delete a PAT
export const DELETE: APIRoute = async (context) => {
  const tenant = context.locals.tenant;
  if (!tenant) {
    return new Response(JSON.stringify({ error: 'Tenant not found' }), {
      status: 404,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  const patId = context.params.id;
  if (!patId) {
    return new Response(JSON.stringify({ error: 'PAT ID required' }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  const { env } = await import('cloudflare:workers');
  const deleted = await deletePat((env as unknown as Env).DB, tenant.id, patId);

  if (!deleted) {
    return new Response(JSON.stringify({ error: 'PAT not found' }), {
      status: 404,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  return new Response(JSON.stringify({ success: true }), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
};

// PATCH /api/pat/:id - Rename a PAT
export const PATCH: APIRoute = async (context) => {
  const tenant = context.locals.tenant;
  if (!tenant) {
    return new Response(JSON.stringify({ error: 'Tenant not found' }), {
      status: 404,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  const patId = context.params.id;
  if (!patId) {
    return new Response(JSON.stringify({ error: 'PAT ID required' }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  let body: { name?: string };
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

  const { env } = await import('cloudflare:workers');
  const updated = await renamePat((env as unknown as Env).DB, tenant.id, patId, name);

  if (!updated) {
    return new Response(JSON.stringify({ error: 'PAT not found' }), {
      status: 404,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  return new Response(JSON.stringify({ success: true }), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
};
