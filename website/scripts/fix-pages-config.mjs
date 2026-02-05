/**
 * Post-build fix for @astrojs/cloudflare Pages compatibility.
 *
 * `astro build` with @astrojs/cloudflare generates:
 *   - dist/_worker.js/wrangler.json with Worker-specific config (main, rules)
 *   - dist/_worker.js/entry.mjs as the worker entry point
 *   - dist/client/ containing static assets (HTML, CSS, JS)
 *   - .wrangler/deploy/config.json redirecting to the generated config
 *
 * Pages expects:
 *   - Static assets at the root of the output directory
 *   - Worker entry at `_worker.js/index.js`
 *   - No Worker-specific config keys like `main`
 *
 * We fix this by:
 *   1. Moving static assets from dist/client/* to dist/
 *   2. Renaming entry.mjs → index.js
 *   3. Removing the generated wrangler.json and deploy redirect
 */
import { renameSync, rmSync, cpSync } from 'node:fs';

// Move static assets to dist root (Pages expects them there)
cpSync('dist/client', 'dist', { recursive: true });
rmSync('dist/client', { recursive: true, force: true });

// Rename worker entry to what Pages expects
renameSync('dist/_worker.js/entry.mjs', 'dist/_worker.js/index.js');

// Remove generated configs so wrangler uses project-level wrangler.jsonc
rmSync('dist/_worker.js/wrangler.json', { force: true });
rmSync('.wrangler/deploy/config.json', { force: true });
