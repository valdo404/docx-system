/**
 * Post-build fix for @astrojs/cloudflare Pages compatibility.
 *
 * `astro build` with @astrojs/cloudflare generates:
 *   - dist/_worker.js/wrangler.json with Worker-specific config (main, rules)
 *   - dist/_worker.js/entry.mjs as the worker entry point
 *   - .wrangler/deploy/config.json redirecting to the generated config
 *
 * Pages expects the entry at `_worker.js/index.js` and doesn't support
 * Worker config keys like `main`. We fix this by:
 *   1. Renaming entry.mjs → index.js
 *   2. Removing the generated wrangler.json and deploy redirect
 *
 * Wrangler then uses the project-level wrangler.jsonc for bindings (D1, KV)
 * and finds the entry point at the expected location.
 */
import { renameSync, rmSync } from 'node:fs';

renameSync('dist/_worker.js/entry.mjs', 'dist/_worker.js/index.js');
rmSync('dist/_worker.js/wrangler.json', { force: true });
rmSync('.wrangler/deploy/config.json', { force: true });
