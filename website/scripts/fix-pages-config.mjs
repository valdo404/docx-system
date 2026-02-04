/**
 * Post-build fix for @astrojs/cloudflare Pages compatibility.
 *
 * `astro build` with @astrojs/cloudflare produces two artefacts that break
 * `wrangler pages deploy`:
 *
 *   1. dist/_worker.js/wrangler.json — Worker-specific keys (main, rules)
 *      and a reserved ASSETS binding that are invalid for Pages.
 *   2. .wrangler/deploy/config.json  — a redirect that tells wrangler to
 *      read (1) instead of the project-level config.
 *
 * Removing only (1) causes wrangler to find the stale redirect in (2) and
 * crash because the target no longer exists.  Removing both lets wrangler
 * fall back to the project-level wrangler.jsonc which carries the real
 * bindings (D1, KV, vars).
 *
 * rmSync with { force: true } is a no-op when the file is absent, so this
 * stays forward-compatible if a future adapter version stops generating
 * either file.
 */
import { rmSync } from 'node:fs';

rmSync('dist/_worker.js/wrangler.json', { force: true });
rmSync('.wrangler/deploy/config.json', { force: true });
