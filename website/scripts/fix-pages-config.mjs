/**
 * Post-build fix for @astrojs/cloudflare Pages compatibility.
 *
 * The adapter generates dist/_worker.js/wrangler.json with Worker-specific
 * keys (main, rules) and a reserved ASSETS binding that confuse
 * `wrangler pages deploy`. Remove the file entirely — the project-level
 * wrangler.jsonc provides all needed config (bindings, vars, etc.).
 */
import { unlinkSync } from 'node:fs';

unlinkSync('dist/_worker.js/wrangler.json');
