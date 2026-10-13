/**
 * Runs the Word export verification against the server converter.
 * Usage: node scripts/verify-docx-export.mjs
 */
import { execFileSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
execFileSync('pnpm', ['exec', 'tsx', 'scripts/verify-docx-export.ts'], {
  cwd: root,
  stdio: 'inherit',
});
