#!/usr/bin/env node
/** Deep sweep of the seeded DOCX round-trip fuzz suite.
 *
 *  CI runs the default 24 seeds (apps/docs/tests/roundtrip-fuzz.test.ts);
 *  this script sweeps a much larger seed space locally, e.g.:
 *
 *    node scripts/roundtrip-sweep.mjs        # 200 seeds
 *    node scripts/roundtrip-sweep.mjs 2000   # 2000 seeds
 *
 *  Every test name carries its seed, so a failure reproduces exactly via
 *  buildFuzzDocx(seed). Exit code mirrors vitest's. */
import { spawnSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const seeds = Number(process.argv[2] ?? 200)
if (!Number.isInteger(seeds) || seeds < 1) {
  console.error(
    `usage: node scripts/roundtrip-sweep.mjs <seeds> (got ${process.argv[2] ?? 'nothing'})`,
  )
  process.exit(2)
}
console.log(`round-trip sweep: ${seeds} seeds`)
const result = spawnSync('npx', ['vitest', 'run', 'apps/docs/tests/roundtrip-fuzz.test.ts'], {
  cwd: root,
  stdio: 'inherit',
  env: { ...process.env, FUZZ_SEEDS: String(seeds) },
})
process.exit(result.status ?? 1)
