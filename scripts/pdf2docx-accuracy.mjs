#!/usr/bin/env node
/** PDF → editable accuracy bench (invoice / report class).
 *
 *  Converts a deterministic corpus (banded e-invoice, the genoffice#2669 redacted
 *  real sample, bordered + borderless reports) with BOTH the genoffice
 *  pipeline and LibreOffice, scores every ground-truth field's survival
 *  (intact in one cell / split across cells / missing) plus the detail
 *  band's column count, and writes a markdown scorecard:
 *
 *    node scripts/pdf2docx-accuracy.mjs                    # print scorecard
 *    node scripts/pdf2docx-accuracy.mjs out/scorecard.md   # write scorecard
 *
 *  The bench itself is tests/accuracy-bench.test.ts, gated behind
 *  PDF2DOCX_ACCURACY=1 so CI never pays for it.
 */
import { spawnSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const out = process.argv[2]
const result = spawnSync(
  'npx',
  ['vitest', 'run', 'packages/pdf2docx/tests/accuracy-bench.test.ts'],
  {
    cwd: root,
    stdio: 'inherit',
    env: {
      ...process.env,
      PDF2DOCX_ACCURACY: '1',
      ...(out ? { PDF2DOCX_ACCURACY_OUT: out } : {}),
    },
  },
)
if (out && result.status === 0) console.log(`scorecard written to ${out}`)
process.exit(result.status ?? 1)
