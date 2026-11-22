#!/usr/bin/env node
/** Round-trip fidelity bench — xlsx / pptx scorecard driver.
 *
 *  Round-trips a deterministic corpus with BOTH the genoffice pipelines
 *  (the streamed xlsx save; the pptx engine's open → save) and LibreOffice,
 *  scores every ground-truth attribute's survival (intact / degraded /
 *  lost), and writes a markdown scorecard:
 *
 *    node scripts/roundtrip-accuracy.mjs           # print both scorecards
 *    node scripts/roundtrip-accuracy.mjs xlsx      # one domain
 *    node scripts/roundtrip-accuracy.mjs out.md    # write both scorecards
 *
 *  The benches are apps/sheets-adjacent gated tests, like the pdf2docx
 *  accuracy bench (#2756): gated behind env vars so CI never pays for them.
 */
import { spawnSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const arg = process.argv[2] ?? ''
const domains = arg === 'xlsx' || arg === 'pptx' ? [arg] : ['xlsx', 'pptx']
const out = arg.endsWith('.md') || !arg ? arg || undefined : undefined

const tests = {
  xlsx: {
    gate: 'XLSX_ROUNDTRIP_ACCURACY',
    out: 'XLSX_ROUNDTRIP_ACCURACY_OUT',
    file: 'packages/xlsx-gateway/tests/accuracy-roundtrip.test.ts',
  },
  pptx: {
    gate: 'PPTX_ROUNDTRIP_ACCURACY',
    out: 'PPTX_ROUNDTRIP_ACCURACY_OUT',
    file: 'packages/pptx-engine/tests/accuracy-roundtrip.test.ts',
  },
}

// both domains sharing one output path would overwrite: suffix the domain
const outFor = (domain) =>
  out === undefined || domains.length === 1 ? out : out.replace(/\.md$/, '') + `-${domain}.md`

let status = 0
for (const domain of domains) {
  const spec = tests[domain]
  const domainOut = outFor(domain)
  const result = spawnSync('npx', ['vitest', 'run', spec.file], {
    cwd: root,
    stdio: 'inherit',
    env: {
      ...process.env,
      [spec.gate]: '1',
      ...(domainOut ? { [spec.out]: domainOut } : {}),
    },
  })
  if (result.status !== 0) status = result.status
}
process.exit(status)
