// Keeps the pagination preview / export path a consumer of the canvas PageCut
// data (apps/docs/src/renderer/pagination-pass.ts, pagination-cuts.ts). New
// lines in the preview sheets, the PDF export and the headless readiness code
// must not re-decide a page fact: no second slicing, no block measurement, no
// page lookup by Y, no header/footer variant or page-number re-derivation, and
// no layout reads (querySelector / getBoundingClientRect / getComputedStyle,
// generic-typed calls included) outside apps/docs/src/renderer/preview-canvas-snapshot.ts,
// the one module that reads the live canvas and the mounted sheets for them.
// A read that really belongs in a reader takes `// page-cut: allow` on the line
// or the adjacent one. Every line of the readers is checked, not only the added ones.
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { spawnSync } from 'node:child_process'

const args = process.argv.slice(2)
// --base is accepted for the callers that pass it to every check; the scan is whole-file
if (args[0] === '--base') args.splice(0, 2)
if (args.length > 0) {
  console.error(`Unexpected argument: ${args[0]}`)
  process.exit(2)
}

const READERS = new Set([
  'apps/docs/src/renderer/components/PaginationPreview.tsx',
  'apps/docs/src/renderer/headless-export.ts',
  'apps/docs/src/renderer/file-actions.ts',
])
const RULES = [
  [
    /\b(?:sliceWithLineSplit|computeSectionedSlices(?:F2)?|computePageSlices|buildPageCuts|measureBlocks|fillLineBoxes)\s*\(/,
    'second slicing / measurement; the canvas pass owns it (pagination-pass)',
  ],
  [/\b(?:pinnedFloatPage|pageOfY|pageAt)\s*\(/, 'page lookup by Y; read cut.floats / PageCut'],
  [
    /\b(?:hfVariantOf|formatPageNumber|pageNoAfterOf|pageNumbers|sectionFirstPages|hfReservedHeightPx|sectionHfHeights|hfVariantHeightsPx)\s*\(/,
    'header/footer or page-number fact re-derived; read cut.hf / cut.display',
  ],
  [
    /\b(?:querySelector(?:All)?|getBoundingClientRect|getComputedStyle)\s*(?:<[^>()]*>)?\s*\(/,
    'layout read in the preview/export path',
  ],
  [/import\.meta\.env\.DEV/, 'DEV-only branch in the preview/export path; make it a counter'],
]
const COMMENT_LINE = /^\s*(\/\/|\/\*|\*)/
const ALLOW = /page-cut: allow\b/

function git(commandArgs, { allowFailure = false } = {}) {
  const result = spawnSync('git', commandArgs, {
    encoding: 'utf8',
    maxBuffer: 256 * 1024 * 1024,
    stdio: ['ignore', 'pipe', 'inherit'],
  })
  if (result.error) {
    console.error(`git ${commandArgs[0]} failed: ${result.error.message}`)
    process.exit(1)
  }
  if (result.status !== 0 && !allowFailure) process.exit(result.status ?? 1)
  return result
}

const repoRoot = git(['rev-parse', '--show-toplevel']).stdout.trim()
const violations = new Map()

function checkLines(file, entries) {
  const path = join(repoRoot, file)
  const fileLines = existsSync(path) ? readFileSync(path, 'utf8').split('\n') : []
  const at = (n) => fileLines[n - 1] ?? ''
  for (const { line, text } of entries) {
    if (COMMENT_LINE.test(text)) continue
    if (ALLOW.test(text) || ALLOW.test(at(line - 1)) || ALLOW.test(at(line + 1))) continue
    for (const [re, why] of RULES) {
      if (re.test(text)) {
        violations.set(`${file}:${line}`, `[${why}] ${text.trim()}`)
        break
      }
    }
  }
}

for (const file of READERS) {
  const path = join(repoRoot, file)
  if (!existsSync(path)) continue
  checkLines(
    file,
    readFileSync(path, 'utf8')
      .split('\n')
      .map((text, i) => ({ line: i + 1, text })),
  )
}

if (violations.size > 0) {
  console.error(
    'Preview/export lines re-decide a page fact or read the DOM themselves. Consume the canvas PageCut / pagination-pass data, or mark a genuine clone-geometry read with `// page-cut: allow`:',
  )
  for (const [where, what] of [...violations].sort()) console.error(`  ${where}: ${what}`)
  process.exit(1)
}
console.log('Page-cut readers check passed.')
