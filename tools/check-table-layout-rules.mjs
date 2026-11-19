// Keeps table column-width decisions in their two homes: what the document declares
// (w:tblGrid / w:tcW / w:tblW / w:tblLayout and the arbitration between them) in
// packages/docx-engine/src/table-widths.ts, everything after parse in
// apps/docs/src/renderer/table-layout/. The callers (convert.ts, table-geometry.ts,
// the emitters, the docx-engine parser) must not compute or rewrite column widths
// themselves: no min/max-content calls, no grid rewrites, no width constants, no
// tcW / gridCol predicates, no width XML reads. Every line of the callers is
// checked, not only the added ones.
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

const CALLERS = new Set([
  'apps/docs/src/renderer/editor/convert.ts',
  'apps/docs/src/renderer/editor/table-geometry.ts',
  'apps/docs/src/renderer/editor/protected-render.ts',
  'apps/docs/src/renderer/editor/extensions.ts',
  'packages/docx-engine/src/parse.ts',
])
const RULES = [
  [
    /\b(?:minContentColTwips|maxContentColTwips|cellMeasurePx|cellMaxWordPx|maxWordWidthPx|maxLineWidthPx|distributeContentAutofit|reclaimFloorTwips|nestedFloorTwips|reclaimSurplus|scaleToFit|layoutTableColumns|clampColumnWidths|nestedTableMinTwips|nestedTablePreferredTwips)\s*\(/,
    'column width computation',
  ],
  [/\bcolWidths(?:Twips|Pct)\s*(?:=[^=>]|:\s*(?!null|undefined))/, 'column grid rewrite'],
  [
    /\b(?:MIN_CONTENT_SLACK|MEASURED_EDGE_PX|MIN_COL_TWIPS|DEFAULT_CELL_MAR|DEFAULT_TABLE_CELL_MAR)\b/,
    'width constant',
  ],
  [
    /\b(?:isUniformGrid|tcwColumnWidths|tableLacksTcW|noWrapPinnedTwips|usableTableWidths|reconcileGridColumns)\s*\(/,
    'tcW / grid predicate',
  ],
  [
    /\b(?:contentAutofit|layoutGrid|autoLayout|dxaWidth|fixedLayout|noWrapPinnedTwips|widthPct)\s*=\s*(?!=)/,
    'width flag assignment',
  ],
  [/['"]w:(?:tcW|gridCol|tblW|tblLayout|noWrap|tblGrid)['"]/, 'width XML read'],
]
const COMMENT_LINE = /^\s*(\/\/|\/\*|\*)/
/** a bare name inside an import / export list: using a module export is not a decision */
const IMPORT_LIST_LINE = /^\s*(?:type\s+)?[A-Za-z_$][\w$]*,?\s*$/
const ALLOW = /table-layout: allow\b/

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

/** added lines of one file: the allow marker sits on the line itself or, where prettier
 *  moved the comment, on the adjacent line of the file as it is now */
function checkLines(file, entries) {
  const path = join(repoRoot, file)
  const fileLines = existsSync(path) ? readFileSync(path, 'utf8').split('\n') : []
  const at = (n) => fileLines[n - 1] ?? ''
  for (const { line, text } of entries) {
    if (COMMENT_LINE.test(text) || IMPORT_LIST_LINE.test(text)) continue
    if (ALLOW.test(text) || ALLOW.test(at(line - 1)) || ALLOW.test(at(line + 1))) continue
    for (const [re, why] of RULES) {
      if (re.test(text)) {
        violations.set(`${file}:${line}`, `[${why}] ${text.trim()}`)
        break
      }
    }
  }
}

for (const file of CALLERS) {
  const path = join(repoRoot, file)
  if (!existsSync(path)) continue
  checkLines(
    file,
    readFileSync(path, 'utf8')
      .split('\n')
      .map((text, i) => ({ line: i + 1, text })),
  )
}

if (violations.size === 0) {
  console.log('No table column-width decisions outside table-widths.ts and table-layout/.')
  process.exit(0)
}
console.error(
  'Table column-width decisions outside their homes. What the document declares belongs in\n' +
    'packages/docx-engine/src/table-widths.ts; min/max-content calls, grid rewrites, width\n' +
    'constants and width flags after parse belong in apps/docs/src/renderer/table-layout/\n' +
    '(with their Word probe and a node unit test); the caller maps the result.\n' +
    'Put `// table-layout: allow` on (or just above) a line that is genuinely a caller mapping.\n',
)
for (const [where, text] of violations) console.error(`  ${where}  ${text}`)
process.exit(1)
