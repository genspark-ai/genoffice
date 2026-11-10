// Rejects new line-metric numbers that cite no measurement. Every added or
// changed `size-adjust` / `*-override` in apps/docs/src/renderer/fonts/fonts.css
// and every numeric return or literal inside knownLineFactor() in
// apps/docs/src/renderer/line-metrics.ts must have, within the 12 preceding
// lines of the resulting file, a comment naming where the value came from
// (a Word probe, a Word PDF, a corpus measurement, the face's hhea table or a
// dated note). Diff-based like check-theme-colors: pre-existing lines never fail.
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { spawnSync } from 'node:child_process'

const args = process.argv.slice(2)
let baseRef = process.env.FORMAT_BASE_REF || ''
if (args[0] === '--base') {
  args.shift()
  baseRef = args.shift() || ''
}
if (/^0+$/.test(baseRef)) baseRef = ''
if (args.length > 0) {
  console.error(`Unexpected argument: ${args[0]}`)
  process.exit(2)
}

const FONTS_CSS = 'apps/docs/src/renderer/fonts/fonts.css'
const LINE_METRICS = 'apps/docs/src/renderer/line-metrics.ts'
const WINDOW = 12
const CSS_METRIC = /^\s*(size-adjust|ascent-override|descent-override|line-gap-override)\s*:/
const TS_FACTOR = /\breturn\s+\d+\.\d+\b|^\s*\d+\.\d+\s*,?\s*$|:\s*\d+\.\d+\s*,?\s*(\/\/.*)?$/
const PROVENANCE = /probe|Word PDF|corpus|hhea|20\d\d-\d\d-\d\d/
const COMMENT_LINE = /^\s*(\/\/|\/\*|\*)/

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

/** 1-based line range of the knownLineFactor function body */
function knownLineFactorRange(lines) {
  const start = lines.findIndex((l) => /^function knownLineFactor\(/.test(l))
  if (start === -1) return null
  let depth = 0
  for (let i = start; i < lines.length; i++) {
    for (const ch of lines[i]) {
      if (ch === '{') depth++
      else if (ch === '}' && --depth === 0) return [start + 1, i + 1]
    }
  }
  return [start + 1, lines.length]
}

/** true when `line` (1-based) is a metric line without provenance nearby */
function unexplained(file, lines, line) {
  const text = lines[line - 1]
  if (COMMENT_LINE.test(text)) return false
  if (file === FONTS_CSS) {
    if (!CSS_METRIC.test(text)) return false
  } else {
    const range = knownLineFactorRange(lines)
    if (!range || line < range[0] || line > range[1]) return false
    if (!TS_FACTOR.test(text)) return false
  }
  const from = Math.max(0, line - 1 - WINDOW)
  return !lines.slice(from, line - 1).some((l) => PROVENANCE.test(l))
}

const violations = new Map()
const fileCache = new Map()
const linesOf = (file) => {
  if (!fileCache.has(file))
    fileCache.set(file, readFileSync(join(repoRoot, file), 'utf8').split('\n'))
  return fileCache.get(file)
}

function scanDiff(diffArgs) {
  const out = git([
    'diff',
    '-U0',
    '--diff-algorithm=histogram',
    ...diffArgs,
    '--',
    FONTS_CSS,
    LINE_METRICS,
  ]).stdout
  let file = null
  let newLine = 0
  for (const line of out.split('\n')) {
    if (line.startsWith('+++ b/')) file = line.slice(6)
    else if (line.startsWith('+++ /dev/null')) file = null
    else if (line.startsWith('@@')) {
      const m = /\+(\d+)/.exec(line)
      newLine = m ? Number(m[1]) : 0
    } else if (file && line.startsWith('+')) {
      const lines = linesOf(file)
      if (newLine <= lines.length && unexplained(file, lines, newLine))
        violations.set(`${file}:${newLine}`, lines[newLine - 1].trim())
      newLine++
    }
  }
}

if (baseRef) {
  const verification = git(['rev-parse', '--verify', `${baseRef}^{commit}`], {
    allowFailure: true,
  })
  if (verification.status === 0) scanDiff([`${baseRef}...HEAD`])
  else
    console.warn(
      `Font-metrics base ${baseRef} is not available in this checkout; ` +
        'checking working-tree changes only.',
    )
}
scanDiff([])
scanDiff(['--cached'])

if (violations.size === 0) {
  console.log('Every new font-metric value cites its measurement.')
  process.exit(0)
}
console.error(
  'New font-metric values without provenance. Within the 12 lines above each\n' +
    'size-adjust / *-override in fonts.css and each factor in knownLineFactor(),\n' +
    'add a comment saying where the number was measured (probe, Word PDF,\n' +
    'corpus, hhea, or a YYYY-MM-DD note). See the docs-fidelity rule: metric\n' +
    'numbers are measured in Word before they are written.\n',
)
for (const [where, text] of violations) console.error(`  ${where}  ${text}`)
process.exit(1)
