// Keeps font metrics on the manifest path. (a) fonts.generated.css and
// line-factors.generated.ts must equal what tools/build-font-aliases.mjs
// emits from packages/font-metrics/measurements; (b) every manifest source
// carries a kind, a date and a probe id or font files, and no new `legacy`
// entry joins the grandfathered list; (c) no added line outside the generated
// files declares size-adjust / *-override in renderer CSS or a factor literal
// inside knownLineFactor(). Diff-based like check-theme-colors: pre-existing
// lines never fail.
import { spawnSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

import { GENERATED_CSS, generate, loadManifest, validateManifest } from './build-font-aliases.mjs'

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

/** legacy entries that predate the manifest; a new one needs a measurement instead */
const GRANDFATHERED_LEGACY = new Set([
  'family:Noto Naskh Arabic W',
  'family:Noto Sans Arabic W',
  'rule:kai-latin-named',
  'rule:times',
  'rule:georgia',
  'rule:arial',
  'rule:calibri',
  'rule:tahoma',
  'rule:consolas',
  'rule:courier',
  'rule:segoe-other',
  'rule:ethiopic',
  'rule:inter',
  'rule:ubuntu',
  'rule:avenir',
  'rule:lucida-sans',
  'rule:bradley-hand',
  'rule:open-sans',
  'rule:shonar-bangla',
])

const LINE_METRICS = 'apps/docs/src/renderer/line-metrics.ts'
const RENDERER_CSS = /^apps\/docs\/src\/renderer\/.*\.css$/
const CSS_METRIC = /^\s*(size-adjust|ascent-override|descent-override|line-gap-override)\s*:/
const TS_FACTOR = /\breturn\s+\d+\.\d+\b/
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
const problems = []

// (a) generated files match the manifest
const manifest = loadManifest(repoRoot)
problems.push(...validateManifest(manifest))
if (problems.length === 0) {
  for (const [rel, text] of Object.entries(generate(repoRoot))) {
    let current = ''
    try {
      current = readFileSync(join(repoRoot, rel), 'utf8')
    } catch {
      /* missing counts as stale */
    }
    if (current !== text) problems.push(`${rel} is stale: run node tools/build-font-aliases.mjs`)
  }
}

// (b) legacy entries are grandfathered, never added
for (const fam of manifest.families)
  if (fam.source?.kind === 'legacy' && !GRANDFATHERED_LEGACY.has(`family:${fam.family}`))
    problems.push(
      `${fam.file}: new legacy source; measure the value (probe, hmtx or corpus) instead`,
    )
for (const rule of manifest.lineFactors.rules || [])
  if (rule.source?.kind === 'legacy' && !GRANDFATHERED_LEGACY.has(`rule:${rule.id}`))
    problems.push(`line-factors.json rule ${rule.id}: new legacy source; measure the value instead`)

// (c) metric literals stay out of hand-written files
const fileCache = new Map()
const linesOf = (file) => {
  if (!fileCache.has(file))
    fileCache.set(file, readFileSync(join(repoRoot, file), 'utf8').split('\n'))
  return fileCache.get(file)
}

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

function offending(file, line) {
  const lines = linesOf(file)
  const text = lines[line - 1] ?? ''
  if (COMMENT_LINE.test(text)) return null
  if (RENDERER_CSS.test(file) && file !== GENERATED_CSS && CSS_METRIC.test(text))
    return 'metric descriptor outside fonts.generated.css; add it to the manifest'
  if (file === LINE_METRICS && TS_FACTOR.test(text)) {
    const range = knownLineFactorRange(lines)
    if (range && line >= range[0] && line <= range[1])
      return 'factor literal in knownLineFactor(); add a rule to line-factors.json'
  }
  return null
}

const violations = new Map()
function scanDiff(diffArgs) {
  const out = git([
    'diff',
    '-U0',
    '--diff-algorithm=histogram',
    ...diffArgs,
    '--',
    'apps/docs/src/renderer',
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
      if (file === LINE_METRICS || RENDERER_CSS.test(file)) {
        const why = offending(file, newLine)
        if (why)
          violations.set(`${file}:${newLine}`, `${linesOf(file)[newLine - 1].trim()}  (${why})`)
      }
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
      `Generated-fonts base ${baseRef} is not available in this checkout; ` +
        'checking working-tree changes only.',
    )
}
scanDiff([])
scanDiff(['--cached'])
for (const [where, text] of violations) problems.push(`${where}  ${text}`)

if (problems.length === 0) {
  console.log('Generated font files match the manifest; no hand-written metric values added.')
  process.exit(0)
}
console.error(
  'Font metrics must flow manifest -> generator -> generated files\n' +
    '(packages/font-metrics/measurements, tools/build-font-aliases.mjs):\n',
)
for (const p of problems) console.error(`  ${p}`)
process.exit(1)
