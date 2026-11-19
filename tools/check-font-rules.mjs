// Keeps face / line-factor / font-chain decisions inside
// apps/docs/src/renderer/text-layout/font-rules/ (and the resolver / manifests
// they read). New lines in line-metrics.ts and the other renderer emitters must
// not decide a face themselves: no fontTable hint reads, no resolver calls, no
// availability probes, no font-name regexes, no CSS chain literals, no numeric
// line factors keyed by name. A rule goes into font-rules/ with its Word probe
// (DESIGN-fonts §3.4). The resolver's own name rows live in
// measurements/substitutions.json and its hint-conditioned steps in
// measurements/substitution-steps.json: a new `[regex, 'Face']` literal or a
// new `if (hints.charset ...)` branch in resolve.ts fails too. Every line of the
// guarded files is checked, not only the added ones; `// font-rules: allow
// <truth>` on or next to a line exempts it.
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

/** the resolver keeps its name rows and hint-conditioned steps in measurements/*.json, not in code */
const RESOLVER = 'packages/font-metrics/src/resolve.ts'
const RESOLVER_RULES = [
  [
    /\[\s*(?:\/[^\n]+?\/[a-z]*|[A-Z_]+_RE)\s*,\s*'[^']+'\s*(?:,\s*(?:true|false))?\s*\]/,
    'substitution row in code; add it to packages/font-metrics/measurements/substitutions.json',
  ],
  [
    /^\s*(?:\}\s*else\s+)?if\b.*\bhints\??\.(?:charset|altName|panose|family|csb0)\b/,
    'hint-conditioned branch in code; add a step to packages/font-metrics/measurements/substitution-steps.json',
  ],
]

const GUARDED = new Set([
  RESOLVER,
  'apps/docs/src/renderer/line-metrics.ts',
  'apps/docs/src/renderer/doc-style-css.ts',
  'apps/docs/src/renderer/editor/para-layout-spec.ts',
  'apps/docs/src/renderer/editor/extensions.ts',
  'apps/docs/src/renderer/editor/protected-render.ts',
  'apps/docs/src/renderer/editor/marks.ts',
  'apps/docs/src/renderer/text-layout/run-faces.ts',
  'apps/docs/src/renderer/text-layout/para-line-box.ts',
  'apps/docs/src/renderer/text-layout/strut-face.ts',
])
const RULES = [
  [/\b(?:fontTableHint|docFontTable|setResolverFontTable)\s*\(/, 'fontTable hint read'],
  [/\bresolveDeclaredFace\s*\(/, 'resolver call'],
  [/\b(?:isFontAvailable|declaredFaceMissing|isBundledFont)\s*\(/, 'availability probe'],
  [/\b(?:LINE_FACTOR_RULES|LINE_FACTORS|LINE_FACTOR_PINS)\b/, 'line-factor table read'],
  [/\b(?:STAND_IN_CHAINS|FACE_CHAIN_RULES|namedFaceChain)\b/, 'chain table read'],
  // a quoted family chain or a stand-in alias literal
  [/'[^']+ GO'|'GenOffice [^']+'|'Noto (?:Sans|Serif) CJK/, 'font chain literal'],
  // a regex over font names (vendor / script family words)
  [
    /\/[^/\n]*(?:simsun|mincho|gothic|batang|malgun|gulim|yahei|dengxian|calibri|cambria|arial|helvetica|segoe|aptos|times|courier|verdana|georgia|tahoma)[^/\n]*\/[a-z]*/i,
    'font-name regex',
  ],
]
const COMMENT_LINE = /^\s*(\/\/|\/\*|\*)/
const ALLOW = /font-rules: allow\b/

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
    if (/^\s*(?:import|export)\b/.test(text) || /^\s*\} from '/.test(text)) continue
    if (ALLOW.test(text) || ALLOW.test(at(line - 1)) || ALLOW.test(at(line + 1))) continue
    for (const [re, why] of file === RESOLVER ? RESOLVER_RULES : RULES) {
      if (re.test(text)) {
        violations.set(`${file}:${line}`, `[${why}] ${text.trim()}`)
        break
      }
    }
  }
}

for (const file of GUARDED) {
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
  console.log('No face decisions outside text-layout/font-rules/.')
  process.exit(0)
}
console.error(
  'New face / line-factor / font-chain decisions outside font-rules. fontTable reads,\n' +
    'resolver calls, availability probes, font-name regexes, chain literals and\n' +
    'line-factor table reads belong in apps/docs/src/renderer/text-layout/font-rules/\n' +
    '(with their Word probe), packages/font-metrics/src/resolve.ts or the measurement\n' +
    'manifests; the emitter maps the result. Put `// font-rules: allow <truth date>` on\n' +
    '(or next to) a line that is genuinely an emitter mapping.\n',
)
for (const [where, text] of violations) console.error(`  ${where}  ${text}`)
process.exit(1)
