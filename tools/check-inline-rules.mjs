// Keeps inline text-layout decisions inside apps/docs/src/renderer/text-layout/.
// New lines in the renderer's emitters (the live decoration plugins, the static
// protected DOM, the marks, doc-style-css, styles.css) must not decide anything
// about text themselves: no script tests, no face-name predicates, no raw CJK
// code-point ranges, no settings.xml compat reads, no ad-hoc Decoration.inline.
// A rule goes into text-layout/ (with its Word probe) and the emitter maps the
// result. Every line of the emitters is checked, not only the added ones.
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

const EMITTERS = new Set([
  'apps/docs/src/renderer/editor/extensions.ts',
  'apps/docs/src/renderer/editor/protected-render.ts',
  'apps/docs/src/renderer/editor/marks.ts',
  'apps/docs/src/renderer/editor/decoration-extensions.ts',
  'apps/docs/src/renderer/editor/script-fonts.ts',
  'apps/docs/src/renderer/doc-style-css.ts',
  'apps/docs/src/renderer/styles.css',
])
const RULES = [
  [
    /\b(?:textHasCjk|textHasHangul|textHasEastAsian|textHasHangulChars|hasBalancedSpaceContext)\s*\(/,
    'script test of paragraph text',
  ],
  [
    /\b(?:isCjkFontName|isCjkNamedFont|isEaLayoutFace|isEastAsianFontName|isEastAsianFaceName|isKoreanFontName|isJapaneseFontName)\s*\(/,
    'face-name predicate',
  ],
  [(text) => cjkCodePoints(text), 'raw CJK code-point range'],
  [
    /\b(?:balanceDbcsSpacing|useFELayout|compressPunctuation|docDefaults\??\.lang|docDefaults\??\.eastAsiaLang)\b/,
    'settings.xml compat read',
  ],
  [/\bDecoration\.inline\s*\(/, 'ad-hoc inline decoration'],
]
const COMMENT_LINE = /^\s*(\/\/|\/\*|\*)/
const CJK_BLOCK = (cp) =>
  (cp >= 0x1100 && cp <= 0x11ff) ||
  (cp >= 0x2e80 && cp <= 0x9fff) ||
  (cp >= 0xac00 && cp <= 0xd7ff) ||
  (cp >= 0xf900 && cp <= 0xfaff) ||
  (cp >= 0xff00 && cp <= 0xffef)
/** a \uXXXX escape (either case) or a literal character inside the CJK blocks */
function cjkCodePoints(text) {
  for (const m of text.matchAll(/\\u([0-9a-fA-F]{4})/g))
    if (CJK_BLOCK(parseInt(m[1], 16))) return true
  for (const ch of text) if (CJK_BLOCK(ch.codePointAt(0))) return true
  return false
}

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

const ALLOW = /inline-rules: allow\b/

/** added lines of one file: the allow marker sits on the line itself or, where prettier
 *  moved the comment, on the adjacent line of the file as it is now */
function checkLines(file, entries) {
  const path = join(repoRoot, file)
  const fileLines = existsSync(path) ? readFileSync(path, 'utf8').split('\n') : []
  const at = (n) => fileLines[n - 1] ?? ''
  for (const { line, text } of entries) {
    if (COMMENT_LINE.test(text)) continue
    if (ALLOW.test(text) || ALLOW.test(at(line - 1)) || ALLOW.test(at(line + 1))) continue
    for (const [re, why] of RULES) {
      if (typeof re === 'function' ? re(text) : re.test(text)) {
        violations.set(`${file}:${line}`, `[${why}] ${text.trim()}`)
        break
      }
    }
  }
}

for (const file of EMITTERS) {
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
  console.log('No inline text-layout decisions outside text-layout/.')
  process.exit(0)
}
console.error(
  'New inline text-layout decisions in an emitter. Script tests, face-name predicates,\n' +
    'CJK code-point ranges, compat reads and ad-hoc inline decorations belong in\n' +
    'apps/docs/src/renderer/text-layout/ (with their Word probe); the emitter maps the\n' +
    'result. Put `// inline-rules: allow` on (or just above) a line that is genuinely an emitter mapping.\n',
)
for (const [where, text] of violations) console.error(`  ${where}  ${text}`)
process.exit(1)
