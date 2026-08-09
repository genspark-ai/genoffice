// Rejects new raw colors in renderer CSS: UI chrome must go through the design
// tokens in packages/ui/src/tokens.css (see CLAUDE.md). Raw values are allowed
// only on custom-property definition lines (token/accent definitions), so the
// check is diff-based and never flags pre-existing lines.
import { existsSync, readFileSync } from 'node:fs'
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

const CSS_SCOPE = /^apps\/[^/]+\/src\/renderer\/.+\.css$/
const RAW_COLOR = /#[0-9a-fA-F]{3,8}\b|\b(?:rgb|rgba|hsl|hsla)\(/
const TOKEN_DEF = /^\s*--[\w-]+:/
const COMMENT_LINE = /^\s*(\/\*|\*)/

function git(commandArgs, { allowFailure = false } = {}) {
  const result = spawnSync('git', commandArgs, {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'inherit'],
  })
  if (result.status !== 0 && !allowFailure) process.exit(result.status ?? 1)
  return result
}

const repoRoot = git(['rev-parse', '--show-toplevel']).stdout.trim()
const violations = new Map()

function flag(file, line, text) {
  violations.set(`${file}:${line}`, text.trim())
}

function checkLine(file, line, text) {
  if (!RAW_COLOR.test(text)) return
  if (TOKEN_DEF.test(text) || COMMENT_LINE.test(text)) return
  flag(file, line, text)
}

function scanDiff(diffArgs) {
  const out = git(['diff', '-U0', ...diffArgs, '--', '*.css']).stdout
  let file = null
  let newLine = 0
  for (const line of out.split('\n')) {
    if (line.startsWith('+++ b/')) {
      const path = line.slice(6)
      file = CSS_SCOPE.test(path) ? path : null
    } else if (line.startsWith('@@')) {
      const m = /\+(\d+)/.exec(line)
      newLine = m ? Number(m[1]) : 0
    } else if (file && line.startsWith('+')) {
      checkLine(file, newLine, line.slice(1))
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
      `Theme-color base ${baseRef} is not available in this checkout; ` +
        'checking working-tree changes only.',
    )
}
scanDiff([])
scanDiff(['--cached'])

for (const file of git(['ls-files', '--others', '--exclude-standard']).stdout.split('\n')) {
  if (!CSS_SCOPE.test(file) || !existsSync(join(repoRoot, file))) continue
  readFileSync(join(repoRoot, file), 'utf8')
    .split('\n')
    .forEach((text, i) => checkLine(file, i + 1, text))
}

if (violations.size === 0) {
  console.log('No new raw colors in renderer CSS.')
  process.exit(0)
}
console.error(
  'New raw colors in renderer CSS. UI chrome must use the design tokens from\n' +
    'packages/ui/src/tokens.css (raw values are only allowed when defining a\n' +
    'custom property). See the theming rules in CLAUDE.md.\n',
)
for (const [where, text] of violations) console.error(`  ${where}  ${text}`)
process.exit(1)
