// Keeps the Docs ribbon and right-click menu on the shared command table: a
// click handler added to those files must run a command, or only open and
// close the control's own popup, or say why it is exempt. Diff-based, so the
// handlers that were there before are not flagged.
//
// Exempt a handler with a comment on one of the three lines above it:
//   // command-exempt: <why this is not a command>
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

const SCOPE =
  /^apps\/docs\/src\/renderer\/components\/(Ribbon\.tsx|ribbon-[\w-]+\.tsx|ContextMenu\.tsx)$/
const HANDLER = /\bonClick\s*[=:]/
// runs a command, directly or through a tab's pick helper
const RUNS_COMMAND = /\b(runCommand|onCommand|pick\w*)\(|[=:]\s*\{?\s*\w+\.run\s*[},]/
// opens or closes the control's own dropdown, submenu or dialog: UI state, not a command
const OWN_POPUP = /\b(toggleDropdown|setDropdown|setSubmenu|set\w*(Open|Dialog|Menu)\w*)\(/
const EXEMPT = /command-exempt:\s*\S/
const BODY_LOOKAHEAD = 14
const EXEMPT_LOOKBEHIND = 3

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
const fileLines = new Map()

function linesOf(file) {
  if (!fileLines.has(file)) {
    const path = join(repoRoot, file)
    fileLines.set(file, existsSync(path) ? readFileSync(path, 'utf8').split('\n') : [])
  }
  return fileLines.get(file)
}

/** the handler's text: its own line, and what follows when the line leaves it unfinished */
function handlerText(lines, index) {
  const first = lines[index] ?? ''
  if (!/([{(]|=>)\s*$/.test(first)) return first
  const indent = first.search(/\S/)
  const body = [first]
  for (const next of lines.slice(index + 1, index + BODY_LOOKAHEAD)) {
    body.push(next)
    // back at the handler's own indentation: its closing line
    if (next.trim() !== '' && next.search(/\S/) <= indent) break
  }
  return body.join('\n')
}

function checkLine(file, line) {
  const lines = linesOf(file)
  const index = line - 1
  const text = lines[index] ?? ''
  if (!HANDLER.test(text)) return
  const handler = handlerText(lines, index)
  if (RUNS_COMMAND.test(handler) || OWN_POPUP.test(handler)) return
  const above = lines.slice(Math.max(0, index - EXEMPT_LOOKBEHIND), index + 1).join('\n')
  if (EXEMPT.test(above)) return
  violations.set(`${file}:${line}`, text.trim())
}

function scanDiff(diffArgs) {
  const out = git(['diff', '-U0', ...diffArgs, '--', '*.tsx']).stdout
  let file = null
  let newLine = 0
  for (const line of out.split('\n')) {
    if (line.startsWith('+++ b/')) {
      const path = line.slice(6)
      file = SCOPE.test(path) ? path : null
    } else if (line.startsWith('@@')) {
      const m = /\+(\d+)/.exec(line)
      newLine = m ? Number(m[1]) : 0
    } else if (file && line.startsWith('+')) {
      checkLine(file, newLine)
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
      `Docs-commands base ${baseRef} is not available in this checkout; ` +
        'checking working-tree changes only.',
    )
}
scanDiff([])
scanDiff(['--cached'])

for (const file of git(['ls-files', '--others', '--exclude-standard']).stdout.split('\n')) {
  if (!SCOPE.test(file)) continue
  linesOf(file).forEach((_, i) => checkLine(file, i + 1))
}

if (violations.size === 0) {
  console.log('No new click handlers outside the Docs command table.')
  process.exit(0)
}
console.error(
  'New click handlers in the Docs ribbon or right-click menu that do not run a\n' +
    'command. Register the action in apps/docs/src/shared/commands.ts, give it a\n' +
    'handler in apps/docs/src/renderer/commands/menu-commands.ts and run it with\n' +
    'runCommand, so the menu bar, the ribbon and the right-click menu share it.\n' +
    'A handler that is not a command takes a "// command-exempt: <why>" comment\n' +
    'on one of the three lines above it.\n',
)
for (const [where, text] of violations) console.error(`  ${where}  ${text}`)
process.exit(1)
