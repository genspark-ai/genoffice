// skills/*/SKILL.md are copied into users' agent directories (Claude Code,
// Codex, ...) and only replaced when metadata.version grows, so a body change
// that keeps the version would never reach anyone. Diff-based like the other
// checks: compares each changed SKILL.md with its base revision.
import { readFileSync } from 'node:fs'
import { basename, dirname, join } from 'node:path'
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

const SKILL_FILE = /^skills\/[^/]+\/SKILL\.md$/

function git(commandArgs, { allowFailure = false } = {}) {
  const result = spawnSync('git', commandArgs, {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', allowFailure ? 'pipe' : 'inherit'],
  })
  if (result.status !== 0 && !allowFailure) process.exit(result.status ?? 1)
  return result
}

const repoRoot = git(['rev-parse', '--show-toplevel']).stdout.trim()
if (!baseRef) baseRef = 'HEAD'

const changed = [
  ...git(['diff', '--name-only', '--diff-filter=AMR', baseRef]).stdout.split('\n'),
  // new skills not yet committed (local runs; CI sees them as committed)
  ...git(['ls-files', '--others', '--exclude-standard', 'skills']).stdout.split('\n'),
].filter((f) => SKILL_FILE.test(f))

/** frontmatter fields the check cares about, and the body with the frontmatter removed */
export function parseSkill(text) {
  const m = /^---\r?\n([\s\S]*?)\r?\n---\r?\n?([\s\S]*)$/.exec(text)
  if (!m) return { name: null, version: null, body: text.trim() }
  const front = m[1]
  const name = /^name:\s*(\S+)\s*$/m.exec(front)?.[1] ?? null
  const version = /^\s+version:\s*['"]?(\d+\.\d+\.\d+)['"]?\s*$/m.exec(front)?.[1] ?? null
  return { name, version, body: m[2].trim() }
}

export function compareSemver(a, b) {
  const pa = a.split('.').map(Number)
  const pb = b.split('.').map(Number)
  for (let i = 0; i < 3; i++) if (pa[i] !== pb[i]) return pa[i] - pb[i]
  return 0
}

const problems = []
for (const file of changed) {
  const head = parseSkill(readFileSync(join(repoRoot, file), 'utf8'))
  const dir = basename(dirname(file))
  if (!head.name) problems.push(`${file}: frontmatter needs a name`)
  else if (head.name !== dir)
    problems.push(`${file}: name "${head.name}" must match its directory "${dir}"`)
  if (!head.version) {
    problems.push(`${file}: frontmatter needs metadata.version (x.y.z)`)
    continue
  }
  const base = git(['show', `${baseRef}:${file}`], { allowFailure: true })
  if (base.status !== 0) continue
  const prev = parseSkill(base.stdout)
  if (!prev.version) continue
  const cmp = compareSemver(head.version, prev.version)
  if (cmp < 0)
    problems.push(`${file}: metadata.version went backwards (${prev.version} -> ${head.version})`)
  else if (cmp === 0 && head.body !== prev.body)
    problems.push(`${file}: body changed but metadata.version is still ${prev.version}; bump it`)
}

if (problems.length) {
  console.error('Skill version check failed:')
  for (const p of problems) console.error(`  ${p}`)
  process.exit(1)
}
console.log(`Skill version check passed (${changed.length} file(s) checked).`)
