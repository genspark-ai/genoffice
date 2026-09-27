import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { defaultRegistry, runCli } from '../src/cli'
import { commandHelp } from '../src/registry'

// Help/registry/README/skill sync: every command and option in the registry must
// surface in `genoffice help`, in the CLI README examples and in the genoffice
// agent skill, so adding a flag without documenting it fails loudly instead of
// drifting. Fast and deterministic: in-process runCli plus file reads only.

const REPO = resolve(__dirname, '../../..')
const README_PATH = resolve(__dirname, '..', 'README.md')
const SKILL_CANDIDATES = [
  resolve(REPO, 'skills/genoffice/SKILL.md'),
  resolve(__dirname, '..', 'skills/genoffice/SKILL.md'),
]

function readFirst(paths: string[], label: string): { path: string; text: string } {
  for (const path of paths) {
    try {
      return { path, text: readFileSync(path, 'utf-8') }
    } catch {
      // try the next candidate location
    }
  }
  throw new Error(`${label} not found (tried ${paths.join(', ')})`)
}

const readme = readFileSync(README_PATH, 'utf-8')
const skillFile = readFirst(SKILL_CANDIDATES, 'genoffice SKILL.md')

function fencesOf(markdown: string): string {
  return [...markdown.matchAll(/```[\s\S]*?```/g)].map((m) => m[0]).join('\n')
}

const readmeFences = fencesOf(readme)

/** Lowercase with every whitespace run collapsed: help pads, tables re-pad. */
function norm(text: string): string {
  return text.toLowerCase().replace(/\s+/g, ' ')
}

/** `--name` with a trailing boundary, so `--max` does not match `--max-rows`. */
function optionPattern(name: string): RegExp {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return new RegExp(`--${escaped}(?![a-z0-9-])`)
}

async function helpFor(argv: string[]): Promise<string> {
  const out: string[] = []
  const err: string[] = []
  await runCli(argv, { io: { stdout: (t) => out.push(t), stderr: (t) => err.push(t) } })
  return out.join('\n')
}

/**
 * Host-only commands the agent skill legitimately omits: `install-cli` puts the
 * launcher on the PATH, not a document workflow an agent reads the skill for.
 * It must stay in `genoffice help` and the CLI README. Enforced in both
 * directions below, so extending the skill (or dropping a command) fails loudly
 * instead of silently changing coverage.
 */
const SKILL_COMMAND_OMISSIONS = new Map([['install-cli', 'host setup, not a document workflow']])

/**
 * Flags the README example fences legitimately omit: the fences show one common
 * invocation per command while `genoffice help <command>` is the complete
 * reference. Enforced exactly like the skill omissions: a new flag must gain a
 * fence example or be added here deliberately, and removing a flag without
 * updating this set fails.
 */
const README_FENCE_OPTION_OMISSIONS = new Set([
  'password',
  'layouts',
  'max-chars',
  'best-effort',
  'stop-on-error',
  'isolation',
  'max-rows',
  'where',
  'stats',
  'styles',
  'sections',
  'fields',
  'notes',
  'track',
  'author',
  'index',
  'fingerprint',
  'size',
  'ref',
  'model',
  'block',
])

/**
 * `--list` is an alias for the `skill list` subcommand form, which is what the
 * README fences and the skill document. Accept the verb form as coverage.
 */
const OPTION_VERB_FORMS = new Map([['list', /\bskill\s+list\b/]])

function covered(text: string, option: string): boolean {
  if (optionPattern(option).test(text)) return true
  const verb = OPTION_VERB_FORMS.get(option)
  return verb !== undefined && verb.test(text)
}

describe('cli help/registry/readme/skill sync', () => {
  it('registers each command once with a usage line naming it', () => {
    const defs = defaultRegistry().list()
    const names = defs.map((d) => d.name)
    expect(new Set(names).size, `duplicate command names: ${names.join(', ')}`).toBe(names.length)
    for (const def of defs) {
      expect(def.summary.trim().length, `${def.name} needs a summary`).toBeGreaterThan(0)
      expect(def.usage.startsWith(def.name), `${def.name} usage must start with its name`).toBe(
        true,
      )
    }
  })

  it('lists every registry command and global flag in genoffice help', async () => {
    const global = norm(await helpFor(['help']))
    const names = defaultRegistry()
      .list()
      .map((d) => d.name)
    const missing = names.filter((n) => !global.includes(n))
    expect(missing, `genoffice help is missing commands: ${missing.join(', ')}`).toEqual([])
    for (const flag of ['--json', '--help', '--version']) {
      expect(global.includes(flag), `genoffice help is missing global flag ${flag}`).toBe(true)
    }
  })

  it('documents every option of every command in its help output', async () => {
    const gaps: string[] = []
    for (const def of defaultRegistry().list()) {
      const built = norm(commandHelp(def))
      const shown = norm(await helpFor(['help', def.name]))
      if (!shown.includes(norm(`usage: genoffice ${def.usage}`))) {
        gaps.push(`${def.name}: 'help ${def.name}' omits its usage line`)
      }
      for (const opt of def.options ?? []) {
        const pattern = optionPattern(opt.name)
        if (!pattern.test(built)) gaps.push(`${def.name}: commandHelp() omits --${opt.name}`)
        if (!pattern.test(shown)) gaps.push(`${def.name}: 'help ${def.name}' omits --${opt.name}`)
      }
    }
    expect(gaps, gaps.join('\n')).toEqual([])
  })

  it('shows every command in packages/cli/README.md and its code fences', () => {
    const body = norm(readme)
    const fence = norm(readmeFences)
    const names = defaultRegistry()
      .list()
      .map((d) => d.name)
    const missingFile = names.filter((n) => !body.includes(norm(`genoffice ${n}`)))
    const missingFence = names.filter((n) => !fence.includes(norm(`genoffice ${n}`)))
    expect(missingFile, `packages/cli/README.md never mentions: ${missingFile.join(', ')}`).toEqual(
      [],
    )
    expect(
      missingFence,
      `packages/cli/README.md code fences show no example for: ${missingFence.join(', ')}`,
    ).toEqual([])
  })

  it('covers every command in the genoffice skill except documented host-only ones', () => {
    const body = norm(skillFile.text)
    const names = defaultRegistry()
      .list()
      .map((d) => d.name)
    const missing = names.filter((n) => !body.includes(norm(`genoffice ${n}`)))
    const unexpected = missing.filter((n) => !SKILL_COMMAND_OMISSIONS.has(n))
    const stale = [...SKILL_COMMAND_OMISSIONS.keys()].filter((n) => !missing.includes(n))
    expect(
      unexpected,
      `${skillFile.path} documents no such command: ${unexpected.join(', ')}. Add a command-table row or extend SKILL_COMMAND_OMISSIONS deliberately.`,
    ).toEqual([])
    expect(
      stale,
      `stale SKILL_COMMAND_OMISSIONS (now documented in ${skillFile.path}): ${stale.join(', ')}. Remove them from the map.`,
    ).toEqual([])
  })

  it('covers every registry option in the README fences and the skill', () => {
    const fence = norm(readmeFences)
    const skillBody = norm(skillFile.text)
    const owners = new Map<string, string[]>()
    for (const def of defaultRegistry().list()) {
      for (const opt of def.options ?? []) {
        const list = owners.get(opt.name) ?? []
        list.push(def.name)
        owners.set(opt.name, list)
      }
    }
    const describe = (options: string[]) =>
      options.map((o) => `--${o} (${(owners.get(o) ?? []).join(', ')})`).join(', ')
    const fenceMissing = [...owners.keys()].filter((o) => !covered(fence, o))
    const fenceUnexpected = fenceMissing.filter((o) => !README_FENCE_OPTION_OMISSIONS.has(o))
    const fenceStale = [...README_FENCE_OPTION_OMISSIONS].filter((o) => !fenceMissing.includes(o))
    const skillMissing = [...owners.keys()].filter((o) => !covered(skillBody, o))
    expect(
      fenceUnexpected,
      `packages/cli/README.md code fences show no such flag: ${describe(fenceUnexpected)}. Add an example or extend README_FENCE_OPTION_OMISSIONS deliberately.`,
    ).toEqual([])
    expect(
      fenceStale,
      `stale README_FENCE_OPTION_OMISSIONS (now shown in a fence): ${describe(fenceStale)}. Remove them from the set.`,
    ).toEqual([])
    expect(
      skillMissing,
      `${skillFile.path} documents no such flag: ${describe(skillMissing)}. Add it to the command table or prose.`,
    ).toEqual([])
  })
})
