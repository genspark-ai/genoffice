import { cpSync, existsSync, readFileSync, readdirSync } from 'node:fs'
import { basename, dirname, isAbsolute, join, resolve } from 'node:path'
import { homedir } from 'node:os'
import { classifySkill, parseSkillFrontmatter } from '@genoffice/agent-core'
import type { FoundSkill } from '../shared/found-skill'

/**
 * Finding the skills a user already has, so GenOffice can offer the ones worth
 * having without asking them to go and copy files.
 *
 * Two places, and the difference matters:
 *
 * - `<default save dir>/skills/<name>/SKILL.md` is *ours*. The user put it
 *   there, so a skill found there is used without asking.
 * - `~/.agents/skills`, `~/.claude/skills` and friends are the coding agents'
 *   own directories. A skill found there is a *candidate*: it is shown with its
 *   description and whether it looks relevant to the formats this app opens,
 *   and only runs if the user picks it. Nothing is copied and nothing runs on
 *   its own — a developer's `.claude` directory holds git, test and deploy
 *   skills, and silently running any of them from inside a word processor would
 *   be the wrong default in both directions.
 */

/** one SKILL.md we found, with everything the palette needs to show it */
export type { FoundSkill }

/** the folder under the save directory that holds skills we may use */
export function skillsRoot(defaultSaveDir: string): string {
  return join(defaultSaveDir, 'skills')
}

/**
 * The coding-agent skill directories on this machine.
 *
 * `~/.agents` is the cross-agent convention and is listed first so a skill
 * living there wins over the same name in one vendor's directory. The rest
 * mirror the table `genoffice skill` already installs into
 * (`packages/cli/src/agent-skills.ts`) — the same machines, the same homes,
 * one list rather than two that can drift.
 */
export function agentSkillDirs(
  env: NodeJS.ProcessEnv = process.env,
  home: string = homedir(),
): { agent: string; dir: string }[] {
  return [
    { agent: 'agents', dir: join(home, '.agents', 'skills') },
    {
      agent: 'claude-code',
      dir: env.CLAUDE_CONFIG_DIR
        ? join(env.CLAUDE_CONFIG_DIR, 'skills')
        : join(home, '.claude', 'skills'),
    },
    {
      agent: 'codex',
      dir: env.CODEX_HOME ? join(env.CODEX_HOME, 'skills') : join(home, '.codex', 'skills'),
    },
    { agent: 'cursor', dir: join(home, '.cursor', 'skills') },
    { agent: 'gemini-cli', dir: join(home, '.gemini', 'skills') },
    { agent: 'copilot', dir: join(home, '.copilot', 'skills') },
    { agent: 'opencode', dir: join(home, '.opencode', 'skills') },
    { agent: 'windsurf', dir: join(home, '.codeium', 'windsurf', 'skills') },
  ]
}

/**
 * One directory of `<root>/<name>/SKILL.md`.
 *
 * A name is skipped when it is not a directory, when SKILL.md is missing, and
 * when the file cannot be read — a permission error on one skill must not hide
 * the other nineteen, so each is read independently and a failure is just an
 * absence.
 */
function readDir(root: string, source: FoundSkill['source'], agent?: string): FoundSkill[] {
  let entries: string[]
  try {
    entries = readdirSync(root)
  } catch {
    return []
  }
  const out: FoundSkill[] = []
  for (const name of entries.sort()) {
    const path = join(root, name, 'SKILL.md')
    if (!existsSync(path)) continue
    let text: string
    try {
      text = readFileSync(path, 'utf8')
    } catch {
      continue
    }
    const fm = parseSkillFrontmatter(text)
    // no name/description frontmatter means it is not a skill we can present
    if (!fm) continue
    out.push({
      name,
      description: fm.description,
      path,
      source,
      ...(agent ? { agent } : {}),
      relevance: classifySkill(fm.name, fm.description),
    })
  }
  return out
}

/**
 * Everything worth offering, ours first.
 *
 * Ours come first and are not de-duplicated against the agents' copies: the
 * user put that folder there, so a same-named skill from `.claude` is not
 * allowed to shadow it. Among the agents' own, the first directory to offer a
 * name wins, which is what makes the `~/.agents` ordering above mean something.
 */
export function findSkills(
  defaultSaveDir: string,
  env?: NodeJS.ProcessEnv,
  home?: string,
): FoundSkill[] {
  const ours = readDir(skillsRoot(defaultSaveDir), 'genoffice')
  const seen = new Set(ours.map((s) => s.name))
  const fromAgents: FoundSkill[] = []
  for (const { agent, dir } of agentSkillDirs(env, home)) {
    for (const skill of readDir(dir, 'agent', agent)) {
      if (seen.has(skill.name)) continue
      seen.add(skill.name)
      fromAgents.push(skill)
    }
  }
  return [...ours, ...fromAgents]
}

/** the skills a user may actually run: the ones they chose to put in our folder */
export function usableSkills(defaultSaveDir: string): FoundSkill[] {
  return readDir(skillsRoot(defaultSaveDir), 'genoffice')
}

/** the body of a skill, for the turn that runs it */
export function readSkillBody(path: string): string {
  const fm = parseSkillFrontmatter(readFileSync(path, 'utf8'))
  return fm?.body ?? ''
}

/**
 * Put a skill the user picked into our own folder, so it is ours to use.
 *
 * This is the only write here, and it happens when a person clicks Import —
 * nothing on this path copies itself just because it was scanned.
 *
 * The whole directory goes rather than the SKILL.md alone: a skill routinely
 * ships reference tables, templates and helper scripts next to its
 * instructions, and half of one reads as a working skill while failing at the
 * step that needed the missing file.
 *
 * The name is taken from the path, never from the caller. Two sources for one
 * destination is how a folder ends up named after something other than what is
 * inside it.
 *
 * Refuses rather than overwrites when the name is taken: a same-named skill in
 * our folder is the one the user put there, and silently replacing it with
 * `.claude`'s copy is the shadowing rule above, just with the winner flipped.
 */
export function importSkill(skillPath: string, defaultSaveDir: string): FoundSkill {
  const name = basename(dirname(skillPath))
  const dest = join(skillsRoot(defaultSaveDir), name)
  if (existsSync(dest)) throw new Error('a skill of that name is already there')
  cpSync(dirname(skillPath), dest, { recursive: true })
  // re-read at the destination rather than trusting the copy: this is the record
  // the UI will show as installed, and it should describe the file now on disk
  const fm = parseSkillFrontmatter(readFileSync(join(dest, 'SKILL.md'), 'utf8'))
  if (!fm) throw new Error('the skill could not be read back')
  return {
    name,
    description: fm.description,
    path: join(dest, 'SKILL.md'),
    source: 'genoffice',
    relevance: classifySkill(fm.name, fm.description),
  }
}

/**
 * Every directory the scan above reads, ours first.
 *
 * Exported so the IPC can name the roots a renderer is allowed to ask about
 * rather than keeping a second list that can drift from this one.
 */
export function knownSkillRoots(
  defaultSaveDir: string,
  env?: NodeJS.ProcessEnv,
  home?: string,
): string[] {
  return [skillsRoot(defaultSaveDir), ...agentSkillDirs(env, home).map((a) => a.dir)]
}

/**
 * Whether a path is one of the SKILL.md files the scan can hand out.
 *
 * `readSkillBody` takes a path from the renderer, so on its own it reads any
 * file on disk and hands back its contents. This is the guard that stops that:
 * the name must be SKILL.md and its directory's directory must be one of the
 * roots — the exact shape `readDir` builds above, so what can be read and what
 * was listed are the same set.
 *
 * The boundary is that exact equality. An escaping path is caught either way:
 * `dirname` only strips one component and never collapses `..`, so
 * `/save/skills/../../.ssh/SKILL.md` keeps its `..` in the grandparent and
 * cannot match a normalized root. Resolving first is what settles a path that
 * climbs out and back in — `/save/skills/../skills/x/SKILL.md` is read as
 * `/save/skills/x/SKILL.md`, the skill it always meant — so the widening stays
 * within the root.
 *
 * Symlinks are deliberately not resolved. The listing follows them (it stats
 * through them), so a linked SKILL.md is one this process already offered to
 * read; resolving here would make listed skills unreadable while stopping
 * nothing a user who can write into their own skills directory could not
 * already do.
 */
export function isKnownSkillPath(path: unknown, roots: Iterable<string>): path is string {
  if (typeof path !== 'string' || path === '' || !isAbsolute(path)) return false
  const file = resolve(path)
  if (basename(file) !== 'SKILL.md') return false
  const root = dirname(dirname(file))
  for (const known of roots) if (resolve(known) === root) return true
  return false
}
