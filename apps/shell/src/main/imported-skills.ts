import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { homedir } from 'node:os'
import { classifySkill, parseSkillFrontmatter, type SkillRelevance } from '@genoffice/agent-core'

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
export interface FoundSkill {
  /** the directory name, which is what `skills/<name>/` is keyed on */
  name: string
  description: string
  /** absolute path of the SKILL.md, read when the user picks it */
  path: string
  /** where it came from: our own folder, or an agent's directory */
  source: 'genoffice' | 'agent'
  /** the agent whose directory it came from, when source is 'agent' */
  agent?: string
  /** whether it talks about a format GenOffice opens, and on what */
  relevance: SkillRelevance
}

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
