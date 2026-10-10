import { app, ipcMain } from 'electron'
import { join } from 'node:path'
import { readSkillBody, usableSkills } from './imported-skills'
import { createCliRunner, type CliRunner } from './mcp/cli-runner'

/**
 * The AI panel's door to the skills a user installed, and to the CLI those
 * skills teach.
 *
 * #1895 shipped the first half: a person can import a skill into
 * `<userData>/skills` and read it in Integrations. This is the second half —
 * what the model gets, and what it may execute:
 *
 * - `ai:skills-list` and `ai:skill-text` hand out only what is installed in
 *   *our* folder. The coding agents' directories are candidates the user has
 *   not chosen, and until they click Import the model does not see them.
 *   `ai:skill-text` takes a name, never a path: the model cannot name a file,
 *   it can only pick an entry from the list this process vouched for.
 * - `ai:run-cli` spawns the bundled `genoffice` CLI and nothing else — the
 *   executable is resolved here, so an argv is a command line for our own
 *   document tool, not a general shell. `--json` is appended by the runner,
 *   which is also what keeps the result a single parseable object instead of
 *   prose the model would have to guess at.
 *
 * There is no approval flow for tool calls (the loop has none), so the
 * boundary is exactly this list: no arbitrary binaries, no shell interpolation
 * (spawn takes an argv array), a bounded argv, and bounded output.
 */

/** argv shape for `ai:run-cli`: a command and its arguments, nothing else */
const MAX_ARGS = 8
const MAX_ARG_CHARS = 200
const MAX_ARGV_CHARS = 2000

/** tool results travel over IPC and into the transcript; keep both bounded */
const MAX_STD_CHARS = 32_000

export interface UserSkillsDeps {
  /** lazy: built on first use, so registering costs nothing */
  cliRunner: () => CliRunner
}

/** what the renderer's skills skill reads back from `ai:skills-list` */
export interface UserSkillEntry {
  name: string
  description: string
  relevant: boolean
}

/** validate the argv the model sent; throws with the reason when it is not one */
export function validateCliArgs(args: unknown): string[] {
  if (!Array.isArray(args) || args.length === 0) {
    throw new Error('args must be a non-empty array of strings')
  }
  if (args.length > MAX_ARGS) throw new Error(`at most ${MAX_ARGS} arguments`)
  let total = 0
  const out: string[] = []
  for (const raw of args) {
    if (typeof raw !== 'string' || raw.trim() === '') {
      throw new Error('every argument must be a non-empty string')
    }
    if (raw.length > MAX_ARG_CHARS) {
      throw new Error(`an argument is longer than ${MAX_ARG_CHARS} characters`)
    }
    total += raw.length
    if (total > MAX_ARGV_CHARS) throw new Error('the command line is too long')
    out.push(raw)
  }
  // the first token selects a CLI subcommand; a flag first is a mistake at
  // best and a way to probe global flags at worst
  if (out[0].startsWith('-')) throw new Error('the first argument must be a command, not a flag')
  return out
}

function clip(text: string): string {
  return text.length > MAX_STD_CHARS
    ? `${text.slice(0, MAX_STD_CHARS)}…(+${text.length - MAX_STD_CHARS} chars)`
    : text
}

export function registerUserSkillsIpc(deps: UserSkillsDeps): void {
  const runner = (): CliRunner => deps.cliRunner()

  ipcMain.handle('ai:skills-list', async (): Promise<UserSkillEntry[]> => {
    const skills = await usableSkills(app.getPath('userData'))
    return skills.map(({ name, description, relevance }) => ({
      name,
      description,
      relevant: relevance.relevant,
    }))
  })

  ipcMain.handle('ai:skill-text', async (_e, name: unknown): Promise<string> => {
    if (typeof name !== 'string' || name.trim() === '') {
      throw new Error('a skill name is required')
    }
    const skills = await usableSkills(app.getPath('userData'))
    const skill = skills.find((s) => s.name === name)
    if (!skill) throw new Error(`no installed skill named "${name}"`)
    return clip(readSkillBody(skill.path))
  })

  ipcMain.handle('ai:run-cli', async (_e, args: unknown) => {
    const argv = validateCliArgs(args)
    const out = await runner().run(argv, { timeoutMs: 120_000 })
    return {
      ok: out.ok,
      code: out.code,
      json: out.json,
      stdout: clip(out.stdout),
      stderr: clip(out.stderr),
    }
  })
}

/** the CLI entry the app ships, dev or packaged — the same one MCP delegates to */
export function bundledCliEntry(): string {
  return app.isPackaged
    ? join(process.resourcesPath, 'cli', 'genoffice.cjs')
    : join(app.getAppPath(), '..', 'packages', 'cli', 'dist', 'genoffice.cjs')
}

export function makeUserCliRunner(): CliRunner {
  return createCliRunner({ executable: process.execPath, entry: bundledCliEntry() })
}
