import { app, ipcMain, type IpcMainInvokeEvent } from 'electron'
import { delimiter, dirname, join, resolve } from 'node:path'
import { HomeCliError, inPlaceTarget, validateHomeCliArgs } from './home-cli'
import { readSkillBody, usableSkills } from './imported-skills'
import { createCliRunner, type CliRunner } from './mcp/cli-runner'

/**
 * The AI panel's door to the skills a user installed, and to the CLI those
 * skills teach.
 *
 * genoffice#1895 shipped the first half: a person can import a skill into
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
 * (spawn takes an argv array), the home screen's command guard (document
 * commands only, no `--force`, `apply`/`replace` write a new file), a bounded
 * argv, and bounded output. The child runs in the open document's folder so
 * the relative paths skills teach land inside the allowed roots.
 */

/** tool results travel over IPC and into the transcript; keep both bounded */
const MAX_STD_CHARS = 32_000

export interface UserSkillsDeps {
  /** lazy: built on first use, so registering costs nothing */
  cliRunner: () => CliRunner
  /** the user's folder roots (save folder + added folders) */
  folderRoots: () => readonly string[]
  /** file of the document the calling window renders, when it has one */
  senderFile: (webContentsId: number) => string | undefined
}

/** PATH-style GENOFFICE_ALLOWED_ROOTS: the open document's folder plus the user's roots */
export function allowedRootsFor(
  folderRoots: readonly string[],
  senderFile: string | undefined,
): string {
  const roots = new Set(folderRoots.map((r) => resolve(r)))
  if (senderFile) roots.add(resolve(dirname(senderFile)))
  return [...roots].join(delimiter)
}

/** what the renderer's skills skill reads back from `ai:skills-list` */
export interface UserSkillEntry {
  name: string
  description: string
  relevant: boolean
}

/** validate the argv the model sent; throws with the reason when it is not one */
export function validateCliArgs(args: unknown): string[] {
  // run_cli has no file staging: a long ops list goes through the document tools instead
  const argv = validateHomeCliArgs(args, 'use the document tools for large edits')
  // unlike the home screen, run_cli keeps no backup, so an edit always goes to a new file
  const inPlace = inPlaceTarget(argv)
  if (inPlace) {
    throw new HomeCliError(`${argv[0]} ${argv[1]} must write a new file: add --out <path>`)
  }
  return argv
}

/** where relative paths resolve: the open document's folder, else the save folder */
export function cliCwdFor(
  folderRoots: readonly string[],
  senderFile: string | undefined,
): string | undefined {
  if (senderFile) return resolve(dirname(senderFile))
  return folderRoots[0] ? resolve(folderRoots[0]) : undefined
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

  ipcMain.handle('ai:run-cli', async (e: IpcMainInvokeEvent, args: unknown) => {
    const argv = validateCliArgs(args)
    const roots = deps.folderRoots()
    const file = deps.senderFile(e.sender.id)
    const cwd = cliCwdFor(roots, file)
    const out = await runner().run(argv, {
      timeoutMs: 120_000,
      ...(cwd ? { cwd } : {}),
      env: { GENOFFICE_ALLOWED_ROOTS: allowedRootsFor(roots, file) },
    })
    return {
      ok: out.ok,
      code: out.code,
      json: out.json,
      stdout: clip(out.stdout),
      stderr: clip(out.stderr),
      cwd,
      allowedRoots: allowedRootsFor(roots, file).split(delimiter).filter(Boolean),
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
