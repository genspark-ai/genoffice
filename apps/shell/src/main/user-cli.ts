import { app, ipcMain, type IpcMainInvokeEvent } from 'electron'
import { delimiter, dirname, join, resolve } from 'node:path'
import { HomeCliError, inPlaceTarget, validateHomeCliArgs } from './home-cli'
import {
  importForeignSkill,
  readSkillBody,
  scanForeignSkills,
  skillRawText,
  usableSkills,
  writeSkill,
} from './imported-skills'
import { pluginSkills, readPluginSkillBody } from './imported-plugins'
import { createCliRunner, type CliRunner } from './mcp/cli-runner'
import type { ForeignSkillCandidate, FoundSkill } from '../shared/found-skill'

/**
 * The AI panel's door to the skills a user installed, and to the CLI those
 * skills teach.
 *
 * genoffice#1895 shipped the first half: a person can import a skill into
 * `<userData>/skills` and read it in Integrations. This is the second half —
 * what the model gets, and what it may execute:
 *
 * - `ai:skills-list` and `ai:skill-text` hand out what is installed in *our*
 *   folder, plus the prompt-level components of the enabled plugins — listed
 *   and invoked as `<plugin>:<name>`, read-only (the plugin's own files are
 *   never written; editing happens on a user-owned copy the foreign-import
 *   flow makes). On a name collision the user's skill wins and the plugin's is
 *   dropped, so a plugin cannot shadow something the person wrote themselves.
 *   Both channels take a name, never a path: the model cannot name a file, it
 *   can only pick an entry from the list this process vouched for.
 * - `ai:skill-raw-text` and `ai:skill-save` are the settings editor's door on
 *   the same folder: the raw SKILL.md of a listed skill in, a written one out.
 *   Nothing new becomes runnable — a saved skill is still just instructions
 *   the next turn may read.
 * - `ai:foreign-skills` and `ai:foreign-import` offer the prompt files the
 *   other agent products keep on this machine as import candidates. The scan
 *   runs here — never in the renderer — and the import runs it again before
 *   writing, so a sourcePath the renderer sends back is an echo of a path this
 *   process vouched for moments ago, never authority of its own. A candidate
 *   is not an installed skill: until it is imported, the model does not see it.
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

/**
 * The listing `ai:skills-list` and `ai:skill-text` resolve against: the user's
 * own skills first, then every enabled plugin's. On a name collision the USER
 * skill wins and the plugin one is dropped — a plugin cannot shadow something
 * the person wrote themselves, and the name the model saw keeps resolving to
 * the user's entry.
 */
async function mergedSkills(userDataDir: string): Promise<FoundSkill[]> {
  const user = await usableSkills(userDataDir)
  const taken = new Set(user.map((s) => s.name))
  const plugins = await pluginSkills(userDataDir)
  return [...user, ...plugins.filter((s) => !taken.has(s.name))]
}

export function registerUserSkillsIpc(deps: UserSkillsDeps): void {
  const runner = (): CliRunner => deps.cliRunner()

  ipcMain.handle('ai:skills-list', async (): Promise<UserSkillEntry[]> => {
    const skills = await mergedSkills(app.getPath('userData'))
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
    // the merged listing is the guard: a `<plugin>:<name>` entry resolves only
    // by the name this process listed, and its file is read read-only — the
    // settings editor's write doors below take user skills only
    const skill = (await mergedSkills(app.getPath('userData'))).find((s) => s.name === name)
    if (!skill) throw new Error(`no installed skill named "${name}"`)
    return clip(skill.plugin ? readPluginSkillBody(skill.path) : readSkillBody(skill.path))
  })

  // the settings editor's pair: raw in (frontmatter included), fields out.
  // Same boundary as ai:skill-text — `name` only ever names a listed skill,
  // and the overwrite flag is what separates save from clobber.
  ipcMain.handle('ai:skill-raw-text', async (_e, name: unknown): Promise<string> => {
    if (typeof name !== 'string') throw new Error('unknown skill')
    return skillRawText(app.getPath('userData'), name)
  })

  ipcMain.handle('ai:skill-save', (_e, request: unknown): FoundSkill => {
    const r = (request ?? {}) as {
      name?: unknown
      description?: unknown
      body?: unknown
      allowOverwrite?: unknown
    }
    if (
      typeof r.name !== 'string' ||
      typeof r.description !== 'string' ||
      typeof r.body !== 'string'
    ) {
      throw new Error('a name, a description and a body are required')
    }
    return writeSkill(
      app.getPath('userData'),
      r.name,
      { description: r.description, body: r.body },
      { allowOverwrite: r.allowOverwrite === true },
    )
  })

  ipcMain.handle('ai:foreign-skills', async (): Promise<ForeignSkillCandidate[]> =>
    scanForeignSkills(app.getPath('home'), app.getPath('userData')),
  )

  // the sourcePath came from the renderer, so the import re-runs the scan and
  // only writes a path that fresh scan vouched for — the rule every path-taking
  // handler here follows, applied to a write
  ipcMain.handle('ai:foreign-import', async (_e, sourcePath: unknown): Promise<FoundSkill> => {
    if (typeof sourcePath !== 'string' || sourcePath.trim() === '') {
      throw new Error('a source path is required')
    }
    return importForeignSkill(sourcePath, app.getPath('home'), app.getPath('userData'))
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
