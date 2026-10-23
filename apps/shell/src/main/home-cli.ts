import { statSync } from 'node:fs'
import { copyFile, mkdir, writeFile } from 'node:fs/promises'
import {
  basename,
  delimiter,
  dirname,
  isAbsolute,
  join,
  normalize,
  relative,
  resolve,
  sep,
} from 'node:path'
import { parseArgs, type ParsedArgs } from '@genoffice/cli/args'
import type { CliRunRequest, CliRunResult, CliStagedFile } from '@genoffice/agent-core'
import type { CliRunner } from './mcp/cli-runner'

/**
 * The home screen's Ask AI running the bundled `genoffice` CLI. The renderer
 * sends a command line and the files it wants staged; this module decides what
 * may run and where it may touch the disk:
 *
 * - only document commands — no `open`, `mcp`, `skill`, `install-cli`;
 * - no `--force`; `apply` / `replace` without `--out` (an in-place edit) only
 *   for a file inside the output directory, and a copy of the original goes
 *   to the work directory first — a starred source elsewhere is never
 *   rewritten;
 * - reads and writes confined by GENOFFICE_ALLOWED_ROOTS to the conversation's
 *   work directory, the save folder, and the folders of the starred files and
 *   of the files attached to the conversation;
 * - staged files land inside the work directory only (relative, no `..`).
 */

export const HOME_CLI_COMMANDS: readonly string[] = [
  'info',
  'convert',
  'create',
  'slides',
  'sheet',
  'docs',
  'merge',
  'pdf',
  'render',
  'guide',
  'search',
  'image',
  'media',
  'capabilities',
]

const MAX_ARGS = 32
const MAX_ARG_CHARS = 1000
const MAX_FILES = 64
const MAX_FILE_CHARS = 4_000_000
const MAX_STD_CHARS = 64_000
const SESSION_ID = /^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$/

export class HomeCliError extends Error {}

export function validateHomeCliArgs(raw: unknown, longArgHint = 'stage it as a file'): string[] {
  if (!Array.isArray(raw) || raw.length === 0) {
    throw new HomeCliError('args must be a non-empty array of strings')
  }
  if (raw.length > MAX_ARGS) throw new HomeCliError(`at most ${MAX_ARGS} arguments`)
  const args: string[] = []
  for (const item of raw) {
    if (typeof item !== 'string' || item.trim() === '') {
      throw new HomeCliError('every argument must be a non-empty string')
    }
    if (item.length > MAX_ARG_CHARS) {
      throw new HomeCliError(
        `an argument is longer than ${MAX_ARG_CHARS} characters; ${longArgHint}`,
      )
    }
    // the runner appends --json itself; a model that adds one is not wrong, just redundant
    if (item === '--json') continue
    args.push(item)
  }
  const command = args[0]
  if (command === undefined) throw new HomeCliError('args must name a command')
  if (command !== '--help' && !HOME_CLI_COMMANDS.includes(command)) {
    throw new HomeCliError(
      `"${command}" is not available here; commands: ${HOME_CLI_COMMANDS.join(', ')}`,
    )
  }
  if (args.includes('--force')) {
    throw new HomeCliError('--force is not allowed: choose a file name that does not exist yet')
  }
  return args
}

/** apply / replace flags that take no value (the CLI list is registry-wide; these are the ones a model sends here) */
const BOOLEAN_FLAGS = new Set([
  'dry-run',
  'track',
  'best-effort',
  'stop-on-error',
  'help',
  'force',
  'json',
])

/**
 * The file an `apply` / `replace` rewrites when no --out is given, or null.
 * Parsed with the CLI's own parseArgs so `--out=dest`, `--ops o.json a.docx`
 * and a flag followed by another flag all read the way the CLI will read them.
 * An apply with no positional at all is refused rather than guessed.
 */
export function inPlaceTarget(args: readonly string[]): string | null {
  const sub = args[1]
  if (sub !== 'apply' && sub !== 'replace') return null
  let parsed: ParsedArgs
  try {
    parsed = parseArgs(args.slice(2), BOOLEAN_FLAGS)
  } catch (err) {
    throw new HomeCliError(err instanceof Error ? err.message : String(err))
  }
  if (parsed.flags.out !== undefined || parsed.flags.help !== undefined) return null
  const target = parsed.positionals[0]
  if (target === undefined) {
    throw new HomeCliError(`${args[0]} ${sub} needs the file to edit as an argument`)
  }
  return target
}

function isInside(root: string, path: string): boolean {
  const rel = relative(resolve(root), resolve(path))
  return rel === '' || (!rel.startsWith('..') && !isAbsolute(rel))
}

export function validateSessionId(raw: unknown): string {
  if (typeof raw !== 'string' || !SESSION_ID.test(raw)) throw new HomeCliError('bad session id')
  return raw
}

/** a staged name stays inside the work directory: relative, normalized, no `..` segment */
export function stagedPath(workDir: string, name: string): string {
  if (isAbsolute(name) || /^[A-Za-z]:/.test(name)) {
    throw new HomeCliError(`staged file names are relative to the work directory: ${name}`)
  }
  const rel = normalize(name.replace(/\\/g, '/'))
  if (rel === '.' || rel.split(sep).includes('..')) {
    throw new HomeCliError(`staged file name escapes the work directory: ${name}`)
  }
  return join(workDir, rel)
}

export function validateStagedFiles(raw: unknown): CliStagedFile[] {
  if (raw === undefined || raw === null) return []
  if (!Array.isArray(raw)) throw new HomeCliError('files must be an array')
  if (raw.length > MAX_FILES) throw new HomeCliError(`at most ${MAX_FILES} files per call`)
  const out: CliStagedFile[] = []
  for (const item of raw) {
    const { name, content } = (item && typeof item === 'object' ? item : {}) as Record<
      string,
      unknown
    >
    if (typeof name !== 'string' || !name.trim() || typeof content !== 'string') {
      throw new HomeCliError('every file needs a string name and string content')
    }
    if (content.length > MAX_FILE_CHARS) throw new HomeCliError(`${name} is too large to stage`)
    out.push({ name, content })
  }
  return out
}

export interface HomeCliDeps {
  runner: () => CliRunner
  /** scratch root for all conversations (under userData, never a document folder) */
  workRoot: () => string
  /** finished documents go here */
  outDir: () => string
  /** the starred files: their folders become readable roots */
  sourcePaths: () => readonly string[]
  recordRecent: (path: string) => void
}

export interface HomeCliRunRequest extends CliRunRequest {
  /** files the user attached to the conversation; their folders become readable roots */
  sources?: string[]
}

/** absolute paths of existing regular files; anything else is dropped silently */
export function validateSources(raw: unknown): string[] {
  if (!Array.isArray(raw)) return []
  const out: string[] = []
  for (const item of raw) {
    if (typeof item !== 'string' || !isAbsolute(item)) continue
    try {
      if (statSync(item).isFile()) out.push(item)
    } catch {
      // vanished since it was attached: no root for it
    }
  }
  return out
}

export interface HomeCliContext {
  workDir: string
  outDir: string
}

export function homeCliContext(deps: HomeCliDeps, sessionId: string): HomeCliContext {
  return { workDir: join(deps.workRoot(), validateSessionId(sessionId)), outDir: deps.outDir() }
}

/** PATH-style GENOFFICE_ALLOWED_ROOTS for one run */
export function allowedRootsFor(ctx: HomeCliContext, sourcePaths: readonly string[]): string {
  const roots = new Set<string>([resolve(ctx.workDir), resolve(ctx.outDir)])
  for (const path of sourcePaths) roots.add(resolve(dirname(path)))
  return [...roots].join(delimiter)
}

function clip(text: string): string {
  return text.length > MAX_STD_CHARS
    ? `${text.slice(0, MAX_STD_CHARS)}\n…(+${text.length - MAX_STD_CHARS} chars)`
    : text
}

export async function runHomeCli(
  deps: HomeCliDeps,
  sessionId: string,
  request: HomeCliRunRequest,
): Promise<CliRunResult> {
  const ctx = homeCliContext(deps, sessionId)
  const args = validateHomeCliArgs(request.args)
  const files = validateStagedFiles(request.files)
  await mkdir(ctx.workDir, { recursive: true })
  for (const file of files) {
    const path = stagedPath(ctx.workDir, file.name)
    await mkdir(dirname(path), { recursive: true })
    await writeFile(path, file.content, 'utf8')
  }
  const inPlace = inPlaceTarget(args)
  if (inPlace) {
    const target = resolve(ctx.workDir, inPlace)
    if (!isInside(ctx.outDir, target)) {
      throw new HomeCliError(
        `${args[0]} ${args[1]} rewrites ${inPlace} in place; that is only allowed inside the output directory — add --out <path under the output directory>`,
      )
    }
    const backupDir = join(ctx.workDir, '.backup')
    await mkdir(backupDir, { recursive: true })
    await copyFile(target, join(backupDir, `${Date.now()}-${basename(target)}`)).catch(() => {
      throw new HomeCliError(`${inPlace} does not exist`)
    })
  }
  const outcome = await deps.runner().run(args, {
    cwd: ctx.workDir,
    env: {
      GENOFFICE_ALLOWED_ROOTS: allowedRootsFor(ctx, [
        ...deps.sourcePaths(),
        ...validateSources(request.sources),
      ]),
    },
    timeoutMs: 300_000,
  })
  const out = outcome.json?.status === 'ok' ? outcome.json.output_path : undefined
  if (out && resolve(dirname(out)) === resolve(ctx.outDir)) deps.recordRecent(out)
  return {
    ok: outcome.ok,
    code: outcome.code,
    json: outcome.json,
    stdout: clip(outcome.stdout),
    stderr: clip(outcome.stderr),
  }
}
