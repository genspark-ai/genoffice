import { existsSync, mkdirSync, readFileSync, realpathSync, statSync } from 'node:fs'
import {
  basename,
  delimiter,
  dirname,
  extname,
  isAbsolute,
  join,
  relative,
  resolve,
  sep,
} from 'node:path'
import { assertNotOpenInGui } from './gui'
import { CliError, EXIT } from './result'

/** What path resolution needs from the command context. */
export interface PathContext {
  cwd: string
  env: NodeJS.ProcessEnv
}

/**
 * GENOFFICE_ALLOWED_ROOTS: PATH-style list of directories genoffice may read from and
 * write to. Unset means unrestricted. Roots are compared after resolving
 * symlinks so a link inside a root cannot point out of it.
 */
export function allowedRoots(env: NodeJS.ProcessEnv): string[] | null {
  const raw = env.GENOFFICE_ALLOWED_ROOTS
  if (raw === undefined || raw.trim() === '') return null
  return raw
    .split(delimiter)
    .map((p) => p.trim())
    .filter(Boolean)
    .map((p) => realizedPath(resolve(p)))
}

/** Real location of a path that may not exist yet: the nearest existing ancestor's realpath plus the rest. */
export function realizedPath(abs: string): string {
  const tail: string[] = []
  let dir = abs
  while (!existsSync(dir)) {
    tail.unshift(basename(dir))
    const parent = dirname(dir)
    if (parent === dir) break
    dir = parent
  }
  const real = existsSync(dir) ? realpathSync(dir) : dir
  return tail.length ? join(real, ...tail) : real
}

function isInside(root: string, path: string): boolean {
  const rel = relative(root, path)
  // a name that merely starts with ".." (e.g. "..hidden") is still inside
  return rel === '' || (rel !== '..' && !rel.startsWith(`..${sep}`) && !isAbsolute(rel))
}

export function assertAllowed(
  abs: string,
  env: NodeJS.ProcessEnv,
  purpose: 'read' | 'write',
): string {
  const roots = allowedRoots(env)
  if (!roots) return abs
  const real = realizedPath(abs)
  if (!roots.some((root) => isInside(root, real))) {
    throw new CliError(
      EXIT.file,
      `refusing to ${purpose} outside GENOFFICE_ALLOWED_ROOTS: ${abs}`,
      {
        allowed_roots: roots,
      },
    )
  }
  return abs
}

export function resolveInput(path: string | undefined, ctx: PathContext): string {
  if (!path) throw new CliError(EXIT.usage, 'missing <file> argument')
  const abs = isAbsolute(path) ? path : resolve(ctx.cwd, path)
  assertAllowed(abs, ctx.env, 'read')
  if (!existsSync(abs)) throw new CliError(EXIT.file, `file not found: ${abs}`)
  if (!statSync(abs).isFile()) throw new CliError(EXIT.file, `not a file: ${abs}`)
  return abs
}

export interface OutputOptions {
  /** used when `spec` is absent (edit in place, derived name) */
  fallback?: string
  /** the user passed --force: overwrite an existing output, write a GUI-open file */
  force?: boolean
  /** refuse an existing output (create/convert) unless forced */
  fresh?: boolean
}

/**
 * The single place an output path is resolved, policy-checked and its
 * directory created. A file the running GenOffice shell has open is refused
 * without --force, so a CLI edit cannot race the editor's own save.
 */
export function resolveOutput(
  spec: string | undefined,
  ctx: PathContext,
  opts: OutputOptions = {},
): string {
  const abs = spec ? (isAbsolute(spec) ? spec : resolve(ctx.cwd, spec)) : opts.fallback
  if (!abs) throw new CliError(EXIT.usage, 'missing --out <path>')
  assertAllowed(abs, ctx.env, 'write')
  if (opts.fresh && !opts.force && existsSync(abs)) {
    throw new CliError(EXIT.file, `output exists: ${abs} (use --force to overwrite)`)
  }
  if (!opts.force) assertNotOpenInGui(abs, ctx.env)
  mkdirSync(dirname(abs), { recursive: true })
  return abs
}

export function readInput(abs: string): Uint8Array {
  const buf = readFileSync(abs)
  return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength)
}

export function extension(path: string): string {
  return extname(path).slice(1).toLowerCase()
}
