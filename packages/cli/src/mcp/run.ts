import { randomBytes } from 'node:crypto'
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { delimiter, join } from 'node:path'
import { runCli } from '../cli'
import type { JsonError, JsonOk } from '../result'
import type { InlineFile } from './tools'

export interface McpContext {
  cwd: string
  env: NodeJS.ProcessEnv
  /** diagnostics; the MCP transport owns stdout, so these go to stderr */
  log: (message: string) => void
  /** where inline parameters land as files for the duration of one call */
  scratchDir: string
}

export type Outcome = { ok: JsonOk; error?: undefined } | { ok?: undefined; error: JsonError }

export function createContext(base: {
  cwd: string
  env: NodeJS.ProcessEnv
  log: (m: string) => void
}): McpContext {
  const scratchDir = join(
    tmpdir(),
    `genoffice-mcp-${process.pid}-${randomBytes(4).toString('hex')}`,
  )
  mkdirSync(scratchDir, { recursive: true })
  return { ...base, env: allowScratch(base.env, scratchDir), scratchDir }
}

export function disposeContext(ctx: McpContext): void {
  rmSync(ctx.scratchDir, { recursive: true, force: true })
}

/** GENOFFICE_ALLOWED_ROOTS stays in force for the caller's paths; the server's own scratch dir is added so inline ops can be read. */
function allowScratch(env: NodeJS.ProcessEnv, scratchDir: string): NodeJS.ProcessEnv {
  const roots = env.GENOFFICE_ALLOWED_ROOTS
  if (roots === undefined || roots.trim() === '') return env
  return { ...env, GENOFFICE_ALLOWED_ROOTS: `${roots}${delimiter}${scratchDir}` }
}

/** One command, as `genoffice <argv> --json` would run it, with the JSON envelope parsed back. */
export async function runJson(argv: string[], ctx: McpContext): Promise<Outcome> {
  const out: string[] = []
  await runCli([...argv, '--json'], {
    cwd: ctx.cwd,
    env: ctx.env,
    io: { stdout: (t) => out.push(t), stderr: (t) => ctx.log(t) },
  })
  const text = out.join('\n')
  let parsed: JsonOk | JsonError
  try {
    parsed = JSON.parse(text) as JsonOk | JsonError
  } catch {
    return {
      error: {
        status: 'error',
        command: argv[0] ?? null,
        code: 3,
        error: 'conversion_failed',
        message: `genoffice printed no JSON: ${text.slice(0, 200)}`,
      },
    }
  }
  return parsed.status === 'error' ? { error: parsed } : { ok: parsed }
}

/** Runs `argv` with every inline file written to the scratch dir and passed as `--<option> <path>`; the files are removed afterwards. */
export async function runWithInline(
  argv: string[],
  inline: InlineFile[],
  ctx: McpContext,
): Promise<Outcome> {
  const paths: string[] = []
  const full = [...argv]
  for (const f of inline) {
    const path = join(ctx.scratchDir, `${f.option}-${randomBytes(4).toString('hex')}${f.ext}`)
    writeFileSync(path, f.text)
    paths.push(path)
    full.push(`--${f.option}`, path)
  }
  try {
    return await runJson(full, ctx)
  } finally {
    for (const p of paths) rmSync(p, { force: true })
  }
}

export interface ImageBlock {
  type: 'image'
  data: string
  mimeType: 'image/png'
}

const MAX_IMAGES = 12
const MAX_IMAGE_BYTES = 6 * 1024 * 1024

/** The rendered PNGs a result lists (detail.files[].path), as image content within a size budget. */
export function imageBlocks(ok: JsonOk): { images: ImageBlock[]; omitted: number } {
  const files = (ok.detail?.files as { path?: string }[] | undefined) ?? []
  const images: ImageBlock[] = []
  let bytes = 0
  let omitted = 0
  for (const f of files) {
    if (!f.path) continue
    if (images.length >= MAX_IMAGES) {
      omitted++
      continue
    }
    const buf = readFileSync(f.path)
    if (bytes + buf.byteLength > MAX_IMAGE_BYTES) {
      omitted++
      continue
    }
    bytes += buf.byteLength
    images.push({ type: 'image', data: buf.toString('base64'), mimeType: 'image/png' })
  }
  return { images, omitted }
}
