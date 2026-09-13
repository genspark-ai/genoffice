import { randomUUID } from 'node:crypto'
import { mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, extname, isAbsolute, join, resolve } from 'node:path'
import { assertAllowed, readInput, type PathContext } from '../fs'
import { CliError, EXIT } from '../result'
import { exportViaApp } from './app-export'
import { rasterizePdf } from './slide-spec'

/** Documents the app can print to PDF, plus PDF itself (rasterized directly). */
export const RENDERABLE = [
  'pdf',
  'docx',
  'xlsx',
  'xlsm',
  'csv',
  'pptx',
  'md',
  'markdown',
  'html',
  'htm',
]

export interface RenderOptions {
  outDir: string
  scale: number
  /** 0-based page to render; all pages when absent */
  only?: number
  /** how the caller named the page flag, for the out-of-range message */
  range?: { flag: string; oneBased: boolean }
  log: (message: string) => void
}

export interface RenderedFile {
  page: number
  path: string
  width: number
  height: number
}

export function parseScale(raw: string | undefined): number {
  const scale = raw === undefined ? 1 : Number(raw)
  if (!(scale > 0 && scale <= 4)) throw new CliError(EXIT.usage, '--scale must be between 0 and 4')
  return scale
}

export function outputDirectory(spec: string | undefined, ctx: PathContext): string {
  if (!spec) throw new CliError(EXIT.usage, 'missing --out <directory>')
  const dir = isAbsolute(spec) ? spec : resolve(ctx.cwd, spec)
  return assertAllowed(dir, ctx.env, 'write')
}

/**
 * One PNG per page: the document is printed to a temporary PDF by the hidden
 * GenOffice process (a PDF input skips that step) and rasterized with pdfium.
 * Files are `<stem>-NN.png`, NN 1-based, in `outDir`.
 */
export async function renderToPngs(
  path: string,
  ctx: PathContext,
  opts: RenderOptions,
): Promise<RenderedFile[]> {
  const ext = extname(path).slice(1).toLowerCase()
  if (!RENDERABLE.includes(ext)) {
    throw new CliError(EXIT.usage, `cannot render .${ext || '?'}`, { supported: RENDERABLE })
  }
  const tmpPdf = ext === 'pdf' ? null : join(tmpdir(), `genoffice-render-${randomUUID()}.pdf`)
  try {
    if (tmpPdf) await exportViaApp(path, 'pdf', tmpPdf, { env: ctx.env, log: opts.log })
    const pages = await rasterizePdf(readInput(tmpPdf ?? path), opts.scale, opts.only, opts.range)
    mkdirSync(opts.outDir, { recursive: true })
    const stem = basename(path, extname(path))
    return pages.map((p) => {
      const out = join(opts.outDir, `${stem}-${String(p.index + 1).padStart(2, '0')}.png`)
      writeFileSync(out, p.png)
      return { page: p.index, path: out, width: p.width, height: p.height }
    })
  } finally {
    if (tmpPdf) rmSync(tmpPdf, { force: true })
  }
}
