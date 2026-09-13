import { basename } from 'node:path'
import { flagString } from '../args'
import { resolveInput } from '../fs'
import { outputDirectory, parseScale, RENDERABLE, renderToPngs } from '../formats/render'
import type { CommandDef } from '../registry'
import { CliError, EXIT } from '../result'

export const renderCommand: CommandDef = {
  name: 'render',
  summary:
    'One PNG per page of a document, as the GenOffice renderer lays it out: the picture an agent looks at to check a Word document, a workbook or a page it just made.',
  usage: 'render <file> --out <dir> [--page n] [--scale n]',
  options: [
    { name: 'out', value: 'dir', description: 'directory for the PNGs (<stem>-NN.png; required)' },
    { name: 'page', value: 'n', description: 'only this 1-based page (default: every page)' },
    {
      name: 'scale',
      value: 'n',
      description: 'pixels per PDF point, 0 < n <= 4 (default 1 = 72 dpi; 2 = 144 dpi)',
    },
  ],
  async run(args, ctx) {
    const path = resolveInput(args.positionals[0], ctx)
    const outDir = outputDirectory(flagString(args, 'out'), ctx)
    const pageRaw = flagString(args, 'page')
    let only: number | undefined
    if (pageRaw !== undefined) {
      const n = Number(pageRaw)
      if (!Number.isInteger(n) || n < 1) throw new CliError(EXIT.usage, '--page must be 1 or more')
      only = n - 1
    }
    const files = await renderToPngs(path, ctx, {
      outDir,
      scale: parseScale(flagString(args, 'scale')),
      only,
      range: { flag: 'page', oneBased: true },
      log: ctx.log,
    })
    return {
      summary: `rendered ${files.length} page(s) of ${basename(path)} to ${outDir}`,
      outputPath: outDir,
      detail: {
        files: files.map((f) => ({ ...f, page: f.page + 1 })),
        formats: RENDERABLE,
        via: path.toLowerCase().endsWith('.pdf')
          ? 'pdfium'
          : 'genoffice --headless-export + pdfium',
      },
    }
  },
}
