import { basename } from 'node:path'
import { PdfLoadError } from '@genoffice/pdf2docx'
import { flagBool, flagString } from '../args'
import { PDF_DEFAULT_PAGES, PDF_PAGE_PREVIEW_CHARS, pdfText } from '../formats/pdf'
import { extension, readInput, resolveInput } from '../fs'
import { previewChars } from '../preview'
import type { CommandContext, CommandDef } from '../registry'
import { CliError, EXIT, type CommandResult } from '../result'

export const pdfCommand: CommandDef = {
  name: 'pdf',
  summary: 'Read the text layer of a PDF page by page, without any app process.',
  usage: 'pdf read <file.pdf> [--page n | --range a-b] [--password <pw>] [--full | --max-chars n]',
  options: [
    { name: 'page', value: 'n', description: 'read: only this 1-based page' },
    {
      name: 'range',
      value: 'a-b',
      description: `read: 1-based page range (default: the first ${PDF_DEFAULT_PAGES} pages)`,
    },
    { name: 'password', value: 'pw', description: 'password for an encrypted PDF' },
    {
      name: 'full',
      description: 'read: every page with its whole text instead of a clipped preview',
    },
    {
      name: 'max-chars',
      value: 'n',
      description: `read: preview length per page (default ${PDF_PAGE_PREVIEW_CHARS}); clipped text ends in …(+n chars)`,
    },
  ],
  async run(args, ctx) {
    const [verb, file] = args.positionals
    if (verb !== 'read') {
      throw new CliError(EXIT.usage, 'expected "pdf read <file.pdf>"', undefined, {
        reason: verb === undefined ? 'missing_argument' : 'invalid_argument',
        suggestion: 'run `genoffice help pdf`',
      })
    }
    return read(file, args, ctx)
  },
}

type Args = Parameters<CommandDef['run']>[0]

async function read(
  file: string | undefined,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  if (extension(path) !== 'pdf') {
    throw new CliError(EXIT.usage, `not a PDF: ${path}`, undefined, {
      reason: 'unsupported',
      suggestion: 'use docs read, sheet read or slides read for Office files',
    })
  }
  const page = flagString(args, 'page')
  const range = flagString(args, 'range')
  if (page !== undefined && range !== undefined) {
    throw new CliError(EXIT.usage, '--page and --range are exclusive', undefined, {
      reason: 'invalid_argument',
    })
  }
  const full = flagBool(args, 'full')
  const password = flagString(args, 'password')
  let windowed = false
  const select = (count: number): [number, number] => {
    if (page !== undefined) return parseSpan(page, count, '--page must be a 1-based page number')
    if (range !== undefined) return parseSpan(range, count, '--range must look like 3 or 3-10')
    if (full) return [1, count]
    windowed = count > PDF_DEFAULT_PAGES
    return [1, Math.min(PDF_DEFAULT_PAGES, count)]
  }
  let doc
  try {
    doc = await pdfText(readInput(path), {
      password,
      maxChars: previewChars(args, PDF_PAGE_PREVIEW_CHARS),
      select,
    })
  } catch (err) {
    throw loadError(err, password)
  }
  const first = doc.pages_read[0]?.page
  const last = doc.pages_read.at(-1)?.page
  const detail: Record<string, unknown> = { ...doc }
  if (first !== undefined && last !== undefined) detail.range = `${first}-${last}`
  if (windowed && last !== undefined) {
    const next = `${last + 1}-${Math.min(last + PDF_DEFAULT_PAGES, doc.pages)}`
    detail.truncated = `pages ${last + 1}-${doc.pages} not read; pass --range ${next} for the next ${PDF_DEFAULT_PAGES}, --page n for one, or --full for every page`
  }
  const read = detail.range ? `, read ${detail.range}` : ''
  return { summary: `${basename(path)}: ${doc.pages} pages${read}`, detail }
}

function parseSpan(spec: string, count: number, usage: string): [number, number] {
  const m = /^(\d+)(?:-(\d+))?$/.exec(spec)
  if (!m) throw new CliError(EXIT.usage, usage, undefined, { reason: 'invalid_argument' })
  const start = Number(m[1])
  const end = m[2] === undefined ? start : Number(m[2])
  if (start < 1 || start > end || end > count) {
    throw new CliError(
      EXIT.usage,
      `page span ${spec} out of bounds (1-${count})`,
      { valid_range: [1, count] },
      { reason: 'out_of_range', suggestion: `use pages between 1 and ${count}` },
    )
  }
  return [start, end]
}

function loadError(err: unknown, password: string | undefined): unknown {
  if (!(err instanceof PdfLoadError)) return err
  if (err.code === 'password-required') {
    return password === undefined
      ? new CliError(EXIT.usage, 'the PDF is encrypted', undefined, {
          reason: 'missing_argument',
          suggestion: 'pass --password <pw>',
        })
      : new CliError(EXIT.usage, 'wrong password for this PDF', undefined, {
          reason: 'invalid_argument',
          suggestion: 'check the --password value',
        })
  }
  return new CliError(EXIT.file, `not a readable PDF (${err.code})`, undefined, {
    reason: 'unsupported',
  })
}
