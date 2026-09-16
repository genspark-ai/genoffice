import { statSync } from 'node:fs'
import { parseDocx } from '@genoffice/docx-engine'
import { openPptx } from '@genoffice/pptx-engine'
import { flagString } from '../args'
import { csvInfo } from '../formats/csv'
import { pdfInfo } from '../formats/pdf'
import { workbookSummary } from '../formats/xlsx'
import { extension, readInput, resolveInput } from '../fs'
import type { CommandDef } from '../registry'
import { CliError, EXIT } from '../result'

const EMU_PER_INCH = 914400

export const infoCommand: CommandDef = {
  name: 'info',
  summary: 'Print metadata and a structure summary of a document.',
  usage: 'info <file> [--password <pw>]',
  options: [{ name: 'password', value: 'pw', description: 'password for an encrypted PDF' }],
  async run(args, ctx) {
    const path = resolveInput(args.positionals[0], ctx)
    const ext = extension(path)
    const stat = statSync(path)
    const base = { path, format: ext, size_bytes: stat.size, modified: stat.mtime.toISOString() }
    const detail = await describe(path, ext, flagString(args, 'password'))
    return { summary: `${ext} · ${detail.headline}`, detail: { ...base, ...detail.fields } }
  },
}

interface Description {
  headline: string
  fields: Record<string, unknown>
}

async function describe(path: string, ext: string, password?: string): Promise<Description> {
  switch (ext) {
    case 'docx':
      return describeDocx(path)
    case 'pptx':
      return describePptx(path)
    case 'xlsx':
    case 'xlsm':
    case 'xls':
    case 'xlsb':
    case 'ods':
      return describeWorkbook(path)
    case 'pdf':
      return describePdf(path, password)
    case 'csv':
      return describeCsv(path)
    case 'md':
    case 'markdown':
    case 'html':
    case 'htm':
    case 'txt':
      return describeText(path, ext)
    default:
      throw new CliError(EXIT.usage, `unsupported file type: .${ext}`, undefined, {
        reason: 'unsupported',
      })
  }
}

async function describeDocx(path: string): Promise<Description> {
  const parsed = await parseDocx(readInput(path))
  const counts: Record<string, number> = {}
  for (const b of parsed.blocks) counts[b.type] = (counts[b.type] ?? 0) + 1
  const headings = parsed.blocks
    .filter((b) => b.type === 'heading')
    .slice(0, 20)
    .map((b) => ({
      level: b.level ?? 1,
      text: (b.runs ?? []).map((r) => r.text).join(''),
    }))
  return {
    headline: `${parsed.blocks.length} blocks, ${counts.heading ?? 0} headings, ${counts.table ?? 0} tables`,
    fields: {
      blocks: parsed.blocks.length,
      block_types: counts,
      headings,
      comments: parsed.comments.length,
      footnotes: parsed.footnotes.length,
      endnotes: parsed.endnotes.length,
      protected: parsed.protection !== null,
    },
  }
}

async function describePptx(path: string): Promise<Description> {
  const { deck } = await openPptx(readInput(path))
  const elements: Record<string, number> = {}
  for (const slide of deck.slides) {
    for (const el of slide.elements) elements[el.type] = (elements[el.type] ?? 0) + 1
  }
  return {
    headline: `${deck.slides.length} slides`,
    fields: {
      slides: deck.slides.length,
      slide_size_in: {
        width: round(deck.size.cx / EMU_PER_INCH),
        height: round(deck.size.cy / EMU_PER_INCH),
      },
      element_types: elements,
    },
  }
}

async function describeWorkbook(path: string): Promise<Description> {
  const wb = await workbookSummary(path)
  return {
    headline: `${wb.sheets.length} sheets`,
    fields: { sheets: wb.sheets, active_sheet: wb.activeSheet, defined_names: wb.definedNames },
  }
}

async function describePdf(path: string, password?: string): Promise<Description> {
  const info = await pdfInfo(readInput(path), password)
  return {
    headline: info.pages === null ? 'encrypted (password required)' : `${info.pages} pages`,
    fields: {
      pages: info.pages,
      encrypted: info.encrypted,
      ...(info.producer ? { producer: info.producer } : {}),
      ...(info.creator ? { creator: info.creator } : {}),
    },
  }
}

function describeCsv(path: string): Description {
  const info = csvInfo(readInput(path))
  return {
    headline: `${info.rows} rows × ${info.columns} columns`,
    fields: { rows: info.rows, columns: info.columns, delimiter: info.delimiter },
  }
}

function describeText(path: string, ext: string): Description {
  const text = Buffer.from(readInput(path)).toString('utf-8')
  const lines = text.split(/\r?\n/)
  const fields: Record<string, unknown> = { lines: lines.length, characters: text.length }
  if (ext === 'md' || ext === 'markdown') {
    fields.headings = lines.filter((l) => /^#{1,6}\s/.test(l)).length
  }
  return { headline: `${lines.length} lines`, fields }
}

function round(n: number): number {
  return Math.round(n * 100) / 100
}
