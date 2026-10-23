import JSZip from 'jszip'
import { assertZipInflatesWithinLimits, assertZipWithinLimits } from '@genoffice/docx-engine'

export interface DocxTextOptions {
  /** stop once this many characters have been collected (the caller slices anyway) */
  maxChars?: number
}

/**
 * Flatten a docx into readable text (structure markers preserved): headings as
 * `# `, list items as `- `, table rows as `a | b`, footnotes and endnotes under
 * their own headers.
 *
 * Reads only the XML parts it needs. The full document model (parseDocx) used
 * to stand here, and it inflates every picture and builds the layout tree: a
 * 19 MB report cost a gigabyte of RSS to index six thousand characters.
 */
export async function docxToText(bytes: Uint8Array, opts: DocxTextOptions = {}): Promise<string> {
  await assertZipInflatesWithinLimits(bytes)
  const zip = await JSZip.loadAsync(bytes)
  assertZipWithinLimits(zip)
  const docPath = await mainDocumentPath(zip)
  const documentXml = docPath ? await zip.file(docPath)?.async('string') : undefined
  if (!documentXml) throw new Error('docx: main document part not found')
  const dir = docPath!.slice(0, docPath!.lastIndexOf('/') + 1)
  const styleFacts = paragraphStyleFacts(await zip.file(`${dir}styles.xml`)?.async('string'))
  const out = new TextSink(opts.maxChars ?? Infinity)
  scanBody(documentXml, styleFacts, out)
  for (const [part, title] of [
    ['footnotes.xml', 'Footnotes'],
    ['endnotes.xml', 'Endnotes'],
  ] as const) {
    if (out.full) break
    const xml = await zip.file(`${dir}${part}`)?.async('string')
    if (!xml) continue
    const notes = new TextSink(Infinity)
    scanBody(xml, styleFacts, notes, { notes: true })
    if (notes.lines.length > 0) {
      out.push(`## ${title}`)
      for (const line of notes.lines) out.push(line)
    }
  }
  return out.lines.join('\n')
}

async function mainDocumentPath(zip: JSZip): Promise<string | null> {
  if (zip.file('word/document.xml')) return 'word/document.xml'
  const rels = await zip.file('_rels/.rels')?.async('string')
  if (!rels) return null
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const tag = m[0]
    if (!/Type="[^"]*\/officeDocument"/.test(tag) || /TargetMode="External"/.test(tag)) continue
    const target = /Target="([^"]+)"/.exec(tag)?.[1]
    if (!target) continue
    const path = target.replace(/^\//, '')
    if (zip.file(path)) return path
  }
  return null
}

interface StyleFacts {
  heading: number
  list: boolean
}

/**
 * Paragraph style id -> what it implies for a paragraph: a heading level (from
 * the style's name, or its outline level) and list membership (numbering
 * attached to the style). Both are inherited along basedOn, as Word resolves
 * them.
 */
function paragraphStyleFacts(stylesXml: string | undefined): Map<string, StyleFacts> {
  const facts = new Map<string, StyleFacts>()
  if (!stylesXml) return facts
  const raw = new Map<string, { heading: number; list: boolean; basedOn?: string }>()
  for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    if (!/w:type="paragraph"/.test(m[1]!)) continue
    const id = /w:styleId="([^"]+)"/.exec(m[1]!)?.[1]
    if (!id) continue
    const body = m[2]!
    const name = /<w:name\b[^>]*w:val="([^"]*)"/.exec(body)?.[1] ?? ''
    const byName = /^heading\s*(\d)$/i.exec(name)
    const outline = /<w:outlineLvl\b[^>]*w:val="(\d)"/.exec(body)
    const heading = byName
      ? Number(byName[1])
      : /^title$/i.test(name)
        ? 1
        : outline
          ? Number(outline[1]) + 1
          : 0
    raw.set(id, {
      heading,
      // numId 0 is Word's "no numbering" override, not a list
      list: /<w:numPr\b/.test(body) && !/<w:numId\b[^>]*w:val="0"/.test(body),
      basedOn: /<w:basedOn\b[^>]*w:val="([^"]*)"/.exec(body)?.[1],
    })
  }
  for (const id of raw.keys()) {
    let heading = 0
    let list = false
    const seen = new Set<string>()
    for (let cur: string | undefined = id; cur && !seen.has(cur); cur = raw.get(cur)?.basedOn) {
      seen.add(cur)
      const r = raw.get(cur)
      if (!r) break
      if (!heading && r.heading) heading = r.heading
      if (r.list) list = true
    }
    facts.set(id, { heading, list })
  }
  return facts
}

class TextSink {
  readonly lines: string[] = []
  private chars = 0
  full = false
  constructor(private readonly maxChars: number) {}
  push(line: string): void {
    if (this.full) return
    this.lines.push(line)
    this.chars += line.length + 1
    if (this.chars >= this.maxChars) this.full = true
  }
}

interface Para {
  text: string
  heading: number
  list: boolean
}

interface Cell {
  paras: string[]
}

interface Table {
  rows: string[][]
  cell: Cell | null
  row: string[] | null
}

const TOKEN = /<(\/?)([\w:.-]+)([^>]*?)(\/?)>|([^<]+)/g

/** subtrees whose text is not part of the readable document */
const SKIP_SUBTREES = new Set([
  'mc:Fallback',
  'w:del',
  'w:moveFrom',
  'w:instrText',
  'w:delInstrText',
  'w:delText',
  'w:tabs',
  'w:tblPr',
  'w:tblGrid',
  'w:trPr',
  'w:tcPr',
  'w:sectPr',
])

/**
 * One pass over a WordprocessingML body. Paragraphs become lines (or table
 * cells); tables emit one line per row; text boxes and content controls
 * contribute their paragraphs where they sit. `mc:Fallback` is skipped so a
 * text box wrapped in AlternateContent is read once.
 */
function scanBody(
  xml: string,
  styleFacts: Map<string, StyleFacts>,
  out: TextSink,
  options: { notes?: boolean } = {},
): void {
  const paras: Para[] = []
  const tables: Table[] = []
  let skip: { name: string; depth: number } | null = null
  let depth = 0
  let inText = false
  let runHidden = false
  let noteSeparator = false
  let noteLines: string[] = []

  const emitPara = (p: Para): void => {
    const table = tables.at(-1)
    if (table?.cell) {
      table.cell.paras.push(p.text)
      return
    }
    if (options.notes) {
      noteLines.push(p.text)
      return
    }
    const line = p.heading
      ? `${'#'.repeat(Math.min(p.heading, 6))} ${p.text}`
      : p.list
        ? `- ${p.text}`
        : p.text
    out.push(line)
  }
  const emitRows = (table: Table): void => {
    const lines = table.rows.map((row) => row.join(' | '))
    const parentCell = tables.at(-1)?.cell
    for (const line of lines) {
      if (parentCell) parentCell.paras.push(line)
      else if (options.notes) noteLines.push(line)
      else out.push(line)
    }
  }
  const appendText = (text: string): void => {
    const p = paras.at(-1)
    if (p) p.text += text
  }

  TOKEN.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = TOKEN.exec(xml)) !== null) {
    if (out.full) return
    const [, close, name, attrs, selfClose, text] = m
    if (text !== undefined) {
      if (inText && !skip && !runHidden) appendText(decodeEntities(text))
      continue
    }
    const opening = !close
    const closing = !!close && !selfClose
    // a self-closing element opens and ends in one token
    const ending = closing || !!selfClose
    if (opening && !selfClose) depth++
    if (skip) {
      if (closing) {
        depth--
        if (depth < skip.depth) skip = null
      }
      continue
    }
    if (opening && SKIP_SUBTREES.has(name!)) {
      if (!selfClose) skip = { name: name!, depth }
      continue
    }
    switch (name) {
      case 'w:p':
        if (opening) paras.push({ text: '', heading: 0, list: false })
        if (ending) {
          const p = paras.pop()
          if (p && !noteSeparator) emitPara(p)
        }
        break
      case 'w:pStyle':
        if (opening) {
          const val = /w:val="([^"]*)"/.exec(attrs!)?.[1]
          const p = paras.at(-1)
          if (p && val) {
            const facts = styleFacts.get(val)
            const level = facts?.heading || Number(/^Heading(\d)$/.exec(val)?.[1] ?? 0)
            if (level) p.heading = level
            if (facts?.list) p.list = true
          }
        }
        break
      case 'w:numPr':
        if (opening) {
          const p = paras.at(-1)
          if (p) p.list = true
        }
        break
      case 'w:numId': {
        // inside the paragraph's own numPr: numId 0 switches inherited numbering off
        const p = paras.at(-1)
        if (opening && p) p.list = !/w:val="0"/.test(attrs!)
        break
      }
      case 'w:outlineLvl':
        if (opening) {
          const p = paras.at(-1)
          const val = /w:val="(\d)"/.exec(attrs!)?.[1]
          if (p && val && !p.heading) p.heading = Number(val) + 1
        }
        break
      case 'w:r':
        if (opening) runHidden = false
        break
      case 'w:vanish':
        if (opening && !/w:val="(0|false)"/.test(attrs!)) runHidden = true
        break
      case 'w:t':
        inText = opening && !selfClose
        break
      case 'w:tab':
      case 'w:ptab':
        if (opening && !runHidden) appendText('\t')
        break
      case 'w:br':
      case 'w:cr':
        // page and column breaks are layout, not text (a heading that opens a column must stay one line)
        if (opening && !runHidden && !/w:type="(page|column)"/.test(attrs!)) appendText('\n')
        break
      case 'w:noBreakHyphen':
        if (opening && !runHidden) appendText('-')
        break
      case 'w:sym':
        if (opening && !runHidden) {
          const code = /w:char="([0-9A-Fa-f]{4})"/.exec(attrs!)?.[1]
          if (code) appendText(String.fromCharCode(parseInt(code, 16)))
        }
        break
      case 'w:tbl':
        if (opening) tables.push({ rows: [], cell: null, row: null })
        if (ending) {
          const table = tables.pop()
          if (table) emitRows(table)
        }
        break
      case 'w:tr': {
        const table = tables.at(-1)
        if (!table) break
        if (opening) {
          table.row = []
          table.rows.push(table.row)
        }
        if (ending) table.row = null
        break
      }
      case 'w:tc': {
        const table = tables.at(-1)
        if (!table) break
        if (opening) table.cell = { paras: [] }
        if (ending && table.cell) {
          // a row is one line: soft breaks inside a cell must not split it
          table.row?.push(table.cell.paras.join(' ').replace(/\s*\n\s*/g, ' '))
          table.cell = null
        }
        break
      }
      case 'w:footnote':
      case 'w:endnote':
        if (opening) {
          noteSeparator = /w:type="(separator|continuationSeparator|continuationNotice)"/.test(
            attrs!,
          )
          noteLines = []
        } else {
          if (!noteSeparator && noteLines.length > 0) out.push(noteLines.join(' '))
          noteSeparator = false
        }
        break
    }
    if (closing) depth--
  }
}

function decodeEntities(text: string): string {
  if (text.indexOf('&') < 0) return text
  return text.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_, e: string) => {
    if (e[0] === '#') {
      const code =
        e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)
      return Number.isFinite(code) ? String.fromCodePoint(code) : ''
    }
    return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[e] ?? ''
  })
}
