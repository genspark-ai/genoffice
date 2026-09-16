import type { Block, FieldDisplay } from '@genoffice/docx-engine'
import { excerpt, PLACEHOLDER, type IssueDraft } from '../check'
import { headerFooterState, listComments, listRevisions, type OpenDocument } from './docx'

export const DOCS_CHECKS = [
  'field_not_evaluated',
  'broken_ref',
  'toc_stale',
  'missing_image',
  'chart_empty',
  'empty_heading',
  'heading_skip',
  'placeholder_left',
  'pending_revisions',
  'unresolved_comments',
] as const

const LIST_CAP = 20

interface PmBlock {
  index: number
  name: string
  text: string
  level?: number
  docxIndex?: number
  field?: FieldDisplay
  brokenImage?: boolean
  chart?: { series: { values: (number | null)[] }[] }
}

export function checkDocument(doc: OpenDocument): IssueDraft[] {
  const blocks = pmBlocks(doc)
  const issues: IssueDraft[] = []
  const blocksByDocx = new Map<number, number>()
  for (const b of blocks) if (b.docxIndex !== undefined) blocksByDocx.set(b.docxIndex, b.index)
  const parsed = doc.parsed.blocks as Block[]
  const at = (docxIndex: number) => blocksByDocx.get(docxIndex) ?? docxIndex

  checkFields(blocks, parsed, at, issues)
  checkRefs(parsed, at, issues)
  checkToc(blocks, issues)
  for (const b of blocks) {
    if (b.brokenImage) {
      issues.push(
        issue(
          'missing_image',
          'error',
          b,
          'image part is missing from the package; the block shows nothing',
        ),
      )
    }
    if (b.chart && b.chart.series.every((s) => s.values.every((v) => v === null))) {
      issues.push(issue('chart_empty', 'warning', b, 'chart has no data in any series'))
    }
    if (b.name === 'docHeading' && !b.text.trim()) {
      issues.push(
        issue('empty_heading', 'warning', b, `heading level ${b.level ?? '?'} has no text`),
      )
    }
  }
  let lastLevel = 0
  for (const b of blocks) {
    if (b.name !== 'docHeading' || !b.level || !b.text.trim()) continue
    if (lastLevel && b.level > lastLevel + 1) {
      issues.push(
        issue(
          'heading_skip',
          'info',
          b,
          `heading level jumps from ${lastLevel} to ${b.level}`,
          b.text,
        ),
      )
    }
    lastLevel = b.level
  }
  checkPlaceholders(doc, blocks, issues)
  const revisions = listRevisions(doc)
  if (revisions.length) {
    const kinds = new Map<string, number>()
    for (const r of revisions) kinds.set(r.kind, (kinds.get(r.kind) ?? 0) + 1)
    issues.push({
      code: 'pending_revisions',
      level: 'info',
      path: `block[${revisions[0]!.blockIndex}]`,
      blockIndex: revisions[0]!.blockIndex,
      message: `${revisions.length} tracked change(s) pending (${[...kinds].map(([k, n]) => `${k} ${n}`).join(', ')}); \`genoffice docs read --revisions\` lists them`,
    })
  }
  const open = listComments(doc).filter((c) => !c.done)
  if (open.length) {
    const first = open.find((c) => c.blockIndex !== undefined)
    issues.push({
      code: 'unresolved_comments',
      level: 'info',
      path: first?.blockIndex !== undefined ? `block[${first.blockIndex}]` : 'comments',
      ...(first?.blockIndex !== undefined ? { blockIndex: first.blockIndex } : {}),
      message: `${open.length} comment thread(s) unresolved; \`genoffice docs read --comments\` lists them`,
    })
  }
  return issues
}

function pmBlocks(doc: OpenDocument): PmBlock[] {
  const out: PmBlock[] = []
  doc.editor.state.doc.forEach((node, _offset, index) => {
    const attrs = node.attrs as Record<string, unknown>
    out.push({
      index,
      name: node.type.name,
      text: node.textContent,
      ...(typeof attrs.level === 'number' ? { level: attrs.level } : {}),
      ...(typeof attrs.docxIndex === 'number' ? { docxIndex: attrs.docxIndex } : {}),
      ...(attrs.fieldDisplay ? { field: attrs.fieldDisplay as FieldDisplay } : {}),
      ...(attrs.brokenImage ? { brokenImage: true } : {}),
      ...(attrs.chartDisplay ? { chart: attrs.chartDisplay as PmBlock['chart'] } : {}),
    })
  })
  return out
}

function issue(
  code: string,
  level: IssueDraft['level'],
  b: PmBlock,
  message: string,
  context?: string,
): IssueDraft {
  return {
    code,
    level,
    path: `block[${b.index}]`,
    blockIndex: b.index,
    message,
    ...(context ? { context: excerpt(context, 80) } : {}),
  }
}

function checkFields(
  blocks: PmBlock[],
  parsed: Block[],
  at: (docxIndex: number) => number,
  issues: IssueDraft[],
): void {
  const toc = blocks.filter((b) => b.field?.kind === 'tocLine' && !b.field.deleted)
  const blank = toc.filter((b) => !(b.field!.right ?? '').trim())
  if (blank.length) {
    issues.push(
      issue(
        'field_not_evaluated',
        'warning',
        blank[0]!,
        `table of contents: ${blank.length} of ${toc.length} entries have no page number; Word computes them on open, or open and save in GenOffice`,
      ),
    )
  }
  parsed.forEach((block, docxIndex) => {
    for (const keyword of emptyFieldResults(block.originalXml)) {
      const index = at(docxIndex)
      issues.push({
        code: 'field_not_evaluated',
        level: 'warning',
        path: `block[${index}]`,
        blockIndex: index,
        message: `${keyword || 'field'} has no cached result; it shows empty until Word updates fields`,
      })
    }
  })
  parsed.forEach((block, docxIndex) => {
    if (block.originalXml && /<w:fldChar\b[^>]*\bw:dirty="(?:true|1)"/.test(block.originalXml)) {
      const index = at(docxIndex)
      issues.push({
        code: 'field_not_evaluated',
        level: 'info',
        path: `block[${index}]`,
        blockIndex: index,
        message: `${fieldKeyword(block.originalXml) || 'field'} is marked dirty; Word recomputes it on open`,
      })
    }
  })
}

const RESULTLESS_FIELDS = new Set(['TOC', 'XE', 'TA', 'INDEX', 'TC', 'PRIVATE', 'RD'])

/** Keywords of complex or simple fields whose result run holds no text (the TOC has its own check). */
function emptyFieldResults(xml: string | null | undefined): string[] {
  if (!xml) return []
  const out: string[] = []
  const keywordOf = (instr: string) => /^\s*\\?\s*([A-Za-z]+)/.exec(instr)?.[1]?.toUpperCase() ?? ''
  for (const m of xml.matchAll(
    /<w:fldSimple\b([^>]*)>([\s\S]*?)<\/w:fldSimple>|<w:fldSimple\b([^>]*)\/>/g,
  )) {
    const instr = /\bw:instr="([^"]*)"/.exec(m[1] ?? m[3] ?? '')?.[1] ?? ''
    const keyword = keywordOf(unescapeAttr(instr))
    if (!RESULTLESS_FIELDS.has(keyword) && !/<w:t\b[^>]*>[^<]*\S/.test(m[2] ?? ''))
      out.push(keyword)
  }
  let instr: string[] | null = null
  let result: string | null = null
  for (const t of xml.matchAll(
    /<w:fldChar\b[^>]*\bw:fldCharType="(begin|separate|end)"|<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>|<w:t\b[^>]*>([^<]*)<\/w:t>/g,
  )) {
    if (t[1] === 'begin') {
      instr = []
      result = null
    } else if (t[1] === 'separate') {
      result = ''
    } else if (t[1] === 'end') {
      if (instr) {
        const keyword = keywordOf(instr.join(''))
        if (!RESULTLESS_FIELDS.has(keyword) && !(result ?? '').trim()) out.push(keyword)
      }
      instr = null
      result = null
    } else if (t[2] !== undefined && instr && result === null) instr.push(t[2])
    else if (t[3] !== undefined && result !== null) result += t[3]
  }
  return out
}

function fieldKeyword(xml: string | null | undefined): string {
  if (!xml) return ''
  const instr =
    /<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/.exec(xml)?.[1] ??
    /<w:fldSimple[^>]*w:instr="([^"]*)"/.exec(xml)?.[1] ??
    ''
  return /^\s*\\?\s*([A-Za-z]+)/.exec(instr)?.[1]?.toUpperCase() ?? ''
}

function checkRefs(parsed: Block[], at: (docxIndex: number) => number, issues: IssueDraft[]): void {
  const bookmarks = new Set<string>()
  for (const b of parsed) {
    for (const name of b.bookmarks ?? []) bookmarks.add(name)
    for (const name of b.hiddenBookmarks ?? []) bookmarks.add(name)
  }
  const seen = new Set<string>()
  parsed.forEach((block, docxIndex) => {
    const xml = block.originalXml
    if (!xml) return
    const targets: { kind: string; name: string }[] = []
    for (const instr of fieldInstructions(xml)) {
      const ref = /^(REF|PAGEREF|NOTEREF)\s+"?([^\s"\\]+)/i.exec(instr)
      if (ref) targets.push({ kind: ref[1]!.toUpperCase(), name: ref[2]! })
      const link = /^HYPERLINK\b[\s\S]*?\\l\s+"?([^\s"\\]+)/i.exec(instr)
      if (link) targets.push({ kind: 'HYPERLINK', name: link[1]! })
    }
    for (const m of xml.matchAll(/<w:hyperlink\b[^>]*\bw:anchor="([^"]+)"/g)) {
      targets.push({ kind: 'hyperlink', name: m[1]! })
    }
    for (const t of targets) {
      if (bookmarks.has(t.name) || seen.has(t.name)) continue
      seen.add(t.name)
      const index = at(docxIndex)
      issues.push({
        code: 'broken_ref',
        level: 'error',
        path: `block[${index}]`,
        blockIndex: index,
        message: `${t.kind} points at bookmark "${t.name}", which does not exist`,
      })
    }
  })
}

/** Field codes, one string per field: Word splits an instruction over several instrText runs. */
function fieldInstructions(xml: string): string[] {
  const out: string[] = []
  for (const m of xml.matchAll(/<w:fldSimple\b[^>]*\bw:instr="([^"]*)"/g)) {
    out.push(unescapeAttr(m[1]!).trim())
  }
  let current: string[] | null = null
  const tokens = xml.matchAll(
    /<w:fldChar\b[^>]*\bw:fldCharType="(begin|separate|end)"|<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g,
  )
  for (const t of tokens) {
    if (t[1] === 'begin') current = []
    else if (t[1] === 'separate' || t[1] === 'end') {
      if (current) out.push(current.join('').trim())
      current = null
    } else if (t[2] !== undefined && current) current.push(t[2])
  }
  return out
}

function unescapeAttr(text: string): string {
  return text
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
}

function checkToc(blocks: PmBlock[], issues: IssueDraft[]): void {
  const toc = blocks.filter((b) => b.field?.kind === 'tocLine' && !b.field.deleted)
  if (!toc.length) return
  const headings = blocks
    .filter((b) => b.name === 'docHeading' && b.text.trim())
    .map((b) => b.text.trim())
  const entries = toc.map((b) => (b.field!.left ?? '').trim())
  const pool = [...headings]
  const unmatched: string[] = []
  for (const e of entries) {
    const i = pool.indexOf(e)
    if (i >= 0) pool.splice(i, 1)
    else unmatched.push(e)
  }
  if (!unmatched.length && !pool.length) return
  const parts: string[] = []
  if (unmatched.length) {
    parts.push(
      `${unmatched.length} entr${unmatched.length === 1 ? 'y' : 'ies'} match no heading (${list(unmatched)})`,
    )
  }
  if (pool.length) parts.push(`${pool.length} heading(s) missing from it (${list(pool)})`)
  issues.push(
    issue(
      'toc_stale',
      'warning',
      toc[0]!,
      `table of contents is out of date: ${parts.join('; ')}; rebuild it with insertToc after deleting the old lines`,
    ),
  )
}

function list(items: readonly string[]): string {
  const shown = items.slice(0, 5).map((t) => `"${excerpt(t, 40)}"`)
  return items.length > 5 ? `${shown.join(', ')}, …` : shown.join(', ')
}

function checkPlaceholders(doc: OpenDocument, blocks: PmBlock[], issues: IssueDraft[]): void {
  const hits = blocks.filter((b) => PLACEHOLDER.test(b.text))
  if (hits.length) {
    issues.push({
      code: 'placeholder_left',
      level: 'warning',
      path: `block[${hits[0]!.index}]`,
      blockIndex: hits[0]!.index,
      message: `${hits.length} block(s) contain placeholder text (template keys, TODO/TBD, lorem ipsum, [insert …])`,
      context:
        hits
          .slice(0, LIST_CAP)
          .map((b) => `#${b.index} "${excerpt(b.text, 40)}"`)
          .join(', ') + (hits.length > LIST_CAP ? ` … (${hits.length} blocks)` : ''),
    })
  }
  const hf = headerFooterState(doc)
  for (const [slot, text] of Object.entries(hf)) {
    if (typeof text === 'string' && PLACEHOLDER.test(text)) {
      issues.push({
        code: 'placeholder_left',
        level: 'warning',
        path: slot,
        message: `${slot} contains placeholder text`,
        context: excerpt(text, 80),
      })
    }
  }
}
