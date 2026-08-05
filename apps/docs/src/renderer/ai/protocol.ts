import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import {
  TABLE_HEADER_FILL,
  type Block,
  type TableCell,
  type TableModel,
} from '@genoffice/docx-engine'
import { pmTableToModel, tableModelToPmNode, type PmMark, type PmNode } from '../editor/convert'
import { equationBlockJson, inlineEquationNodeJson } from '../editor/equation'
import { TRACK_IGNORE } from '../editor/revisions'
import { countWords } from '../word-count'

/**
 * Agent protocol: the model runs a multi-turn tool-use loop — it reads the document skeleton
 * (block list with numbered addressing), reads full block content on demand,
 * and mutates the document exclusively through tools (tools.ts). Questions
 * are answered directly in chat without touching the document. Context size
 * stays bounded: previews are clipped and full content is pulled lazily.
 */

// ---- context budgets (characters, ≈4 chars/token) ----

const SELECTION_MAX_CHARS = 24_000
const DOC_CONTEXT_MAX_CHARS = 8_000
const PREVIEW_MAX_CHARS = 60
const PREVIEW_TIGHT_CHARS = 20

/**
 * Guide for the apply_commands tool: format,
 * structure and batch operations travel as a batchUpdate-style JSON envelope
 * executed deterministically by commands.ts.
 */
export const COMMANDS_GUIDE = [
  "The command format mimics documents.batchUpdate of the Google Docs API: each command is a single-key object; the fields array has the same semantics as Google's FieldMask — only the listed fields are updated, everything not listed is left untouched. Differences from Google: target uses semantic addressing (node type / text containment / block index), colors are 6-digit hex (no #), lengths are in twips, font sizes are in half-points.",
  '',
  'The commands input is an array of commands executed in order. Command schema (TypeScript):',
  '```',
  'interface Target {  // all conditions are ANDed; provide at least one',
  "  nodeType?: 'docHeading' | 'docParagraph' | 'docListItem' | 'image'  // image = image block",
  "  headingLevel?: number        // only together with nodeType: 'docHeading'",
  '  containsText?: string        // block plain text contains this substring',
  '  matchCase?: boolean          // defaults to true',
  '  blockIndexes?: number[]      // indexes from the "document block list"',
  "  scope?: 'selection' | 'document'  // selection = only blocks covered by the current selection",
  '}',
  "{ updateTextStyle: { target, style: { color?, highlight?, sizeHalfPoints?, font?, bold?, italic?, underline?, strike?, baselineOffset?: 'SUPERSCRIPT'|'SUBSCRIPT'|'NONE', link?: {url}|null }, fields: string[] } }  // value null = clear that property",
  "{ updateParagraphStyle: { target, style: { align?: 'left'|'center'|'right'|'justify', lineSpacing?, indentLeft?, indentRight?, indentFirstLine?, spaceBefore?, spaceAfter?, pageBreakBefore?, shadingFill?, borders?: subset of 'tblr'|null }, fields: string[] } }  // indentFirstLine positive = first-line indent, negative = hanging; a two-character indent for CJK text ≈ font size in pt × 40 twips",
  '{ setHeadingLevel: { target, level: 0-6 } }  // 0 = demote to body paragraph',
  '{ replaceAllText: { containsText, replaceText, matchCase? } }',
  '{ deleteBlocks: { target } }',
  '{ moveBlocks: { blockIndexes: number[], afterBlockIndex: number } }  // -1 = start of document',
  '{ createParagraphBullets: { target, bulletPreset? } }  // paragraphs to list; BULLET_* prefix = unordered, NUMBERED_* = ordered (default unordered); heading blocks are not converted',
  '{ deleteParagraphBullets: { target } }  // list items back to body paragraphs',
  "{ updateImageProperties: { target, properties: { widthPx?, heightPx?, align?: 'left'|'center'|'right' }, fields } }  // image blocks only; giving one dimension scales proportionally",
  '{ insertToc: { afterBlockIndex } }  // insert a TOC field after that block index (-1 = start of document); entries come from current headings, Word computes page numbers on open; fails if the document has no headings',
  '```',
  '',
  'Common recipes:',
  '- "Make all headings red" → commands: [{"updateTextStyle":{"target":{"nodeType":"docHeading"},"style":{"color":"FF0000"},"fields":["color"]}}]',
  '- "1.5 line spacing for the whole document" → one updateParagraphStyle each for docParagraph / docHeading / docListItem, style {"lineSpacing":1.5}, fields ["lineSpacing"]',
  '- "Demote the \'Risk Notice\' level-2 heading to level 3" → commands: [{"setHeadingLevel":{"target":{"nodeType":"docHeading","headingLevel":2,"containsText":"Risk Notice"},"level":3}}]',
  '- "Delete blocks 3 to 5" → commands: [{"deleteBlocks":{"target":{"blockIndexes":[3,4,5]}}}]',
  '- "Turn blocks 2 to 4 into a numbered list" → commands: [{"createParagraphBullets":{"target":{"blockIndexes":[2,3,4]},"bulletPreset":"NUMBERED_DECIMAL_ALPHA_ROMAN"}}]',
  '- "Center all images" → commands: [{"updateImageProperties":{"target":{"nodeType":"image"},"properties":{"align":"center"},"fields":["align"]}}]',
  '- "Two-character first-line indent for the whole document" (11 pt font) → one updateParagraphStyle for docParagraph, style {"indentFirstLine":440}, fields ["indentFirstLine"]',
  '- "Insert a table of contents at the beginning" → commands: [{"insertToc":{"afterBlockIndex":-1}}]; if the user wants the TOC on its own page, also apply updateParagraphStyle to the first body block after the TOC with style {"pageBreakBefore":true}, fields ["pageBreakBefore"]',
  '- Cover page recipe: use insert_content to insert one paragraph each for title/subtitle/date at the start of the document, then updateTextStyle to enlarge the title font (e.g. {"sizeHalfPoints":72}, fields ["sizeHalfPoints"]), updateParagraphStyle to center everything and give the title {"spaceBefore":4800}, and finally set {"pageBreakBefore":true} on the first body block after the cover',
  '',
  'Discipline:',
  '- Read the "document block list" before issuing commands; anything like "paragraph N / a certain section" must be addressed with block indexes from the list — never guess;',
  '- Put only the properties the user explicitly asked for into fields;',
  '- Protected blocks such as images cannot be modified with style commands (they are skipped); to change table content, use read_blocks to get the <table> and rewrite it wholesale with replace_blocks.',
  '',
  'Known error cases (must avoid):',
  'BC-1 Putting properties the user did not ask for into fields, wiping existing formatting by mistake;',
  'BC-2 A "whole document" formatting operation only changed docParagraph and missed docHeading/docListItem;',
  'BC-3 Addressing by word/character position (no such addressing exists); use block indexes or text containment instead;',
  'BC-4 Using replaceAllText sentence by sentence for rewrites such as translation/abbreviation; use the replace_blocks tool instead;',
].join('\n')

const HTML_RULES = [
  'The html tool input is a restricted HTML fragment. Rules:',
  '- Only these tags are allowed: h1 h2 h3 h4 h5 h6 p ul ol li strong em u s a br table thead tbody tr th td pre code blockquote',
  '- Tables: use <th> for the first (header) row; cells contain plain text only (<br> may split lines); nested tables / merged cells are not supported; once inserted the table is protected as a whole, only cell text remains editable',
  '- Use <pre> for code samples (monospace font + shading, line breaks preserved); use <blockquote> for quotations (indent + left bar)',
  '- Use <formula>LaTeX</formula> for math (produces native Word equations): as a top-level block it becomes its own centered paragraph; placed inside <p>/<li>/<h*> text it is an inline formula flowing with the text, e.g. <p>From <formula>E = mc^2</formula> we know…</p>; supports the common subset of \\frac \\sqrt super/subscripts \\sum \\int \\lim matrix environments Greek letters etc., the align environment is not supported; invalid LaTeX fails the whole call — fix and retry',
  '- Do not include <html>/<body>, markdown code fences, or explanatory text',
  '- Organize long content into sections with h2/h3 (unless the user only wants a single paragraph)',
  '- Keep the same language as the original document / user instruction, unless translation is requested',
].join('\n')

/**
 * Agent system prompt: document-first, intent resolution, act via tools,
 * answer in chat.
 */
export const AGENT_SYSTEM_PROMPT = [
  'You are the document assistant built into the local document editor GenOffice Docs. You read and modify the currently open document exclusively through tools; there is no other modification channel.',
  '',
  '# Intent resolution',
  '- The user asks to modify/generate/translate/format → call the appropriate tools, then summarize what was done in one or two sentences;',
  '- The user is asking a question or consulting (word count, structure, "what is this about", writing advice, etc.) → answer directly in plain text without calling modification tools;',
  '- For statistics such as word count or block count, quote the "full-text stats" from the "document block list" directly — do not count yourself;',
  '- When intent is unclear, read the document first (the block list in the message, and read_blocks for full text if needed), then decide.',
  '',
  '# Tool usage',
  '- Every user message carries the latest "document block list" (index|type|content preview; previews may be truncated); after modifications, call get_document_context if you need the latest state;',
  '- When a list preview is truncated, read the full content with read_blocks before rewriting; never rewrite based on a truncated preview;',
  '- Content changes: use insert_content for new content, and replace_blocks to rewrite/replace existing blocks (pass a block index range and the new HTML);',
  '- Formatting, structure, and batch operations (color/font size/line spacing/alignment/indent/heading level/find & replace/delete/move/list conversion) go through apply_commands — do not rewrite whole blocks with replace_blocks;',
  '- When the user has text selected, the message includes the selection block indexes and content; rewrite-style requests apply to the selection by default;',
  '- Web search: use web_search when you need up-to-date information/data/fact checking; search before writing about uncertain facts — do not fabricate;',
  '- Illustrations: when the user wants pictures, first image_search (English keywords work better) → pick a suitable result → insert_image with its imageUrl to insert into the document;',
  '- Charts: use insert_chart for data visualization (bar/line/pie; saved as native Word charts); use edit_chart to change the data of an existing chart block in the block list; data must be real, from the document or search results;',
  '- One reply may chain multiple tools; after everything is done, always finish with a short plain-text summary.',
  '',
  '# HTML fragment rules',
  HTML_RULES,
  '',
  '# apply_commands command guide',
  COMMANDS_GUIDE,
].join('\n')

export { countWords }

// ---- selection scope ----

export interface SelectionScope {
  /** top-level child indexes covered by the selection (inclusive) */
  startIndex: number
  endIndex: number
  /** true when the user has an actual range selected (not just a caret) */
  isRange: boolean
}

export function getSelectionScope(editor: Editor): SelectionScope {
  const { from, to, empty } = editor.state.selection
  const doc = editor.state.doc
  let startIndex = -1
  let endIndex = -1
  let index = 0
  doc.forEach((node, offset) => {
    const nodeFrom = offset
    const nodeTo = offset + node.nodeSize
    if (nodeTo > from && nodeFrom < to) {
      if (startIndex === -1) startIndex = index
      endIndex = index
    } else if (empty && from >= nodeFrom && from <= nodeTo && startIndex === -1) {
      startIndex = index
      endIndex = index
    }
    index++
  })
  if (startIndex === -1) {
    startIndex = doc.childCount - 1
    endIndex = doc.childCount - 1
  }
  return { startIndex, endIndex, isRange: !empty }
}

/** ProseMirror positions of a top-level child index range */
export function blockRangePositions(
  editor: Editor,
  startIndex: number,
  endIndex: number,
): { from: number; to: number } {
  const doc = editor.state.doc
  let from = 0
  let to = 0
  let index = 0
  doc.forEach((node, offset) => {
    if (index === startIndex) from = offset
    if (index === endIndex) to = offset + node.nodeSize
    index++
  })
  return { from, to }
}

// ---- blocks -> restricted HTML ----

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function inlineToHtml(content: PmNode[] | undefined): string {
  if (!content) return ''
  let html = ''
  for (const node of content) {
    if (node.type === 'hardBreak') {
      html += '<br>'
      continue
    }
    if (node.type === 'docInlineMath') {
      // editor-created formulas round-trip as LaTeX; parsed-from-docx ones
      // degrade to the flat token strip (structure lost if the model rewrites)
      const source = String(node.attrs?.latex ?? '') || String(node.attrs?.text ?? '')
      if (source) html += `<formula>${escapeHtml(source)}</formula>`
      continue
    }
    if (node.type === 'docRuby') {
      html += escapeHtml(String(node.attrs?.base ?? ''))
      continue
    }
    if (node.type !== 'text' || !node.text) continue
    let text = escapeHtml(node.text)
    const marks = node.marks ?? []
    const has = (type: string) => marks.some((m: PmMark) => m.type === type)
    if (has('bold')) text = `<strong>${text}</strong>`
    if (has('italic')) text = `<em>${text}</em>`
    if (has('underline')) text = `<u>${text}</u>`
    if (has('strike')) text = `<s>${text}</s>`
    const link = marks.find((m: PmMark) => m.type === 'link')
    if (link?.attrs?.href) text = `<a href="${escapeHtml(String(link.attrs.href))}">${text}</a>`
    html += text
  }
  return html
}

/** plain text of a paragraph's inline content, hard breaks as newlines */
function inlineToPlainText(content: PmNode[] | undefined): string {
  if (!content) return ''
  return content
    .map((n) => (n.type === 'hardBreak' ? '\n' : n.type === 'text' ? (n.text ?? '') : ''))
    .join('')
}

/** paragraph carries the <pre> preset (see CODE_BLOCK_PRESET) */
function isCodeParagraph(node: PmNode): boolean {
  return (
    node.attrs?.shadingFill === CODE_BLOCK_PRESET.shadingFill &&
    node.attrs?.borders === CODE_BLOCK_PRESET.borders
  )
}

/** paragraph carries the <blockquote> preset (see QUOTE_BLOCK_PRESET) */
function isQuoteParagraph(node: PmNode): boolean {
  return node.attrs?.borders === QUOTE_BLOCK_PRESET.borders && Number(node.attrs?.indentLeft) > 0
}

/** table block -> restricted <table> HTML so the model can read and rewrite it */
function tableToHtml(table: TableModel): string {
  const headerRow =
    table.rows.length > 1 && table.rows[0].every((c) => c.bold || c.fill !== undefined)
  const rows = table.rows.map((row, r) => {
    const tag = headerRow && r === 0 ? 'th' : 'td'
    const cells = row
      .map((cell) => `<${tag}>${cell.paras.map(escapeHtml).join('<br>')}</${tag}>`)
      .join('')
    return `<tr>${cells}</tr>`
  })
  return `<table>${rows.join('')}</table>`
}

/**
 * Serialize top-level nodes [startIndex, endIndex] into restricted HTML.
 * Consecutive list items of the same kind are grouped into ul/ol.
 */
export function serializeRangeToHtml(editor: Editor, startIndex: number, endIndex: number): string {
  const json = editor.getJSON() as PmNode
  const children = (json.content ?? []).slice(startIndex, endIndex + 1)
  const parts: string[] = []
  let listBuffer: { kind: string; items: string[] } | null = null

  const flushList = () => {
    if (!listBuffer) return
    const tag = listBuffer.kind === 'ordered' ? 'ol' : 'ul'
    parts.push(`<${tag}>${listBuffer.items.join('')}</${tag}>`)
    listBuffer = null
  }

  for (const node of children) {
    if (node.type === 'docHeading') {
      flushList()
      const level = Math.min(Math.max(Number(node.attrs?.level) || 1, 1), 6)
      parts.push(`<h${level}>${inlineToHtml(node.content)}</h${level}>`)
    } else if (node.type === 'docListItem') {
      const kind = (node.attrs?.kind as string) ?? 'bullet'
      if (!listBuffer || listBuffer.kind !== kind) {
        flushList()
        listBuffer = { kind, items: [] }
      }
      listBuffer.items.push(`<li>${inlineToHtml(node.content)}</li>`)
    } else if (node.type === 'docParagraph') {
      flushList()
      if (isCodeParagraph(node)) {
        parts.push(`<pre>${escapeHtml(inlineToPlainText(node.content))}</pre>`)
      } else if (isQuoteParagraph(node)) {
        parts.push(`<blockquote>${inlineToHtml(node.content)}</blockquote>`)
      } else {
        parts.push(`<p>${inlineToHtml(node.content)}</p>`)
      }
    } else if (node.type === 'docTable') {
      flushList()
      parts.push(tableToHtml(pmTableToModel(node)))
    } else {
      flushList()
      const formula = node.attrs?.formulaDisplay as { tokens?: string[]; latex?: string } | null
      if (formula?.latex) {
        // full LaTeX recovered: the model can rewrite the formula losslessly
        parts.push(`<formula>${escapeHtml(formula.latex)}</formula>`)
      } else if (formula?.tokens?.length) {
        parts.push(
          `<p>[Protected formula: ${escapeHtml(formula.tokens.join(' '))}, structure is read-only, kept as is]</p>`,
        )
      } else {
        const label = String(node.attrs?.label ?? node.attrs?.blockType ?? 'content')
        parts.push(`<p>[Protected content: ${escapeHtml(label)}, kept as is]</p>`)
      }
    }
  }
  flushList()
  return parts.join('\n')
}

// ---- document context + layered request builders ----

interface ContextEntry {
  index: number
  type: string
  preview: string
  isHeading: boolean
}

/**
 * Numbered structure snapshot for command-mode addressing: one line per
 * top-level block. Block numbers MUST equal the live PM doc order at execution
 * time, so this is rebuilt per request and never cached.
 */
export function buildDocumentContext(editor: Editor): string {
  const entries: ContextEntry[] = []
  let index = 0
  editor.state.doc.forEach((node) => {
    let type: string
    let preview: string
    let isHeading = false
    if (node.type.name === 'docHeading') {
      type = `h${Math.min(Math.max(Number(node.attrs.level) || 1, 1), 6)}`
      preview = node.textContent
      isHeading = true
    } else if (node.type.name === 'docListItem') {
      type = 'li'
      preview = node.textContent
    } else if (node.type.name === 'docTable') {
      type = 'table'
      preview = node.textContent
    } else if (node.type.name === 'docProtected') {
      type = String(node.attrs.label || node.attrs.blockType || 'protected')
      preview = String(node.attrs.previewText ?? '')
    } else {
      type = 'p'
      preview = node.textContent
    }
    entries.push({ index, type, preview: preview.replace(/\s+/g, ' ').trim(), isHeading })
    index++
  })

  const render = (bodyMax: number) =>
    entries.map(
      (e) => `${e.index}|${e.type}|${clip(e.preview, e.isHeading ? PREVIEW_MAX_CHARS : bodyMax)}`,
    )
  let lines = render(PREVIEW_MAX_CHARS)
  if (lines.join('\n').length > DOC_CONTEXT_MAX_CHARS) lines = render(PREVIEW_TIGHT_CHARS)
  if (lines.join('\n').length > DOC_CONTEXT_MAX_CHARS) {
    // keep both ends (numbering must stay verifiable), elide the middle
    let dropStart = Math.floor(lines.length / 3)
    let dropEnd = lines.length - Math.floor(lines.length / 3)
    while (
      dropStart > 1 &&
      dropEnd < lines.length - 1 &&
      [...lines.slice(0, dropStart), ...lines.slice(dropEnd)].join('\n').length >
        DOC_CONTEXT_MAX_CHARS
    ) {
      dropStart = Math.max(1, dropStart - 10)
      dropEnd = Math.min(lines.length - 1, dropEnd + 10)
    }
    lines = [
      ...lines.slice(0, dropStart),
      `…(${dropEnd - dropStart} blocks elided here; numbering is continuous, extrapolate indexes if needed)…`,
      ...lines.slice(dropEnd),
    ]
  }

  const scope = getSelectionScope(editor)
  const selLine = scope.isRange
    ? scope.startIndex === scope.endIndex
      ? `Current selection: block ${scope.startIndex}`
      : `Current selection: blocks ${scope.startIndex}-${scope.endIndex}`
    : `No selection; cursor is in block ${scope.startIndex}`

  // authoritative stats for answer mode (previews above may be clipped)
  const fullText = editor.state.doc.textContent
  const statsLine = `Full-text stats: words ${countWords(fullText)}, characters (no spaces) ${
    fullText.replace(/\s/g, '').length
  }, characters (with spaces) ${fullText.length}`

  return [
    `The document has ${entries.length} blocks, listed below (index|type|content preview):`,
    ...lines,
    statsLine,
    selLine,
  ].join('\n')
}

function clip(text: string, max: number): string {
  return text.length > max ? text.slice(0, max) + '…' : text
}

/** Blank = exactly one empty paragraph; a textContent check would misread image/chart-only documents as blank. */
export function isBlankDocument(editor: Editor): boolean {
  const doc = editor.state.doc
  if (doc.childCount !== 1) return false
  const first = doc.child(0)
  return first.type.name === 'docParagraph' && first.content.size === 0
}

/**
 * Per-turn context: fresh document skeleton + selection fragment injected
 * as a selection chip.
 */
export function buildDocContext(editor: Editor): string {
  const isEmptyDoc = isBlankDocument(editor)
  const scope = getSelectionScope(editor)
  const selectionHtml = scope.isRange
    ? clip(serializeRangeToHtml(editor, scope.startIndex, scope.endIndex), SELECTION_MAX_CHARS)
    : ''
  return [
    isEmptyDoc
      ? 'The document is currently blank.'
      : `Document block list:\n${buildDocumentContext(editor)}`,
    selectionHtml
      ? `Content selected by the user (blocks ${scope.startIndex}-${scope.endIndex}):\n${selectionHtml}`
      : '',
  ]
    .filter(Boolean)
    .join('\n\n')
}

// ---- restricted HTML fragment -> PmNode[] ----

export interface NumIds {
  bullet: string | null
  ordered: string | null
}

/**
 * Preset paragraph formats for <pre> / <blockquote>: combinations of existing
 * Block-model properties, so these blocks stay fully editable and round-trip
 * byte-stable through parse -> generate -> parse (no schema/signature change).
 */
export const CODE_BLOCK_PRESET = {
  font: 'Consolas',
  shadingFill: 'F2F2F2',
  borders: 'tblr',
} as const
export const QUOTE_BLOCK_PRESET = {
  color: '666666',
  indentLeft: 720,
  borders: 'l',
} as const

/** pick an existing numbering id of the right kind so new list items join real docx numbering */
export function findNumId(blocks: Block[], kind: 'bullet' | 'ordered'): string | null {
  for (const block of blocks) {
    if (block.type === 'listItem' && block.list?.kind === kind) return block.list.numId
  }
  return null
}

const INLINE_MARK_TAGS: Record<string, string> = {
  strong: 'bold',
  b: 'bold',
  em: 'italic',
  i: 'italic',
  u: 'underline',
  s: 'strike',
  del: 'strike',
  strike: 'strike',
}

function parseInline(element: Node, marks: PmMark[]): PmNode[] {
  const nodes: PmNode[] = []
  element.childNodes.forEach((child) => {
    if (child.nodeType === Node.TEXT_NODE) {
      const text = (child.textContent ?? '').replace(/\s+/g, ' ')
      if (text)
        nodes.push({ type: 'text', text, ...(marks.length > 0 ? { marks: [...marks] } : {}) })
      return
    }
    if (child.nodeType !== Node.ELEMENT_NODE) return
    const el = child as Element
    const tag = el.tagName.toLowerCase()
    if (tag === 'br') {
      nodes.push({ type: 'hardBreak' })
      return
    }
    if (tag === 'a') {
      const href = el.getAttribute('href') ?? ''
      nodes.push(...parseInline(el, [...marks, { type: 'link', attrs: { href, rId: null } }]))
      return
    }
    if (tag === 'formula') {
      const latex = (el.textContent ?? '').trim()
      if (!latex) return
      try {
        nodes.push(inlineEquationNodeJson(latex))
      } catch (e) {
        const reason = e instanceof Error ? e.message : String(e)
        throw new Error(`Cannot parse the LaTeX in <formula> (${reason}): ${latex}`, { cause: e })
      }
      return
    }
    const markType = INLINE_MARK_TAGS[tag]
    if (markType) {
      const next = marks.some((m) => m.type === markType) ? marks : [...marks, { type: markType }]
      nodes.push(...parseInline(el, next))
      return
    }
    // unknown inline tag (span etc.): keep its text content
    nodes.push(...parseInline(el, marks))
  })
  return nodes
}

function blockNode(type: string, attrs: Record<string, unknown>, content: PmNode[]): PmNode {
  const node: PmNode = {
    type,
    attrs: { docxIndex: null, styleId: null, aiChanged: true, ...attrs },
  }
  if (content.length > 0) node.content = content
  return node
}

function parseList(
  el: Element,
  kind: 'bullet' | 'ordered',
  ilvl: number,
  numIds: NumIds,
  out: PmNode[],
): void {
  el.childNodes.forEach((child) => {
    if (child.nodeType !== Node.ELEMENT_NODE) return
    const item = child as Element
    const tag = item.tagName.toLowerCase()
    if (tag === 'ul' || tag === 'ol') {
      parseList(item, tag === 'ol' ? 'ordered' : 'bullet', ilvl + 1, numIds, out)
      return
    }
    if (tag !== 'li') return
    // extract nested lists first so their items follow this one
    const nested: Element[] = []
    item.querySelectorAll(':scope > ul, :scope > ol').forEach((n) => {
      nested.push(n)
      n.remove()
    })
    out.push(
      blockNode(
        'docListItem',
        { kind, numId: numIds[kind], ilvl: Math.min(ilvl, 4) },
        parseInline(item, []),
      ),
    )
    for (const n of nested) {
      parseList(n, n.tagName.toLowerCase() === 'ol' ? 'ordered' : 'bullet', ilvl + 1, numIds, out)
    }
  })
}

/** paragraph texts of one table cell: <br> and nested <p>/<div> split paragraphs */
function cellParas(cell: Element): string[] {
  const paras: string[] = []
  let current = ''
  const push = () => {
    paras.push(current.replace(/\s+/g, ' ').trim())
    current = ''
  }
  const walk = (node: Node) => {
    node.childNodes.forEach((child) => {
      if (child.nodeType === Node.TEXT_NODE) {
        current += child.textContent ?? ''
        return
      }
      if (child.nodeType !== Node.ELEMENT_NODE) return
      const el = child as Element
      const tag = el.tagName.toLowerCase()
      if (tag === 'br') push()
      else if (tag === 'p' || tag === 'div') {
        if (current.trim()) push()
        walk(el)
        push()
      } else walk(el)
    })
  }
  walk(cell)
  if (current.trim()) push()
  const out = paras.filter((p, i) => p !== '' || (i > 0 && i < paras.length - 1))
  return out.length > 0 ? out : ['']
}

/**
 * <table> -> protected table block: display TableModel + generated w:tbl
 * fragment. Cell texts are patched into the fragment at save time via the
 * existing genXml branch of pmDocToSavePlan (patchTableCellTexts), so the
 * model row/col counts must mirror the generated grid.
 */
function parseTable(el: Element): PmNode | null {
  const trs = Array.from(el.querySelectorAll('tr'))
  if (trs.length === 0) return null
  const rawRows = trs.map((tr) => Array.from(tr.children).filter((c) => /^t[hd]$/i.test(c.tagName)))
  const cols = Math.max(...rawRows.map((r) => r.length))
  if (cols === 0) return null
  const headerRow = rawRows[0].some((c) => c.tagName.toLowerCase() === 'th')
  const rows: TableCell[][] = rawRows.map((cells, r) => {
    const isHeader = headerRow && r === 0
    const row: TableCell[] = cells.map((cell) => ({
      paras: cellParas(cell),
      ...(isHeader ? { bold: true, fill: TABLE_HEADER_FILL } : {}),
    }))
    while (row.length < cols)
      row.push({ paras: [''], ...(isHeader ? { bold: true, fill: TABLE_HEADER_FILL } : {}) })
    return row
  })
  const table: TableModel = { rows }
  return tableModelToPmNode(table)
}

/** <pre> -> mono/shaded paragraph; newlines preserved as hard breaks */
function parseCodeBlock(el: Element): PmNode | null {
  const text = (el.textContent ?? '').replace(/^\n/, '').replace(/\s+$/, '')
  if (!text) return null
  const mark: PmMark = { type: 'docTextStyle', attrs: { font: CODE_BLOCK_PRESET.font } }
  const content: PmNode[] = []
  text.split('\n').forEach((line, i) => {
    if (i > 0) content.push({ type: 'hardBreak' })
    if (line !== '') content.push({ type: 'text', text: line, marks: [mark] })
  })
  return blockNode(
    'docParagraph',
    { shadingFill: CODE_BLOCK_PRESET.shadingFill, borders: CODE_BLOCK_PRESET.borders },
    content,
  )
}

/** <blockquote> -> indented gray paragraph with a left border */
function parseBlockquote(el: Element): PmNode | null {
  const mark: PmMark = { type: 'docTextStyle', attrs: { color: QUOTE_BLOCK_PRESET.color } }
  const inline = parseInline(el, [mark])
  if (inline.length === 0) return null
  return blockNode(
    'docParagraph',
    { indentLeft: QUOTE_BLOCK_PRESET.indentLeft, borders: QUOTE_BLOCK_PRESET.borders },
    inline,
  )
}

/**
 * Parse a restricted HTML fragment returned by the model into top-level PmNodes.
 * Tolerates markdown code fences and plain-text responses.
 */
export function parseHtmlFragment(raw: string, numIds: NumIds): PmNode[] {
  let text = raw.trim()
  const fence = /```(?:html)?\s*([\s\S]*?)```/.exec(text)
  if (fence) text = fence[1].trim()
  if (!text) return []

  // plain text response (no tags): one paragraph per blank-line-separated chunk
  if (!/<[a-z][\s\S]*>/i.test(text)) {
    return text
      .split(/\n{2,}/)
      .map((para) =>
        blockNode('docParagraph', {}, [{ type: 'text', text: para.replace(/\s+/g, ' ').trim() }]),
      )
      .filter((n) => n.content?.[0]?.text)
  }

  const parsed = new DOMParser().parseFromString(text, 'text/html')
  const out: PmNode[] = []
  const pushFormula = (el: Element) => {
    const latex = (el.textContent ?? '').trim()
    if (!latex) return
    try {
      out.push(equationBlockJson(latex))
    } catch (e) {
      const reason = e instanceof Error ? e.message : String(e)
      throw new Error(`Cannot parse the LaTeX in <formula> (${reason}): ${latex}`, { cause: e })
    }
  }
  parsed.body.childNodes.forEach((child) => {
    if (child.nodeType === Node.TEXT_NODE) {
      const t = (child.textContent ?? '').trim()
      if (t) out.push(blockNode('docParagraph', {}, [{ type: 'text', text: t }]))
      return
    }
    if (child.nodeType !== Node.ELEMENT_NODE) return
    const el = child as Element
    const tag = el.tagName.toLowerCase()
    const headingMatch = /^h([1-6])$/.exec(tag)
    if (headingMatch) {
      out.push(blockNode('docHeading', { level: Number(headingMatch[1]) }, parseInline(el, [])))
    } else if (tag === 'ul' || tag === 'ol') {
      parseList(el, tag === 'ol' ? 'ordered' : 'bullet', 0, numIds, out)
    } else if (tag === 'table') {
      const node = parseTable(el)
      if (node) out.push(node)
    } else if (tag === 'pre') {
      const node = parseCodeBlock(el)
      if (node) out.push(node)
    } else if (tag === 'blockquote') {
      const node = parseBlockquote(el)
      if (node) out.push(node)
    } else if (tag === 'formula') {
      pushFormula(el)
    } else if (tag === 'p' || tag === 'div') {
      const inline = parseInline(el, [])
      if (inline.length > 0) out.push(blockNode('docParagraph', {}, inline))
    } else {
      // unknown block-ish tag: salvage the text
      const t = (el.textContent ?? '').trim()
      if (t) out.push(blockNode('docParagraph', {}, [{ type: 'text', text: t }]))
    }
  })
  return out
}

// ---- applying parsed fragments ----

/** record the AI's content change as tracked revisions under this author */
export interface AiTrack {
  author: string
}

function revisionDate(): string {
  return new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
}

/** blocks whose content ins/del marks can fully represent (tracked replace) */
const TRACKABLE_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

/**
 * Replace the top-level block range with the parsed nodes (marked aiChanged).
 * With `track`, and when both sides are plain text blocks, this becomes a
 * tracked rewrite instead: the old blocks stay struck through
 * (del) and the new blocks follow with ins marks — accept/reject via Review.
 */
export function replaceBlockRange(
  editor: Editor,
  startIndex: number,
  endIndex: number,
  nodes: PmNode[],
  track?: AiTrack,
): boolean {
  if (nodes.length === 0) return false
  const { from, to } = blockRangePositions(editor, startIndex, endIndex)
  const pmNodes = nodes.map((n) => editor.schema.nodeFromJSON(n))

  let oldTrackable = true
  editor.state.doc.nodesBetween(from, to, (node, _pos, parent) => {
    if (parent === editor.state.doc && !TRACKABLE_TYPES.has(node.type.name)) oldTrackable = false
    return false
  })
  const newTrackable = nodes.every((n) => TRACKABLE_TYPES.has(n.type))
  if (track && oldTrackable && newTrackable) {
    const { ins, del } = editor.schema.marks
    const date = revisionDate()
    // revision marks are the change indicator; no yellow aiChanged on top
    const tracked = nodes.map((n) =>
      editor.schema.nodeFromJSON({ ...n, attrs: { ...n.attrs, aiChanged: false } }),
    )
    const inserted = tracked.reduce((size, n) => size + n.nodeSize, 0)
    const tr = editor.state.tr
    tr.setMeta(TRACK_IGNORE, true)
    if (editor.state.doc.textBetween(from, to, '\n').trim() === '') {
      // nothing to strike through (blank paragraphs): replace outright
      tr.replaceWith(from, to, tracked)
      tr.addMark(from, from + inserted, ins.create({ author: track.author, date }))
    } else {
      tr.insert(to, tracked)
      tr.addMark(to, to + inserted, ins.create({ author: track.author, date }))
      tr.addMark(from, to, del.create({ author: track.author, date }))
    }
    editor.view.dispatch(tr)
    return true
  }
  if (track) {
    const date = revisionDate()
    const oldNodes: ProseMirrorNode[] = []
    editor.state.doc.nodesBetween(from, to, (node, _pos, parent) => {
      if (parent === editor.state.doc) {
        oldNodes.push(
          node.type.create(
            {
              ...node.attrs,
              aiChanged: false,
              blockRevision: { kind: 'del', author: track.author, date },
            },
            node.content,
            node.marks,
          ),
        )
      }
      return false
    })
    const inserted = pmNodes.map((node) =>
      node.type.create(
        {
          ...node.attrs,
          aiChanged: false,
          blockRevision: { kind: 'ins', author: track.author, date },
        },
        node.content,
        node.marks,
      ),
    )
    const tr = editor.state.tr.replaceWith(from, to, [...oldNodes, ...inserted])
    tr.setMeta(TRACK_IGNORE, true)
    editor.view.dispatch(tr)
    return true
  }
  editor.view.dispatch(editor.state.tr.replaceWith(from, to, pmNodes))
  return true
}

/** insert the parsed nodes after the given top-level block index */
export function insertBlocksAfter(
  editor: Editor,
  index: number,
  nodes: PmNode[],
  track?: AiTrack,
): boolean {
  if (nodes.length === 0) return false
  const { to } = blockRangePositions(editor, index, index)
  const pmNodes = nodes.map((n) =>
    editor.schema.nodeFromJSON(track ? { ...n, attrs: { ...n.attrs, aiChanged: false } } : n),
  )
  const tr = editor.state.tr.insert(to, pmNodes)
  if (track) {
    const date = revisionDate()
    let offset = to
    for (const node of pmNodes) {
      if (TRACKABLE_TYPES.has(node.type.name)) {
        tr.addMark(
          offset,
          offset + node.nodeSize,
          editor.schema.marks.ins.create({ author: track.author, date }),
        )
      } else {
        tr.setNodeMarkup(offset, undefined, {
          ...node.attrs,
          blockRevision: { kind: 'ins', author: track.author, date },
        })
      }
      offset += node.nodeSize
    }
    tr.setMeta(TRACK_IGNORE, true)
  }
  editor.view.dispatch(tr)
  return true
}
