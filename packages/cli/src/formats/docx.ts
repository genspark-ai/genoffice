import { clipText } from '../preview'
import type { Editor } from '@tiptap/core'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  findChartWorkbookPath,
  nextNoteId,
  parseChartPartXml,
  parseDocx,
  patchChartPartXml,
  patchChartWorkbookXlsxBase64,
  pendingHeadingLevel,
  readDocxPartBase64,
  readSections,
  saveDocx,
  type CommentInfo,
  type HeaderFooter,
  type HfPartInfo,
  type NoteInfo,
  type SaveOptions,
  type StyleUpsert,
  type PictureWatermarkSpec,
  type WatermarkSpec,
} from '@genoffice/docx-engine'
import type { FloatSpec } from '../../../../apps/docs/src/renderer/ai/floating-ops'
import type { AiNotesAccess, NoteKind } from '../../../../apps/docs/src/renderer/ai/note-ops'
import type { AiCommentsAccess } from '../../../../apps/docs/src/renderer/ai/tools'
import { ensureDom } from '../dom'
import type { PathContext } from '../fs'
import type { BatchMode } from '../batch'
import { classifyOpError, opSuggestion, type OpFailure } from '../op-errors'
import {
  signatureFromSchema,
  withFingerprint,
  type JsonSchema,
  type OpCatalog,
  type OpEntry,
} from '../op-catalog'
import { CliError, EXIT } from '../result'
import { readImageSource } from './image-source'
import { imageSize } from './image-size'
import { applySectionEdits, pageSetupAccess } from './docx-sections'

/**
 * The Word editing surface lives in the docs renderer (Tiptap document,
 * restricted-HTML parser, op executor, save plan). Those modules are pure
 * apart from needing a DOM, so genoffice runs them under jsdom. They are loaded
 * lazily, after the DOM exists, and by relative path until they move into a
 * package of their own.
 */
async function docsModules() {
  await ensureDom()
  const [
    { Editor },
    extensions,
    convert,
    protocol,
    ops,
    tools,
    locale,
    comments,
    hfText,
    revisions,
    pageSetup,
    revisionOps,
    floating,
    noteOps,
    tableOps,
  ] = await Promise.all([
    import('@tiptap/core'),
    import('../../../../apps/docs/src/renderer/editor/extensions'),
    import('../../../../apps/docs/src/renderer/editor/convert'),
    import('../../../../apps/docs/src/renderer/ai/protocol'),
    import('../../../../apps/docs/src/renderer/ai/ops'),
    import('../../../../apps/docs/src/renderer/ai/tools'),
    import('../../../../apps/docs/src/renderer/i18n/locale'),
    import('../../../../apps/docs/src/renderer/editor/comments'),
    import('../../../../apps/docs/src/renderer/editor/hf-text'),
    import('../../../../apps/docs/src/renderer/editor/revisions'),
    import('../../../../apps/docs/src/renderer/ai/page-setup'),
    import('../../../../apps/docs/src/renderer/ai/revision-ops'),
    import('../../../../apps/docs/src/renderer/ai/floating-ops'),
    import('../../../../apps/docs/src/renderer/ai/note-ops'),
    import('../../../../apps/docs/src/renderer/ai/table-ops'),
  ])
  locale.setModuleLang('en')
  return {
    Editor,
    extensions,
    convert,
    protocol,
    ops,
    tools,
    comments,
    hfText,
    revisions,
    pageSetup,
    revisionOps,
    floating,
    noteOps,
    tableOps,
  }
}

type Parsed = Awaited<ReturnType<typeof parseDocx>>
type Modules = Awaited<ReturnType<typeof docsModules>>

type HfView = 'default' | 'first' | 'even'
type HfSlot = `${'header' | 'footer'}${'' | 'First' | 'Even'}`

/** Header/footer parts, comments and notes live outside the ProseMirror document; edits are kept here until save. */
interface SideState {
  hf: Partial<Record<HfSlot, HeaderFooter | null>>
  hfDirty: Set<HfSlot>
  titlePg: boolean
  evenOddHf: boolean
  titlePgDirty: boolean
  evenOddHfDirty: boolean
  comments: CommentInfo[]
  commentsDirty: boolean
  /** pending sectPr rewrites by the docxIndex of the block carrying them (section breaks, trailing) */
  sectPr: Map<number, string>
  /** define_style results by styleId, written into styles.xml on save */
  styleUpserts: Map<string, StyleUpsert>
  /** set_watermark result; undefined = header untouched */
  watermark?: WatermarkSpec | PictureWatermarkSpec | null
  watermarkText: string | null
  footnotes: NoteInfo[]
  endnotes: NoteInfo[]
  notesDirty: boolean
}

export interface OpenDocument {
  parsed: Parsed
  editor: Editor
  numIds: { bullet: string | null; ordered: string | null }
  mods: Modules
  side: SideState
}

function hfFromPart(part: HfPartInfo | null | undefined): HeaderFooter | null {
  if (
    !part ||
    (!part.text && !part.hasPageNumber && part.paras.length === 0 && !part.images?.length)
  )
    return null
  return {
    text: part.text,
    pageNumber: part.hasPageNumber,
    paras: part.paras.length > 0 ? part.paras : undefined,
  }
}

function sideStateOf(parsed: Parsed): SideState {
  const dflt = (kind: 'header' | 'footer'): HeaderFooter | null => {
    const text = kind === 'header' ? parsed.headerText : parsed.footerText
    const pageNumber = kind === 'header' ? parsed.headerHasPageNumber : parsed.footerHasPageNumber
    const paras = kind === 'header' ? parsed.headerParas : parsed.footerParas
    return text || pageNumber || paras?.length
      ? { text: text ?? '', pageNumber, paras: paras ?? undefined }
      : null
  }
  return {
    hf: {
      header: dflt('header'),
      footer: dflt('footer'),
      headerFirst: hfFromPart(parsed.headerFirst),
      footerFirst: hfFromPart(parsed.footerFirst),
      headerEven: hfFromPart(parsed.headerEven),
      footerEven: hfFromPart(parsed.footerEven),
    },
    hfDirty: new Set(),
    titlePg: parsed.titlePg ?? false,
    evenOddHf: parsed.evenAndOddHeaders ?? false,
    titlePgDirty: false,
    evenOddHfDirty: false,
    comments: [...parsed.comments],
    commentsDirty: false,
    sectPr: new Map(),
    styleUpserts: new Map(),
    watermarkText: parsed.watermarkText ?? null,
    footnotes: [...parsed.footnotes],
    endnotes: [...parsed.endnotes],
    notesDirty: false,
  }
}

export async function openDocument(bytes: Uint8Array): Promise<OpenDocument> {
  const mods = await docsModules()
  const parsed = await parseDocx(bytes)
  const editor = new mods.Editor({
    element: document.createElement('div'),
    extensions: mods.extensions.editorExtensions,
  })
  editor.commands.setContent(
    mods.convert.blocksToPmDoc(parsed.blocks, readSections(parsed)) as never,
  )
  editor.storage.listNumbering.styles = parsed.styles
  const numIds = {
    bullet: mods.protocol.findNumId(parsed.blocks, 'bullet') ?? BLANK_BULLET_NUM_ID,
    ordered: mods.protocol.findNumId(parsed.blocks, 'ordered') ?? BLANK_ORDERED_NUM_ID,
  }
  return { parsed, editor, numIds, mods, side: sideStateOf(parsed) }
}

export async function blankDocument(): Promise<OpenDocument> {
  return openDocument(await buildBlankDocx())
}

export async function saveDocument(doc: OpenDocument): Promise<Uint8Array> {
  const plan = doc.mods.convert.pmDocToSavePlan(doc.editor.getJSON() as never, doc.parsed.blocks)
  const { side } = doc
  const hf = (slot: HfSlot) => (side.hfDirty.has(slot) ? (side.hf[slot] ?? undefined) : undefined)
  const sectionEdits = applySectionEdits(doc, plan)
  const options: SaveOptions = {
    ...sectionEdits.options,
    header: hf('header'),
    footer: hf('footer'),
    headerFirst: hf('headerFirst'),
    footerFirst: hf('footerFirst'),
    headerEven: hf('headerEven'),
    footerEven: hf('footerEven'),
    titlePg: side.titlePgDirty ? side.titlePg : undefined,
    evenAndOddHeaders: side.evenOddHfDirty ? side.evenOddHf : undefined,
    comments: side.commentsDirty ? side.comments : undefined,
    styleUpserts: side.styleUpserts.size > 0 ? [...side.styleUpserts.values()] : undefined,
    watermark: side.watermark,
    footnotes: side.notesDirty ? side.footnotes : undefined,
    endnotes: side.notesDirty ? side.endnotes : undefined,
    ...(await chartPartPatches(doc, plan.chartPatches)),
  }
  return saveDocx(doc.parsed, sectionEdits.saveBlocks, options)
}

/** edit_chart changes live in the chart's own part and its embedded workbook, not in the body XML */
async function chartPartPatches(
  doc: OpenDocument,
  patches: ReturnType<Modules['convert']['pmDocToSavePlan']>['chartPatches'],
): Promise<Pick<SaveOptions, 'partXml' | 'partBinary'>> {
  const partXml: Record<string, string> = {}
  const partBinary: Record<string, string> = {}
  const original = doc.parsed.internal.originalBytes
  for (const { partPath, patch } of patches) {
    const part = doc.parsed.extras.chartParts[partPath]
    if (!part) continue
    const patched = patchChartPartXml(part, patch)
    partXml[partPath] = patched
    const wbPath = await findChartWorkbookPath(original, partPath)
    const wb = wbPath ? await readDocxPartBase64(original, wbPath) : null
    const display = wb ? parseChartPartXml(patched, partPath) : null
    if (wbPath && wb && display) {
      const updated = await patchChartWorkbookXlsxBase64(
        wb,
        display.categories,
        display.series.map((s, i) => ({
          name: s.name ?? `Series${i + 1}`,
          values: s.values as (number | null)[],
        })),
      )
      if (updated) partBinary[wbPath] = updated
    }
  }
  return {
    ...(Object.keys(partXml).length ? { partXml } : {}),
    ...(Object.keys(partBinary).length ? { partBinary } : {}),
  }
}

export function closeDocument(doc: OpenDocument): void {
  doc.editor.destroy()
}

export interface BlockSummary {
  index: number
  type: string
  level?: number
  /** protected blocks: image, chart, field, formula … (targets for setImageProperties / edit_chart) */
  kind?: string
  text: string
  /** native tables: grid size, merged cells and the table style, for the table ops */
  table?: ReturnType<Modules['tableOps']['describeTable']>
  /** the text was clipped to the preview length; `--full` or `--max-chars` returns the rest */
  truncated?: true
}

export const PREVIEW_CHARS = 200

/** Block list an agent targets ops at: index, node type, heading level, text (clipped to `maxChars`). */
export function describeDocument(
  doc: OpenDocument,
  range?: [number, number],
  maxChars = PREVIEW_CHARS,
): BlockSummary[] {
  const out: BlockSummary[] = []
  const root = doc.editor.state.doc
  const [start, end] = range ?? [0, root.childCount - 1]
  for (let i = start; i <= Math.min(end, root.childCount - 1); i++) {
    const node = root.child(i)
    const { text, truncated } = clipText(node.textContent, maxChars)
    out.push({
      index: i,
      type: node.type.name.replace(/^doc/, '').toLowerCase(),
      ...(typeof node.attrs.level === 'number' ? { level: node.attrs.level } : {}),
      ...(node.type.name === 'docProtected' ? { kind: protectedKind(node.attrs) } : {}),
      ...(node.type.name === 'docTable' ? { table: doc.mods.tableOps.describeTable(node) } : {}),
      text,
      ...(truncated ? { truncated } : {}),
    })
  }
  return out
}

/** table styles defined in styles.xml (setTableStyle accepts the id or the name) */
export function listTableStyles(doc: OpenDocument): Array<{ id: string; name: string }> {
  return [...doc.parsed.styles.values()]
    .filter((s) => s.type === 'table')
    .map((s) => ({ id: s.styleId, name: s.name }))
}

/** native docx charts are passthrough blocks carrying chartDisplay; edit_chart targets them like AI-made ones */
function protectedKind(attrs: Record<string, unknown>): string {
  if (attrs.chartDisplay) return 'chart'
  return typeof attrs.blockType === 'string' ? attrs.blockType : 'protected'
}

export function blockRangeHtml(doc: OpenDocument, start: number, end: number): string {
  return doc.mods.protocol.serializeRangeToHtml(doc.editor, start, end)
}

export interface ExportHtml {
  html: string
  /** blocks Markdown cannot carry, by kind */
  skipped: { images: number; fields: number }
  /** formulas whose LaTeX was not recoverable; emitted as their token text */
  formulasAsText: number
}

/**
 * The whole document as HTML for a human-facing export. The AI protocol's
 * serializer describes protected blocks (images, fields, formulas without
 * LaTeX) with "[Protected …]" notes meant for the model; here they become a
 * math node, the field's visible text, or nothing, and inline <formula> tags
 * take the markdown editor's math markup.
 */
export function documentHtml(doc: OpenDocument): ExportHtml {
  const top = doc.editor.state.doc
  const parts: string[] = []
  const skipped = { images: 0, fields: 0 }
  let formulasAsText = 0
  let runStart = -1
  const flush = (end: number) => {
    if (runStart >= 0) parts.push(blockRangeHtml(doc, runStart, end))
    runStart = -1
  }
  top.forEach((node, _offset, index) => {
    if (node.type.name !== 'docProtected') {
      if (runStart < 0) runStart = index
      return
    }
    flush(index - 1)
    // pending tracked deletions are dropped, as serializeRangeToHtml drops them
    if (doc.mods.protocol.isTrackedDeleted(node)) return
    const formula = node.attrs.formulaDisplay as { tokens?: string[]; latex?: string } | null
    if (formula?.latex) {
      parts.push(`<div data-type="block-math" data-latex="${escapeHtml(formula.latex)}"></div>`)
    } else if (formula?.tokens?.length) {
      formulasAsText++
      parts.push(`<p>${escapeHtml(formula.tokens.join(' '))}</p>`)
    } else if (node.attrs.blockType === 'image') {
      skipped.images++
    } else {
      const preview = String(node.attrs.previewText ?? '')
        .replace(/\s+/g, ' ')
        .trim()
      if (preview) parts.push(`<p>${escapeHtml(preview)}</p>`)
      else skipped.fields++
    }
  })
  flush(top.childCount - 1)
  const html = parts
    .join('\n')
    // the protocol escapes & < > in formula text but not quotes, which would end the attribute
    .replace(
      /<formula>([\s\S]*?)<\/formula>/g,
      (_m, latex: string) =>
        `<span data-type="inline-math" data-latex="${latex.replace(/"/g, '&quot;')}"></span>`,
    )
  return { html, skipped, formulasAsText }
}

function escapeHtml(s: string): string {
  return s.replace(
    /[&<>"]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!,
  )
}

/** Replaces the whole body with a restricted-HTML fragment (the create_document contract). */
export function fillFromHtml(doc: OpenDocument, html: string): number {
  const nodes = doc.mods.protocol.parseHtmlFragment(html, doc.numIds)
  if (nodes.length === 0) {
    throw new CliError(EXIT.usage, 'no content could be parsed from the HTML', undefined, {
      reason: 'invalid_argument',
    })
  }
  const last = doc.editor.state.doc.childCount - 1
  doc.mods.protocol.replaceBlockRange(doc.editor, 0, last, nodes)
  return nodes.length
}

export interface DocOpResult {
  index: number
  op: string
  output: string
}

export interface DocOpsOutcome {
  results: DocOpResult[]
  failures: OpFailure[]
  /** ops after the first rejection under stop_on_error, never attempted */
  skipped: number
}

const HTML_TOOLS = new Set(['insert_content', 'replace_blocks'])
/** app tools reachable headless besides the html pair; insert_image is built here (no renderer image bridge) */
const SIDE_TOOLS = new Set([
  'insert_chart',
  'edit_chart',
  'set_header_footer',
  'set_page_setup',
  'insert_section_break',
  'reply_comment',
  'resolve_comment',
  'add_comment',
  'delete_comment',
  'read_comments',
  'read_revisions',
  'accept_changes',
  'reject_changes',
  'define_style',
  'list_styles',
  'set_watermark',
  'insert_text_box',
  'insert_footnote',
  'insert_endnote',
  'delete_note',
  'edit_note',
  'read_notes',
])

export interface ApplyOptions {
  ctx?: PathContext
  /** folder of the ops file, for relative image paths */
  baseDir?: string
  /** record content edits as tracked changes by this author */
  trackAuthor?: string
  mode?: BatchMode
}

/**
 * Every entry goes through the app's executeTool: apply_ops entries as an
 * `apply_ops` call, the html and side tools as themselves. That is what keeps
 * the tool layer's stale-document guard in step (it records the document it
 * last touched), so mixed batches behave like one in-app turn.
 *
 * A dry run executes the batch on the in-memory editor and simply never
 * saves: that is the only way index shifts from earlier entries are taken
 * into account when later ones are validated.
 */
export async function applyDocOps(
  doc: OpenDocument,
  ops: Record<string, unknown>[],
  opts: ApplyOptions = {},
): Promise<DocOpsOutcome> {
  const mode = opts.mode ?? 'atomic'
  const results: DocOpResult[] = []
  const failures: OpFailure[] = []
  const comments = commentsAccess(doc)
  const hf = headerFooterAccess(doc)
  const pageSetup = pageSetupAccess(doc)
  const extras = docExtras(doc)
  const notes = notesAccess(doc)
  for (const [index, op] of ops.entries()) {
    const name = typeof op.op === 'string' ? op.op : ''
    // a rejected op left the document as it was: the executor applies each call whole or not at all
    const rejected = (output: string): boolean => {
      const failure = classifyOpError(index, name, output)
      if (mode === 'atomic') {
        throw new CliError(
          EXIT.usage,
          `op ${index} (${name || '?'}) rejected: ${output}`,
          { failures: [failure] },
          { reason: failure.reason, suggestion: opSuggestion(failure, 'docs') },
        )
      }
      failures.push(failure)
      return true
    }
    let failed = false
    if (name === 'insert_image' || name === 'insert_picture') {
      const r = await insertImage(doc, op, opts, name === 'insert_picture')
      if (r.error) failed = rejected(r.error)
      else results.push({ index, op: name, output: r.output! })
    } else if (name === 'set_watermark' && typeof op.image === 'string') {
      const r = await setPictureWatermark(doc, op, opts)
      if (r.error) failed = rejected(r.error)
      else results.push({ index, op: name, output: r.output! })
    } else {
      const call = HTML_TOOLS.has(name)
        ? { id: `genoffice-${index}`, name, input: htmlToolInput(doc, name, op) }
        : SIDE_TOOLS.has(name)
          ? { id: `genoffice-${index}`, name, input: sideToolInput(doc, name, op) }
          : { id: `genoffice-${index}`, name: 'apply_ops', input: { ops: [op] } }
      const exec = await doc.mods.tools.executeTool(
        doc.editor,
        call,
        doc.numIds,
        opts.trackAuthor ? { author: opts.trackAuthor } : undefined,
        undefined,
        null,
        comments,
        hf,
        undefined,
        pageSetup,
        extras,
        notes,
      )
      // the in-app executor reports a target that matched nothing as a soft result; for a
      // scripted batch that is a wrong index, so the op counts as rejected
      if (exec.isError || /^No matching blocks/i.test(exec.output)) failed = rejected(exec.output)
      else results.push({ index, op: name, output: exec.output })
    }
    if (failed && mode === 'stop_on_error') {
      return { results, failures, skipped: ops.length - index - 1 }
    }
  }
  return { results, failures, skipped: 0 }
}

/** insert_content / insert_chart without a position append at the end; the in-app default ("after the cursor") has no meaning headless. */
function htmlToolInput(
  doc: OpenDocument,
  name: string,
  op: Record<string, unknown>,
): Record<string, unknown> {
  const { op: _op, ...input } = op
  if (name === 'insert_content' && input.afterBlockIndex === undefined) {
    input.afterBlockIndex = doc.editor.state.doc.childCount - 1
  }
  return input
}

function sideToolInput(
  doc: OpenDocument,
  name: string,
  op: Record<string, unknown>,
): Record<string, unknown> {
  const { op: _op, ...input } = op
  if (name === 'insert_chart' && input.afterBlockIndex === undefined) {
    input.afterBlockIndex = doc.editor.state.doc.childCount - 1
  }
  return input
}

const CLI_COMMENT_AUTHOR = 'AI Assistant'

const commentDate = (): string => new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')

function commentsAccess(doc: OpenDocument): AiCommentsAccess {
  const { side } = doc
  return {
    list: () => side.comments,
    add: (range, text, meta) => {
      const id = doc.mods.comments.nextCommentId(side.comments)
      if (!doc.mods.comments.addCommentToRange(doc.editor, range.from, range.to, id)) return null
      side.comments.push({
        id,
        author: meta.author ?? CLI_COMMENT_AUTHOR,
        ...(meta.initials ? { initials: meta.initials } : {}),
        date: commentDate(),
        text,
      })
      side.commentsDirty = true
      return id
    },
    remove: (id) => {
      const victims = new Set([
        id,
        ...side.comments.filter((c) => c.parentId === id).map((c) => c.id),
      ])
      if (!side.comments.some((c) => c.id === id)) return false
      for (const v of victims) doc.mods.comments.removeCommentFromDoc(doc.editor, v)
      side.comments = side.comments.filter((c) => !victims.has(c.id))
      side.commentsDirty = true
      return true
    },
    reply: (parentId, text) => {
      const id = doc.mods.comments.nextCommentId(side.comments)
      if (!doc.mods.comments.addReplyToCommentRange(doc.editor, parentId, id)) return false
      side.comments.push({ id, author: CLI_COMMENT_AUTHOR, date: commentDate(), text, parentId })
      side.commentsDirty = true
      return true
    },
    resolve: (id) => {
      if (!side.comments.some((c) => c.id === id)) return false
      side.comments = side.comments.map((c) =>
        c.id === id || c.parentId === id ? { ...c, done: true } : c,
      )
      side.commentsDirty = true
      return true
    },
  }
}

function notesAccess(doc: OpenDocument): AiNotesAccess {
  const { side } = doc
  const listOf = (kind: NoteKind) => (kind === 'footnote' ? side.footnotes : side.endnotes)
  const setList = (kind: NoteKind, next: NoteInfo[]) => {
    if (kind === 'footnote') side.footnotes = next
    else side.endnotes = next
    side.notesDirty = true
  }
  return {
    list: listOf,
    add: (kind, text) => {
      const id = nextNoteId(listOf(kind))
      setList(kind, [...listOf(kind), { id, text }])
      return id
    },
    remove: (kind, id) => {
      if (!listOf(kind).some((n) => n.id === id)) return false
      setList(
        kind,
        listOf(kind).filter((n) => n.id !== id),
      )
      return true
    },
    replace: (kind, id, next) => {
      if (!listOf(kind).some((n) => n.id === id)) return false
      setList(
        kind,
        listOf(kind).map((n) => (n.id === id ? next : n)),
      )
      return true
    },
    protectedMarkBlock: (kind, id) =>
      doc.mods.noteOps.protectedNoteMarkBlock(
        doc.editor.state.doc,
        (block) => protectedXml(doc, block),
        kind,
        id,
      ),
  }
}

export interface NoteSummary {
  kind: NoteKind
  id: string
  /** display number by document order of the reference marks; absent when the text has no mark */
  num?: number
  blockIndex?: number
  text: string
}

/** Footnotes and endnotes with the block holding their reference mark; ids as delete_note takes them. */
export function listNotes(doc: OpenDocument): NoteSummary[] {
  const anchors = doc.mods.noteOps.noteAnchors(doc.editor.state.doc)
  // marks inside protected blocks are only in the block XML; the parse numbered those
  const inXml = new Map<string, number>()
  doc.editor.state.doc.forEach((block, _offset, blockIndex) => {
    if (block.type.name !== 'docProtected') return
    const xml = protectedXml(doc, block)
    for (const m of xml.matchAll(/<w:(footnote|endnote)Reference\b[^>]*\bw:id="([^"]+)"/g)) {
      inXml.set(`${m[1]}:${m[2]}`, blockIndex)
    }
  })
  const parsedNumbers = doc.parsed.noteNumbers as Record<string, number> | undefined
  const summarize = (kind: NoteKind, notes: NoteInfo[]) =>
    notes.map((n) => {
      const key = `${kind}:${n.id}`
      const a = anchors.get(key)
      const xmlBlock = inXml.get(key)
      const num = a?.num ?? parsedNumbers?.[key]
      const blockIndex = a?.blockIndex ?? xmlBlock
      return {
        kind,
        id: n.id,
        ...(num !== undefined ? { num } : {}),
        ...(blockIndex !== undefined ? { blockIndex } : {}),
        text: n.text,
      }
    })
  return [...summarize('footnote', doc.side.footnotes), ...summarize('endnote', doc.side.endnotes)]
}

export interface FieldSummary {
  blockIndex: number
  /** field keyword: SEQ, REF, DATE, PAGE, TOC … */
  type: string
  /** full instruction text */
  instr: string
  /** cached result as the document shows it */
  result: string
  /** Word recomputes the field when the file opens */
  dirty: boolean
  /** the field sits in a protected block (only Word or updateFields-free rebuilds change it) */
  protected?: true
}

const keywordOf = (instr: string) => /^\s*([A-Za-z]+)/.exec(instr)?.[1]?.toUpperCase() ?? ''

/** complex (fldChar) fields at depth 1 and fldSimple fields in a block's XML */
function xmlFields(xml: string): Array<Pick<FieldSummary, 'instr' | 'result' | 'dirty'>> {
  const out: Array<Pick<FieldSummary, 'instr' | 'result' | 'dirty'>> = []
  const decode = (t: string) =>
    t
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
  let depth = 0
  let instr = ''
  let result: string | null = null
  let dirty = false
  const simple: Array<{ instr: string; dirty: boolean; result: string }> = []
  const re =
    /<w:fldChar\b([^>]*)\/?>|<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>|<w:t\b[^>]*>([^<]*)<\/w:t>|<w:fldSimple\b([^>]*)>|<\/w:fldSimple>/g
  for (const m of xml.matchAll(re)) {
    if (m[1] !== undefined) {
      const type = /w:fldCharType="(\w+)"/.exec(m[1])?.[1]
      if (type === 'begin') {
        depth++
        if (depth === 1) {
          instr = ''
          result = null
          dirty = /\bw:dirty="(?:true|1)"/.test(m[1])
        }
      } else if (type === 'separate') {
        if (depth === 1) result = ''
      } else if (type === 'end') {
        if (depth === 1)
          out.push({ instr: decode(instr).trim(), result: decode(result ?? ''), dirty })
        depth = Math.max(0, depth - 1)
      }
    } else if (m[2] !== undefined) {
      if (depth === 1 && result === null) instr += m[2]
    } else if (m[3] !== undefined) {
      if (depth === 1 && result !== null) result += m[3]
      else if (depth === 0 && simple.length) simple[simple.length - 1]!.result += m[3]
    } else if (m[4] !== undefined) {
      simple.push({
        instr: decode(/\bw:instr="([^"]*)"/.exec(m[4])?.[1] ?? '').trim(),
        dirty: /\bw:dirty="(?:true|1)"/.test(m[4]),
        result: '',
      })
    } else {
      const done = simple.pop()
      if (done) out.push({ instr: done.instr, result: decode(done.result), dirty: done.dirty })
    }
  }
  return out
}

function protectedXml(doc: OpenDocument, block: Editor['state']['doc']): string {
  if (typeof block.attrs.genXml === 'string') return block.attrs.genXml
  const docxIndex = block.attrs.docxIndex
  return (typeof docxIndex === 'number' ? doc.parsed.blocks[docxIndex]?.originalXml : null) ?? ''
}

/** Every field with its block: editable inline fields from the editor, the rest from the block XML. */
export function listFields(doc: OpenDocument): FieldSummary[] {
  const out: FieldSummary[] = []
  doc.editor.state.doc.forEach((block, _offset, blockIndex) => {
    if (block.type.name === 'docProtected') {
      for (const f of xmlFields(protectedXml(doc, block))) {
        out.push({ blockIndex, type: keywordOf(f.instr), ...f, protected: true })
      }
      return
    }
    // a result split by formatting is one field: join text nodes that touch and carry the same mark
    type FieldMark = Editor['state']['doc']['marks'][number]
    let prev: { mark: FieldMark; end: number; field: FieldSummary } | null = null
    block.descendants((node, pos) => {
      if (!node.isText) return true
      const instr = node.marks.find((m) => m.type.name === 'instrField')
      const ref = node.marks.find((m) => m.type.name === 'refField')
      const mark = instr ?? ref
      if (!mark) return true
      if (prev && prev.end === pos && prev.mark.eq(mark)) {
        prev.field.result += node.text ?? ''
        prev.end += node.nodeSize
        return true
      }
      const field: FieldSummary = instr
        ? {
            blockIndex,
            type: keywordOf(String(instr.attrs.instr)),
            instr: String(instr.attrs.instr),
            result: node.text ?? '',
            dirty: instr.attrs.dirty === true,
          }
        : {
            blockIndex,
            type: 'REF',
            instr: mark.attrs.instr
              ? String(mark.attrs.instr).trim()
              : `REF ${String(mark.attrs.name)} \\h`,
            result: node.text ?? '',
            dirty: mark.attrs.dirty === true,
          }
      out.push(field)
      prev = { mark, end: pos + node.nodeSize, field }
      return true
    })
  })
  return out
}

const slotOf = (kind: 'header' | 'footer', view: HfView): HfSlot =>
  view === 'default' ? kind : `${kind}${view === 'first' ? 'First' : 'Even'}`

export interface HeaderFooterState {
  header: string
  footer: string
  headerFirst: string | null
  footerFirst: string | null
  headerEven: string | null
  footerEven: string | null
  titlePg: boolean
  evenOddHf: boolean
  multiSection: boolean
  /** the default header's watermark: text, picture, or none; pending set_watermark results included */
  watermark: WatermarkState
}

export type WatermarkState =
  | { kind: 'none'; text: null }
  | { kind: 'text'; text: string }
  | { kind: 'picture'; text: null; widthPt: number; heightPt: number; washout: boolean }

function watermarkState(doc: OpenDocument): WatermarkState {
  const { side, parsed } = doc
  if (side.watermark !== undefined) {
    const wm = side.watermark
    if (wm === null) return { kind: 'none', text: null }
    if ('image' in wm) {
      // the frame is sized against the page on save; report the natural size until then
      const box = { widthPt: (wm.image.widthPx * 72) / 96, heightPt: (wm.image.heightPx * 72) / 96 }
      const f = wm.scale === undefined ? 1 : wm.scale / 100
      return {
        kind: 'picture',
        text: null,
        widthPt: Math.round(box.widthPt * f * 100) / 100,
        heightPt: Math.round(box.heightPt * f * 100) / 100,
        washout: wm.washout !== false,
      }
    }
    return { kind: 'text', text: wm.text }
  }
  if (side.watermarkText) return { kind: 'text', text: side.watermarkText }
  const pic = parsed.watermarkPicture
  if (pic)
    return {
      kind: 'picture',
      text: null,
      widthPt: pic.widthPt,
      heightPt: pic.heightPt,
      washout: pic.washout,
    }
  return { kind: 'none', text: null }
}

export function headerFooterState(doc: OpenDocument): HeaderFooterState {
  const { side } = doc
  const textOf = (slot: HfSlot) => {
    const v = side.hf[slot]
    return v ? doc.mods.hfText.hfEditText(v) : ''
  }
  return {
    header: textOf('header'),
    footer: textOf('footer'),
    headerFirst: side.titlePg ? textOf('headerFirst') : null,
    footerFirst: side.titlePg ? textOf('footerFirst') : null,
    headerEven: side.evenOddHf ? textOf('headerEven') : null,
    footerEven: side.evenOddHf ? textOf('footerEven') : null,
    titlePg: side.titlePg,
    evenOddHf: side.evenOddHf,
    multiSection: readSections(doc.parsed).length > 1,
    watermark: watermarkState(doc),
  }
}

/** The default variant is written to the trailing section's part, as the app does for single-section documents. */
function headerFooterAccess(doc: OpenDocument): {
  read(): HeaderFooterState
  set(kind: 'header' | 'footer', view: HfView, text: string): string | null
} {
  const { side } = doc
  return {
    read: () => headerFooterState(doc),
    set: (kind, view, text) => {
      if (view === 'first' && !side.titlePg) {
        side.titlePg = true
        side.titlePgDirty = true
      }
      if (view === 'even' && !side.evenOddHf) {
        side.evenOddHf = true
        side.evenOddHfDirty = true
      }
      const slot = slotOf(kind, view)
      side.hf[slot] = doc.mods.hfText.applyHfText(side.hf[slot] ?? null, text)
      side.hfDirty.add(slot)
      return null
    },
  }
}

const IMAGE_MIMES = new Set(['image/png', 'image/jpeg', 'image/gif'])

/**
 * The app's insert_image / insert_picture download through the renderer
 * bridge and measure with an <img>; neither exists here, so the node is built
 * from the bytes directly, then the executor is told the document is still
 * the one it saw. `picture` = the sized / floating insert_picture variant.
 */
async function insertImage(
  doc: OpenDocument,
  op: Record<string, unknown>,
  opts: ApplyOptions,
  picture = false,
): Promise<{ output?: string; error?: string }> {
  const url = typeof op.url === 'string' ? op.url.trim() : ''
  if (!url) return { error: 'url must be a local path, a data: URL or an http(s) URL' }
  if (!opts.ctx) return { error: 'image sources need a path context' }
  const source = await readImageSource(url, opts.ctx, opts.baseDir)
  if (!source) return { error: `image not found or not downloadable: ${url}` }
  const size = imageSize(source.bytes)
  if (!size || !source.mime || !IMAGE_MIMES.has(source.mime)) {
    return { error: 'unsupported image format (only png, jpg and gif can be embedded)' }
  }
  const at = doc.mods.floating.insertPosition(doc.editor, op.afterBlockIndex)
  if ('error' in at) return { error: at.error }
  const base64 = Buffer.from(source.bytes).toString('base64')
  const mime = source.mime as 'image/png' | 'image/jpeg' | 'image/gif'
  let float: FloatSpec | undefined
  if (picture && op.float !== undefined) {
    const r = doc.mods.floating.resolveFloat(op.float)
    if ('error' in r) return { error: r.error }
    float = r.float
  }
  const built = doc.mods.floating.pictureNode({
    base64,
    mime,
    naturalWidth: size.width,
    naturalHeight: size.height,
    label: picture ? 'Picture' : 'Image',
    ...(picture
      ? {
          width: op.width,
          height: op.height,
          float,
          ...(typeof op.altText === 'string' ? { altText: op.altText } : {}),
        }
      : { width: `${Math.min(size.width, Number(op.maxWidthPx) || 480)}px` }),
  })
  if ('error' in built) return { error: built.error }
  doc.editor.chain().insertContentAt(at.pos, built.node).run()
  doc.mods.tools.markDocSeen(doc.editor)
  return {
    output: `Inserted the ${float ? 'floating ' : ''}${picture ? 'picture' : 'image'} (${built.widthPx}x${built.heightPx}px) after block ${at.after}.`,
  }
}

/** set_watermark with an image: the app downloads through the renderer bridge; here the bytes come from a path / data: / http(s) URL. */
async function setPictureWatermark(
  doc: OpenDocument,
  op: Record<string, unknown>,
  opts: ApplyOptions,
): Promise<{ output?: string; error?: string }> {
  const url = String(op.image).trim()
  if (!url) return { error: 'image must be a local path, a data: URL or an http(s) URL' }
  if (!opts.ctx) return { error: 'image sources need a path context' }
  const source = await readImageSource(url, opts.ctx, opts.baseDir)
  if (!source) return { error: `image not found or not downloadable: ${url}` }
  const size = imageSize(source.bytes)
  if (!size || !source.mime || !IMAGE_MIMES.has(source.mime)) {
    return { error: 'unsupported image format (only png, jpg and gif can be embedded)' }
  }
  const resolved = doc.mods.floating.resolvePictureWatermark(op, {
    base64: Buffer.from(source.bytes).toString('base64'),
    mime: source.mime as 'image/png' | 'image/jpeg' | 'image/gif',
    widthPx: size.width,
    heightPx: size.height,
  })
  if ('error' in resolved) return { error: resolved.error }
  doc.side.watermark = resolved.spec
  return { output: `Picture watermark set (${size.width}x${size.height}px source).` }
}

type DocExtras = NonNullable<Parameters<Modules['tools']['executeTool']>[10]>

/** styles.xml catalog plus this batch's define_style entries, and the header watermark. */
function docExtras(doc: OpenDocument): DocExtras {
  return {
    styles: {
      list: () => listStyles(doc),
      upsert: (up) => {
        const prev = doc.side.styleUpserts.get(up.styleId)
        doc.side.styleUpserts.set(
          up.styleId,
          prev
            ? {
                ...prev,
                ...up,
                pPr: up.pPr || prev.pPr ? { ...prev.pPr, ...up.pPr } : undefined,
                rPr: up.rPr || prev.rPr ? { ...prev.rPr, ...up.rPr } : undefined,
              }
            : up,
        )
        return null
      },
    },
    watermark: {
      current: () => watermarkState(doc).text,
      set: (spec) => {
        doc.side.watermark = spec
        return null
      },
    },
  }
}

export interface StyleSummary {
  styleId: string
  name: string
  type: 'paragraph' | 'character' | 'table'
  basedOn?: string
  headingLevel?: number
  /** blocks (paragraphs, headings, list items, runs) carrying the style */
  inUse: number
  /** defined or changed by this batch, not saved yet */
  pending?: boolean
}

/** Every style the document defines (plus pending define_style entries), with usage counts. */
export function listStyles(doc: OpenDocument): StyleSummary[] {
  const counts = new Map<string, number>()
  const headings = new Map<number, number>()
  const bump = (id: string) => counts.set(id, (counts.get(id) ?? 0) + 1)
  doc.editor.state.doc.descendants((node) => {
    if (typeof node.attrs?.styleId === 'string' && node.attrs.styleId) bump(node.attrs.styleId)
    else if (node.type.name === 'docHeading') {
      const level = Number(node.attrs.level) || 1
      headings.set(level, (headings.get(level) ?? 0) + 1)
    }
    for (const mark of node.marks) {
      if (typeof mark.attrs?.styleId === 'string' && mark.attrs.styleId) bump(mark.attrs.styleId)
    }
    return true
  })
  const out = new Map<string, StyleSummary>()
  for (const s of doc.parsed.styles.values()) {
    if (s.linkedCharShell) continue
    let inUse = counts.get(s.styleId) ?? 0
    if (s.headingLevel && /^Heading[1-9]$/.test(s.styleId))
      inUse += headings.get(s.headingLevel) ?? 0
    out.set(s.styleId, {
      styleId: s.styleId,
      name: s.name,
      type: s.type,
      ...(s.basedOn ? { basedOn: s.basedOn } : {}),
      ...(s.headingLevel ? { headingLevel: s.headingLevel } : {}),
      inUse,
    })
  }
  const parsed = (id: string) => doc.parsed.styles.get(id)
  for (const up of doc.side.styleUpserts.values()) {
    const cur = out.get(up.styleId)
    const basedOn = up.basedOn === undefined ? cur?.basedOn : (up.basedOn ?? undefined)
    const headingLevel = pendingHeadingLevel(
      up.styleId,
      (id) => doc.side.styleUpserts.get(id),
      parsed,
    )
    out.set(up.styleId, {
      styleId: up.styleId,
      name: up.name ?? cur?.name ?? up.styleId,
      type: cur?.type ?? up.type ?? 'paragraph',
      ...(basedOn ? { basedOn } : {}),
      ...(headingLevel ? { headingLevel } : {}),
      inUse: cur?.inUse ?? counts.get(up.styleId) ?? 0,
      pending: true,
    })
  }
  return [...out.values()]
}

export interface CommentSummary extends CommentInfo {
  blockIndex?: number
  anchorText?: string
}

/** Every comment with the block its anchor sits in, ids as reply_comment / resolve_comment / delete_comment take them. */
export function listComments(doc: OpenDocument): CommentSummary[] {
  const anchors = doc.mods.protocol.commentAnchors(doc.editor)
  return doc.side.comments.map((c) => {
    const a = anchors.get(c.parentId ?? c.id)
    const base = { ...c, done: c.done === true }
    return a ? { ...base, blockIndex: a.blockIndex, anchorText: a.excerpt } : base
  })
}

export interface RevisionSummary {
  /** positional: r1 = first pending change in document order; renumbered after every edit */
  id: string
  type: 'insertion' | 'deletion' | 'formatting' | 'move'
  kind: string
  author: string
  date?: string
  blockIndex: number
  text: string
  /** formatting changes: `field: old → new` pairs */
  change?: string
}

/** Pending tracked changes in document order, ids as accept_changes / reject_changes take them. */
export function listRevisions(doc: OpenDocument): RevisionSummary[] {
  return doc.mods.revisionOps
    .listRevisionEntries(doc.editor.state.doc)
    .map(({ from: _from, to: _to, ...rest }) => rest)
}

const CONTENT_OPS = [
  'insert_content',
  'replace_blocks',
  'insert_image',
  'insert_picture',
  'insert_text_box',
  'insert_chart',
  'edit_chart',
  'set_header_footer',
  'set_page_setup',
  'insert_section_break',
  'set_watermark',
  'define_style',
  'add_comment',
  'reply_comment',
  'resolve_comment',
  'accept_changes',
  'reject_changes',
  'insert_footnote',
  'insert_endnote',
  'edit_note',
  'delete_note',
  'delete_comment',
] as const

const CONTENT_NOTES: Record<(typeof CONTENT_OPS)[number], string> = {
  insert_content: 'new blocks after an index (-1 = start of document; omitted = end of document)',
  replace_blocks:
    "rewrite a block range; the new blocks inherit the replaced blocks' paragraph and text formatting",
  insert_image:
    'url is a local path (absolute, or relative to the current directory then to the ops file), a data: URL or an http(s) URL; png/jpg/gif; scaled down to maxWidthPx; appended when afterBlockIndex is omitted',
  insert_picture:
    'sized / floating variant of insert_image (same url forms); width/height/float offsets take "2.54cm", "1in", "72pt", "96px" or points; float.anchor paragraph|page|margin, float.wrap square|tight|topAndBottom|behind|inFront',
  insert_text_box:
    'floating text box anchored after a block; x/y/width/height in cm/in/pt/px; anchor paragraph (default) or page; wrap topAndBottom (text resumes below) or none',
  set_watermark:
    'watermark behind every page, written into the page header (Word Design > Watermark): gray text (color "#RRGGBB", opacity 0-1, diagonal true/false) or a picture (image = local path / data: / http(s) URL; scale percent, default fits the margins; washout true/false); text null or image null removes it',
  define_style:
    'create or patch a style in styles.xml; only the given fields change; put it on paragraphs with the block op { op: "applyStyle", target, styleId }; `genoffice docs read --styles` lists ids',
  insert_chart:
    'native Word chart; values per series match the categories; appended when afterBlockIndex is omitted',
  edit_chart:
    'change the data of a chart block (`genoffice docs read` lists blocks with kind "chart"); counts must match the original, null keeps a position',
  set_header_footer:
    'plain text, \\n between lines, {PAGE} / {NUMPAGES} become page-number fields, "" clears; view first/even switches the different-first-page / odd-even setting on; `genoffice docs read --header-footer` shows the current text',
  set_page_setup:
    'paper (A3/A4/A5/B5/Letter/Legal/Tabloid) or width/height, orientation, margins {top,right,bottom,left,header,footer,gutter}, columns {count,spacing}, titlePg, pageNumberStart/pageNumberFormat; lengths as "2.54cm" / "1in" / "72pt" or twips; section or blockIndex picks one section, neither = all; `genoffice docs read --sections` lists them',
  insert_section_break:
    'the blocks after afterBlockIndex become a new section (a copy of the current setup, then change it with set_page_setup section: N); type nextPage (default) / continuous / evenPage / oddPage',
  add_comment:
    'new thread on a block, or on an exact text span inside it (occurrence picks one of several matches); the document text is untouched; author defaults to "AI Assistant"',
  reply_comment: 'ids from `genoffice docs read --comments`; replies attach to the thread root',
  resolve_comment: 'ids from `genoffice docs read --comments`',
  accept_changes:
    'selector: all: true, ids from `genoffice docs read --revisions` (positional, re-read after every edit), or author / type / blockIndex / blockRange / before; insertions become plain text, deleted text goes',
  reject_changes:
    'same selector as accept_changes; inserted text goes, deleted text comes back, formatting reverts',
  insert_footnote:
    'a superscript reference mark goes into blockIndex (right after afterText, else at the block end) and the note text prints at the page bottom; `genoffice docs read --notes` lists notes',
  insert_endnote: 'like insert_footnote, but the note text collects at the end of the document',
  edit_note:
    'findReplace inside one footnote or endnote (id from `genoffice docs read --notes`; kind only when a footnote and an endnote share the id): every occurrence of find becomes replace, the note keeps its id, reference mark and formatting; matchCase defaults to true',
  delete_note:
    'kind footnote|endnote and an id from `genoffice docs read --notes`; removes the reference mark too',
  delete_comment:
    'removes the comment and its anchor; a thread root with replies needs withReplies: true, a reply id removes just that reply',
}

const AFTER_BLOCK_INDEX =
  'insert after this block index; -1 = start of document; omitted = end of document'

const INSERT_IMAGE_SCHEMA: JsonSchema = {
  type: 'object',
  properties: {
    url: { type: 'string', description: 'local path, data: URL or http(s) URL of a png/jpg/gif' },
    maxWidthPx: { type: 'integer', description: 'maximum width in px, default 480' },
    afterBlockIndex: { type: 'integer', description: AFTER_BLOCK_INDEX },
  },
  required: ['url'],
}

/** insert_picture / set_watermark headless: the in-app schema, with the image source widened to local paths */
function insertPictureSchema(tool: JsonSchema, key: 'url' | 'image' = 'url'): JsonSchema {
  const copy = cliToolSchema(tool)
  const src = copy.properties?.[key]
  if (src)
    src.description =
      key === 'url'
        ? 'local path, data: URL or http(s) URL of a png/jpg/gif'
        : 'picture watermark: local path, data: URL or http(s) URL of a png/jpg/gif, centered behind the body; null removes the watermark'
  return copy
}

/** The in-app tool schemas describe the cursor default; headless, an omitted position appends. */
function cliToolSchema(schema: JsonSchema): JsonSchema {
  const copy = JSON.parse(JSON.stringify(schema)) as JsonSchema
  const after = copy.properties?.afterBlockIndex
  if (after) after.description = AFTER_BLOCK_INDEX
  return copy
}

/** The engine signatures read `{ op: "name", fields }  // note`; the catalog keeps the two apart. */
function splitSignature(signature: string): { signature: string; note?: string } {
  const [sig, note] = signature.split(/\s+\/\/\s+/, 2)
  const fields = sig!.replace(/^\{\s*op:\s*"[^"]+",\s*/, '{ ')
  return note ? { signature: fields, note } : { signature: fields }
}

export async function docsCatalog(): Promise<{ catalog: OpCatalog; htmlRules: string }> {
  const mods = await docsModules()
  const block: OpEntry[] = mods.ops.opCatalog().map((d) => ({
    op: d.name,
    group: 'block',
    ...splitSignature(d.signature),
    target: d.target,
    keys: [...d.keys],
  }))
  const tools = new Map(mods.tools.AGENT_TOOLS.map((t) => [t.name, t]))
  const content: OpEntry[] = CONTENT_OPS.map((op) => {
    const tool = tools.get(op)
    if (!tool && op !== 'insert_image') throw new Error(`docs op catalog: no tool named ${op}`)
    const schema =
      op === 'insert_image'
        ? INSERT_IMAGE_SCHEMA
        : op === 'insert_picture'
          ? insertPictureSchema(tool!.inputSchema as JsonSchema)
          : op === 'set_watermark'
            ? insertPictureSchema(tool!.inputSchema as JsonSchema, 'image')
            : cliToolSchema(tool!.inputSchema as JsonSchema)
    return {
      op,
      group: 'content',
      signature: signatureFromSchema(schema, []),
      note: CONTENT_NOTES[op],
      schema,
    }
  })
  const groups = [
    {
      name: 'block',
      summary: 'edits of existing blocks, addressed by target',
      ops: block.map((o) => o.op),
    },
    {
      name: 'content',
      summary:
        'new content, pictures, text boxes, charts, header/footer, page setup, watermark, styles, comments, tracked changes, footnotes',
      ops: [...CONTENT_OPS],
    },
  ]
  return {
    catalog: withFingerprint({ domain: 'docs', groups, ops: [...block, ...content] }),
    htmlRules: mods.protocol.HTML_RULES,
  }
}
