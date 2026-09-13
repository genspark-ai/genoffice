import type { Editor } from '@tiptap/core'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  findChartWorkbookPath,
  parseChartPartXml,
  parseDocx,
  patchChartPartXml,
  patchChartWorkbookXlsxBase64,
  readDocxPartBase64,
  readSections,
  saveDocx,
  type CommentInfo,
  type HeaderFooter,
  type HfPartInfo,
  type SaveOptions,
} from '@genoffice/docx-engine'
import { ensureDom } from '../dom'
import type { PathContext } from '../fs'
import { CliError, EXIT } from '../result'
import { readImageSource } from './image-source'
import { imageSize } from './image-size'

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
  ])
  locale.setModuleLang('en')
  return { Editor, extensions, convert, protocol, ops, tools, comments, hfText, revisions }
}

type Parsed = Awaited<ReturnType<typeof parseDocx>>
type Modules = Awaited<ReturnType<typeof docsModules>>

type HfView = 'default' | 'first' | 'even'
type HfSlot = `${'header' | 'footer'}${'' | 'First' | 'Even'}`

/** Header/footer parts and comments live outside the ProseMirror document; edits are kept here until save. */
interface SideState {
  hf: Partial<Record<HfSlot, HeaderFooter | null>>
  hfDirty: Set<HfSlot>
  titlePg: boolean
  evenOddHf: boolean
  titlePgDirty: boolean
  evenOddHfDirty: boolean
  comments: CommentInfo[]
  commentsDirty: boolean
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
  const options: SaveOptions = {
    header: hf('header'),
    footer: hf('footer'),
    headerFirst: hf('headerFirst'),
    footerFirst: hf('footerFirst'),
    headerEven: hf('headerEven'),
    footerEven: hf('footerEven'),
    titlePg: side.titlePgDirty ? side.titlePg : undefined,
    evenAndOddHeaders: side.evenOddHfDirty ? side.evenOddHf : undefined,
    comments: side.commentsDirty ? side.comments : undefined,
    ...(await chartPartPatches(doc, plan.chartPatches)),
  }
  return saveDocx(doc.parsed, plan.saveBlocks, options)
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
}

/** Block list an agent targets ops at: index, node type, heading level, text (clipped unless `full`). */
export function describeDocument(
  doc: OpenDocument,
  range?: [number, number],
  full = false,
): BlockSummary[] {
  const out: BlockSummary[] = []
  const root = doc.editor.state.doc
  const [start, end] = range ?? [0, root.childCount - 1]
  for (let i = start; i <= Math.min(end, root.childCount - 1); i++) {
    const node = root.child(i)
    const text = node.textContent
    out.push({
      index: i,
      type: node.type.name.replace(/^doc/, '').toLowerCase(),
      ...(typeof node.attrs.level === 'number' ? { level: node.attrs.level } : {}),
      ...(node.type.name === 'docProtected' ? { kind: protectedKind(node.attrs) } : {}),
      text: !full && text.length > 200 ? `${text.slice(0, 197)}...` : text,
    })
  }
  return out
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
  if (nodes.length === 0) throw new CliError(EXIT.usage, 'no content could be parsed from the HTML')
  const last = doc.editor.state.doc.childCount - 1
  doc.mods.protocol.replaceBlockRange(doc.editor, 0, last, nodes)
  return nodes.length
}

export interface DocOpResult {
  op: string
  output: string
}

const HTML_TOOLS = new Set(['insert_content', 'replace_blocks'])
/** app tools reachable headless besides the html pair; insert_image is built here (no renderer image bridge) */
const SIDE_TOOLS = new Set([
  'insert_chart',
  'edit_chart',
  'set_header_footer',
  'reply_comment',
  'resolve_comment',
  'read_comments',
  'read_revisions',
])

export interface ApplyOptions {
  ctx?: PathContext
  /** folder of the ops file, for relative image paths */
  baseDir?: string
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
): Promise<DocOpResult[]> {
  const results: DocOpResult[] = []
  const comments = commentsAccess(doc)
  const hf = headerFooterAccess(doc)
  for (const [index, op] of ops.entries()) {
    const name = typeof op.op === 'string' ? op.op : ''
    const reject = (output: string) =>
      new CliError(EXIT.usage, `op ${index} (${name || '?'}) rejected: ${output}`, {
        failures: [{ index, op: name, error: output }],
      })
    if (name === 'insert_image') {
      const r = await insertImage(doc, op, opts)
      if (r.error) throw reject(r.error)
      results.push({ op: name, output: r.output! })
      continue
    }
    const call = HTML_TOOLS.has(name)
      ? { id: `genoffice-${index}`, name, input: htmlToolInput(doc, name, op) }
      : SIDE_TOOLS.has(name)
        ? { id: `genoffice-${index}`, name, input: sideToolInput(doc, name, op) }
        : { id: `genoffice-${index}`, name: 'apply_ops', input: { ops: [op] } }
    const exec = await doc.mods.tools.executeTool(
      doc.editor,
      call,
      doc.numIds,
      undefined,
      undefined,
      null,
      comments,
      hf,
    )
    // the in-app executor reports a target that matched nothing as a soft result; for a
    // scripted batch that is a wrong index, so the whole batch is refused before saving
    if (exec.isError || /^No matching blocks/i.test(exec.output)) throw reject(exec.output)
    results.push({ op: name, output: exec.output })
  }
  return results
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

function commentsAccess(doc: OpenDocument): {
  list(): CommentInfo[]
  reply(parentId: string, text: string): boolean
  resolve(id: string): boolean
} {
  const { side } = doc
  return {
    list: () => side.comments,
    reply: (parentId, text) => {
      const id = doc.mods.comments.nextCommentId(side.comments)
      if (!doc.mods.comments.addReplyToCommentRange(doc.editor, parentId, id)) return false
      const date = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
      side.comments.push({ id, author: CLI_COMMENT_AUTHOR, date, text, parentId })
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
 * The app's insert_image downloads through the renderer bridge and measures
 * with an <img>; neither exists here, so the node is built from the bytes
 * directly, then the executor is told the document is still the one it saw.
 */
async function insertImage(
  doc: OpenDocument,
  op: Record<string, unknown>,
  opts: ApplyOptions,
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
  const maxW = Number(op.maxWidthPx) || 480
  const scale = Math.min(1, maxW / size.width)
  const w = Math.max(1, Math.round(size.width * scale))
  const h = Math.max(1, Math.round(size.height * scale))
  const count = doc.editor.state.doc.childCount
  const afterRaw = op.afterBlockIndex === undefined ? count - 1 : Number(op.afterBlockIndex)
  if (!Number.isInteger(afterRaw) || afterRaw < -1 || afterRaw >= count) {
    return { error: `afterBlockIndex must be -1..${count - 1}` }
  }
  const after = afterRaw
  const pos = after < 0 ? 0 : doc.mods.protocol.blockRangePositions(doc.editor, after, after).to
  const base64 = Buffer.from(source.bytes).toString('base64')
  const mime = source.mime as 'image/png' | 'image/jpeg' | 'image/gif'
  doc.editor
    .chain()
    .insertContentAt(pos, {
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'image',
        label: 'Image',
        imageDataUrl: `data:${mime};base64,${base64}`,
        imageWidthPx: w,
        imageHeightPx: h,
        genImage: { base64, mime, widthPx: w, heightPx: h },
      },
    })
    .run()
  doc.mods.tools.markDocSeen(doc.editor)
  return { output: `Inserted the image (${w}x${h}px) after block ${after}.` }
}

export interface CommentSummary extends CommentInfo {
  blockIndex?: number
  anchorText?: string
}

/** Every comment with the block its anchor sits in, ids as reply_comment / resolve_comment take them. */
export function listComments(doc: OpenDocument): CommentSummary[] {
  const anchors = doc.mods.protocol.commentAnchors(doc.editor)
  return doc.side.comments.map((c) => {
    const a = anchors.get(c.parentId ?? c.id)
    return a ? { ...c, blockIndex: a.blockIndex, anchorText: a.excerpt } : { ...c }
  })
}

export interface RevisionSummary {
  blockIndex: number
  kind: string
  author: string
  date?: string
  text: string
}

/** Pending tracked changes in document order. */
export function listRevisions(doc: OpenDocument): RevisionSummary[] {
  const pm = doc.editor.state.doc
  return doc.mods.revisions.collectRevisions(pm).map((rev) => ({
    blockIndex: blockIndexOfPos(pm, rev.from),
    kind: rev.kind,
    author: rev.author,
    ...(rev.date ? { date: rev.date } : {}),
    text: pm.textBetween(rev.from, rev.to, '\n', ' ').replace(/\s+/g, ' ').trim(),
  }))
}

function blockIndexOfPos(pm: Editor['state']['doc'], pos: number): number {
  let index = 0
  let found = 0
  pm.forEach((node, offset) => {
    if (pos >= offset && pos < offset + node.nodeSize) found = index
    index++
  })
  return found
}

export async function docsGuide(): Promise<{ signatures: string[]; htmlRules: string }> {
  const mods = await docsModules()
  return { signatures: mods.ops.opSignatures(), htmlRules: mods.protocol.HTML_RULES }
}
