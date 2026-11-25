// Modified for ForaOffice from GenOffice v0.12.852.
import { getDocsHost } from '../host/docs-host'
import { embeddedImageBytes } from './image-input'
import { readToolImage, toolImageSource, type DocsImageSource } from './image-source'
import { docsToolCatalog } from '../../shared/tool-catalog'
import type { DocsErrorCode } from '../../shared/operation-result'
import type { Editor } from '@tiptap/core'
import type { Mark, Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { ChartDisplay, CommentInfo, NewChart, NoteInfo } from '@genoffice/docx-engine'
import type { ToolDisplay } from '@genoffice/agent-core'
import type { AgentToolCall, CreateDocumentType } from '../../shared/ipc'
import { t } from '../i18n/locale'
import { MIND_MAP_LIMITS, parseMindMapTree } from '@genoffice/agent-core/mind-map'
import { executeOps } from './ops'
import {
  describeStructureDeletion,
  inventoryStructures,
  preservedInlineStructures,
  preservedTableStructures,
  validateStructureDeletion,
} from './structure-inventory'
import { clipExcerpt, repliesOf, resolveCommentAnchor } from './comment-ops'
import { insertPictureAtSelection } from '../editor/insert-picture'
import {
  SECTION_BREAK_TYPES,
  resolvePageSetup,
  sectionIndexOfBlock,
  pageSetupContextLines,
  sectionLine,
  type AiPageSetupAccess,
  type SectionBreakType,
} from './page-setup'
import {
  insertPosition,
  pictureNode,
  resolveFloat,
  resolvePictureWatermark,
  resolveWatermark,
  textBoxNode,
  type AiWatermarkAccess,
  type FloatSpec,
} from './floating-ops'
import { describeStyles, resolveStyleDefinition, type AiStyleAccess } from './style-ops'
import {
  applyRevisionSelection,
  describePending,
  listRevisionEntries,
  validateSelector,
  type StoryRevisionAccess,
} from './revision-ops'
import {
  buildNotesContext,
  editNoteText,
  insertNoteRef,
  noteInsertPos,
  removeNoteRefs,
  type AiNotesAccess,
  type NoteKind,
} from './note-ops'
import {
  blockRangePositions,
  buildCommentsContext,
  buildDocumentContext,
  buildRevisionsContext,
  insertBlocksAfter,
  inheritBlockFormatting,
  preservedImageStructureIds,
  isBlankDocument,
  isTrackedDeleted,
  liveInlineRanges,
  liveTextBetween,
  parseHtmlFragment,
  readOnlyHtmlError,
  resolveImageRefs,
  countInlineImages,
  parseInlineFragment,
  replaceBlockRange,
  replaceInlineRange,
  serializeRangeToHtml,
  type AiHfState,
  type AiTrack,
  type NumIds,
  type SelectionScope,
} from './protocol'

/**
 * Local agent tools: a document-context reader plus a script-style execution
 * channel. The "execution backend" is the in-process ProseMirror doc, split
 * into three safe primitives plus the deterministic command engine.
 */

const READ_MAX_CHARS = 24_000

export const DOCS_EDITOR_TOOLS = docsToolCatalog()

/**
 * App-owned header/footer state, handed to the tool executor. Writes run the
 * same commit path as on-canvas editing (variant routing, per-section edits,
 * dirty flags), so the docx save path needs no changes.
 */
export interface AiHeaderFooterAccess {
  read(): AiHfState
  /** returns an error message, or null on success */
  set(
    kind: 'header' | 'footer',
    view: 'default' | 'first' | 'even',
    text: string,
    section?: number,
    scope?: 'section' | 'shared',
    textboxIndex?: number,
  ): string | null
}

/**
 * The App-owned comments store, handed to the tool executor. Mutations run
 * the same review-actions code paths as the comments pane, so AI replies and
 * resolves behave exactly like manual ones (anchors, dirty flags, docx save).
 */
/** Document-level stores beyond the PM doc (styles.xml, the page watermark), handed to the tool executor. */
export interface AiDocExtras {
  styles?: AiStyleAccess
  watermark?: AiWatermarkAccess
  /** protection (read-only or enforced tracked changes) forbids accepting / rejecting revisions */
  revisionsLocked?: boolean
  /** tracked changes in notes, header / footer parts and textboxes */
  storyRevisions?: StoryRevisionAccess
}

export interface AiCommentsAccess {
  list(): CommentInfo[]
  /** false when the parent thread or its anchor no longer exists */
  reply(parentId: string, text: string): boolean
  /** false when the thread does not exist */
  resolve(id: string): boolean
  /** new thread anchored to the range; the new comment's id, null when the range holds no text */
  add?(range: { from: number; to: number }, text: string, meta: CommentAuthorMeta): string | null
  /** the comment and its replies, anchors included; false when the id does not exist */
  remove?(id: string): boolean
}

export interface CommentAuthorMeta {
  author?: string
  initials?: string
}

/**
 * Selection frozen at context build, valid only while `doc` is still the live
 * document. A user click elsewhere keeps the doc identical (selection-only
 * transaction) so the freeze holds; once any edit lands, the live selection —
 * which ProseMirror has remapped through those edits — is the correct target
 * again and the frozen block indexes would drift, so the freeze is dropped.
 */
export interface FrozenSelection {
  scope: SelectionScope
  doc: ProseMirrorNode
}

export interface ToolExecution {
  /** result text fed back to the model */
  output: string
  isError?: boolean
  errorCode?: DocsErrorCode
  /** Independent host effects may commit without changing this open document. */
  committed?: boolean | null
  /** true when the tool changed the document */
  mutated: boolean
  /** short human-readable label for the chat activity chip */
  summary: string
  /** Exact match/change counts when all operations report those units. */
  matchCount?: number
  changeCount?: number
  /** UI-only side channel (never enters the LLM context) */
  display?: ToolDisplay
}

const fail = (
  summary: string,
  output: string,
  errorCode: DocsErrorCode = 'invalid_input',
): ToolExecution => ({
  output,
  isError: true,
  mutated: false,
  summary,
  errorCode,
})

const clipText = (text: string, max: number): string =>
  text.length > max ? `${text.slice(0, max)}…` : text

/**
 * A model that saw gateway-flattened tool results can regurgitate them as the
 * html argument (raw {"index":…} block dumps, literal </tool_response> tags);
 * reject those so protocol artifacts never land in the document as text.
 */
function toolEchoError(html: string): string | null {
  if (/<\/?tool_response>/i.test(html)) {
    return 'html contains a literal <tool_response> tag — that is tool-protocol output, not document content; retry with the actual restricted-HTML fragment'
  }
  // the context/read dump shape is screened anywhere in the payload (fenced or
  // prose-wrapped dumps included) — it is never legitimate document content
  if (/"index"\s*:\s*\d+\s*,\s*"type"\s*:\s*"/.test(html)) {
    return 'html contains a raw JSON block dump, not an HTML fragment; retry with restricted HTML (e.g. <p>…</p>)'
  }
  // unwrap only a fully fenced payload; an embedded fence inside otherwise
  // valid HTML must not shadow the real content
  let candidate = html.trim()
  const fence = /^```[a-z]*\s*([\s\S]*?)```\s*$/i.exec(candidate)
  if (fence) candidate = fence[1].trim()
  if (!/^[[{]/.test(candidate)) return null
  try {
    JSON.parse(candidate)
    return 'html is raw JSON, not an HTML fragment; retry with restricted HTML (e.g. <p>…</p>)'
  } catch {
    return null // brace-led plain text is legitimate content
  }
}

/** <sel> only marks the user's selection in the context; a model echoing it would write the marker into the document */
function contextMarkerError(html: string): string | null {
  return /<\/?sel\s*>/i.test(html)
    ? 'html contains <sel>: that marker only delimits the selection in the context and is not document content; retry without it'
    : readOnlyHtmlError(html)
}

/** Old read responses used plain paragraphs; refuse their echo only at protected targets. */
function legacyProtectedEchoError(
  editor: Editor,
  start: number,
  end: number,
  html: string,
): string | null {
  let protectedTarget = false
  for (let i = start; i <= end; i++) {
    const node = editor.state.doc.child(i)
    if (node.type.name === 'docProtected' && !isTrackedDeleted(node)) protectedTarget = true
  }
  if (!protectedTarget) return null
  const parsed = new DOMParser().parseFromString(html, 'text/html')
  const legacy = /^\[Protected (?:content|formula):[\s\S]*, kept as is\]$/
  if (
    [parsed.body, ...parsed.body.querySelectorAll('p,div')].some((node) =>
      legacy.test(node.textContent?.trim() ?? ''),
    )
  )
    return 'HTML echoes an old read-only protected-content description, not an object reference. No changes were made. Read the current blocks and edit surrounding text in place, or use explicit structure operations for intentional deletion.'
  return null
}

function replacementStructureNote(
  editor: Editor,
  start: number,
  end: number,
  tracked: boolean,
): string {
  const labels: string[] = []
  let fields = false
  for (let i = start; i <= end; i++) {
    const node = editor.state.doc.child(i)
    if (isTrackedDeleted(node)) continue
    if (node.type.name === 'docProtected')
      labels.push(String(node.attrs.label || node.attrs.blockType || 'content'))
    node.descendants((child) => {
      if (
        !child.marks.some((mark) => mark.type.name === 'del') &&
        child.marks.some((mark) => mark.type.name === 'instrField' || mark.type.name === 'refField')
      )
        fields = true
    })
  }
  const objects = labels.length
    ? ` ${labels.length} protected block(s) (${[...new Set(labels)].join(', ')}) ${tracked ? 'were marked for deletion' : 'were replaced'}; their read-only descriptions do not preserve objects.`
    : ''
  return (
    objects +
    (fields
      ? ' The range contained Word fields; visible HTML text does not encode their field instructions.'
      : '')
  )
}

function preservedImageStructures(
  editor: Editor,
  kept: Array<{ block: number; index: number }>,
): string[] {
  return preservedImageStructureIds(editor.state.doc, kept)
}

/** Pure refusal preflight, also called by the runtime before opening an undo group. */
export function inspectHtmlStructureIntent(
  editor: Editor,
  call: AgentToolCall,
  scope?: SelectionScope | null,
  numIds: NumIds = { bullet: null, ordered: null },
): { error: string | null; plan?: import('../../shared/structure-plan').DocsStructurePlan } {
  if (call.name !== 'replace_blocks' && call.name !== 'replace_selection') return { error: null }
  const inspect = (from: number, to: number, preserved: string[]) => {
    const affected = inventoryStructures(editor.state.doc, from, to)
    return {
      error: validateStructureDeletion(
        editor.state.doc,
        from,
        to,
        call.input.deleteStructures,
        preserved,
      ),
      plan: describeStructureDeletion(
        affected,
        affected.filter((item) => !preserved.includes(item.id)),
        (call.input.deleteStructures ?? []) as string[],
      ),
    }
  }
  try {
    const html = String(call.input.html ?? '')
    const echo = toolEchoError(html) ?? contextMarkerError(html)
    if (echo) return { error: echo }
    if (call.name === 'replace_blocks') {
      const range = validRange(editor, call.input.startBlockIndex, call.input.endBlockIndex)
      if (!range) return { error: rangeError(editor) }
      const protectedEcho = legacyProtectedEchoError(editor, range.start, range.end, html)
      if (protectedEcho) return { error: protectedEcho }
      const resolved = resolveImageRefs(editor, parseHtmlFragment(html, numIds), html)
      if (!resolved.nodes.length) return { error: 'html did not parse into any content blocks' }
      const nodes = inheritBlockFormatting(editor, range.start, range.end, resolved.nodes, false)
      const positions = blockRangePositions(editor, range.start, range.end)
      const kept = [
        ...preservedImageStructures(editor, resolved.kept),
        ...preservedTableStructures(
          editor.state.doc,
          positions.from,
          positions.to,
          nodes.map((node) => editor.schema.nodeFromJSON(node)),
        ),
        ...preservedInlineStructures(
          editor.state.doc,
          positions.from,
          positions.to,
          nodes.map((node) => editor.schema.nodeFromJSON(node)),
          call.input.deleteStructures,
        ),
      ]
      return inspect(positions.from, positions.to, kept)
    }
    const range = selectionRange(editor, scope)
    if (!range) return { error: 'nothing is selected: replace_selection needs a range selection' }
    const $from = editor.state.doc.resolve(range.from)
    const $to = editor.state.doc.resolve(range.to)
    if (!$from.sameParent($to) || !$from.parent.isTextblock)
      return { error: 'replace_selection only edits text inside one paragraph/heading/list item' }
    if (!liveInlineRanges(editor.state.doc, range.from, range.to).length)
      return { error: 'the selection is a pending tracked deletion' }
    const resolved = resolveImageRefs(editor, parseInlineFragment(html), html)
    if (!resolved.nodes.length) return { error: 'html did not parse into any text' }
    return inspect(range.from, range.to, preservedImageStructures(editor, resolved.kept))
  } catch (error) {
    return { error: error instanceof Error ? error.message : String(error) }
  }
}

export function validateHtmlStructureIntent(
  ...args: Parameters<typeof inspectHtmlStructureIntent>
): string | null {
  return inspectHtmlStructureIntent(...args).error
}

/** the range selection tools act on: the frozen scope when valid, else the live selection; null for a caret */
function selectionRange(
  editor: Editor,
  scope?: SelectionScope | null,
): { from: number; to: number } | null {
  if (scope) {
    if (!scope.isRange) return null
    if (scope.from !== undefined && scope.to !== undefined && scope.from < scope.to)
      return { from: scope.from, to: scope.to }
  }
  const { from, to, empty } = editor.state.selection
  return empty ? null : { from, to }
}

/** formatting the fragment expresses itself (tags), so it must not also be inherited from the old text */
const FRAGMENT_MARK_TYPES = new Set(['bold', 'italic', 'underline', 'strike', 'link'])
const REVISION_MARK_TYPES = new Set(['ins', 'del'])

/** Doc as last seen by the AI pipeline (context build / read / own write); a differing doc means the user edited in between. */
const docBaseline = new WeakMap<Editor, ProseMirrorNode>()

export interface DocReadState {
  isCurrent(): boolean
  markSeen(): void
  apply?<T>(operation: () => T): T
  selection?(): SelectionScope | null
  /** Project image identity before paging, so continuation offsets address the returned HTML. */
  imageHtml?(html: string): string
}

type SavedSideState = { before: unknown; after: unknown }
type BaselineCarry = (
  before: ProseMirrorNode,
  after: ProseMirrorNode,
  savedSideState?: SavedSideState,
) => void
const baselineCarries = new WeakMap<Editor, Set<BaselineCarry>>()

/** Trusted load/field refreshes may carry an existing read, never create a new read. */
export function onDocSeenCarry(editor: Editor, callback: BaselineCarry): () => void {
  const listeners = baselineCarries.get(editor) ?? new Set<BaselineCarry>()
  listeners.add(callback)
  baselineCarries.set(editor, listeners)
  return () => {
    listeners.delete(callback)
  }
}

export function markDocSeen(editor: Editor, readState?: DocReadState): void {
  if (readState) readState.markSeen()
  else docBaseline.set(editor, editor.state.doc)
}

/** Trusted representation refreshes carry an existing read, never acknowledge unseen edits. */
export function carryDocSeen(
  editor: Editor,
  before: ProseMirrorNode,
  savedSideState?: SavedSideState,
): void {
  if (docBaseline.get(editor) === before) docBaseline.set(editor, editor.state.doc)
  for (const carry of baselineCarries.get(editor) ?? [])
    carry(before, editor.state.doc, savedSideState)
}

function editedExternally(editor: Editor, readState?: DocReadState): boolean {
  if (readState) return !readState.isCurrent()
  const seen = docBaseline.get(editor)
  return seen !== undefined && seen !== editor.state.doc
}

/** tools addressing the document by block index: refused after an external edit until the model re-reads */
const INDEX_WRITE_SUMMARIES: Record<string, () => string> = {
  insert_content: () => t('aiSumInsertContent'),
  replace_blocks: () => t('aiSumReplaceContent'),
  replace_selection: () => t('aiSumReplaceSelection'),
  apply_ops: () => t('aiSumApplyCommands'),
  insert_chart: () => t('aiSumInsertChart'),
  edit_chart: () => t('aiSumEditChart'),
  insert_section_break: () => t('aiSumInsertSectionBreak'),
  accept_changes: () => t('aiSumAcceptChanges'),
  reject_changes: () => t('aiSumRejectChanges'),
  insert_text_box: () => t('aiSumInsertTextBox'),
  insert_picture: () => t('aiSumInsertPicture'),
  insert_footnote: () => t('aiSumInsertFootnote'),
  insert_endnote: () => t('aiSumInsertEndnote'),
  add_comment: () => t('aiSumAddComment'),
}

/** set_page_setup only addresses the document by index through blockIndex; section / all-sections stay valid */
function staleSummaryOf(call: AgentToolCall): (() => string) | undefined {
  if (call.name === 'set_page_setup') {
    const { blockIndex } = call.input
    return blockIndex === undefined || blockIndex === null
      ? undefined
      : () => t('aiSumSetPageSetup')
  }
  return INDEX_WRITE_SUMMARIES[call.name]
}

const STALE_DOC_ERROR =
  'The document was edited by the user since it was last read; block indexes may be stale. ' +
  'Call get_document_context (or read_blocks) to get the current state, then retry.'

function rangeError(editor: Editor): string {
  return `block index invalid or out of range (the document has ${editor.state.doc.childCount} blocks); call get_document_context for fresh indexes`
}

/** the embedded picture of a top-level block; null when the block is not an image block */
export function imageBlockSource(
  editor: Editor,
  index: unknown,
): { src: string; label: string | undefined } | null {
  if (!Number.isInteger(index)) return null
  const i = Number(index)
  const doc = editor.state.doc
  if (i < 0 || i >= doc.childCount) return null
  const node = doc.child(i)
  if (node.type.name !== 'docProtected' || node.attrs.blockType !== 'image') return null
  const src = typeof node.attrs.imageDataUrl === 'string' ? node.attrs.imageDataUrl : ''
  if (!src) return null
  return { src, label: typeof node.attrs.label === 'string' ? node.attrs.label : undefined }
}

function validRange(
  editor: Editor,
  start: unknown,
  end: unknown,
): { start: number; end: number } | null {
  const count = editor.state.doc.childCount
  if (!Number.isInteger(start) || !Number.isInteger(end)) return null
  const s = Number(start)
  const e = Number(end)
  // an out-of-range end must surface as an error, not silently clamp onto the wrong blocks
  if (s < 0 || e < s || e >= count) return null
  return { start: s, end: e }
}

/** Read the natural size of a dataURL image. */
function imageSizeOf(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve({ width: img.naturalWidth || 1, height: img.naturalHeight || 1 })
    img.onerror = () => reject(new Error('image load failed'))
    img.src = dataUrl
  })
}

/** Async tools: web search / image search / insert web image. */
async function executeAsyncTool(
  editor: Editor,
  call: AgentToolCall,
  signal?: AbortSignal,
  readState?: DocReadState,
  scope?: SelectionScope | null,
): Promise<ToolExecution> {
  const originalSelection = { from: editor.state.selection.from, to: editor.state.selection.to }
  const imageSelection = () => {
    const current = readState?.selection?.() ?? scope
    return current?.from !== undefined && current.to !== undefined
      ? { from: current.from, to: current.to }
      : originalSelection
  }
  switch (call.name) {
    case 'insert_picture':
      return insertPicture(editor, call, signal, readState)
    case 'insert_image': {
      const url = toolImageSource(call.input)
      return insertImageFromUrl(
        editor,
        url,
        Number(call.input.maxWidthPx) || 480,
        signal,
        {
          failLabel: t('aiSumInsertImage'),
          doneLabel: t('aiSumInsertWebImage'),
          blockLabel: typeof call.input.label === 'string' ? call.input.label : 'Image (web)',
        },
        readState,
        imageSelection,
      )
    }
    case 'create_document': {
      const typeRaw = call.input.type === undefined ? 'docx' : String(call.input.type)
      if (typeRaw !== 'docx' && typeRaw !== 'pdf' && typeRaw !== 'md' && typeRaw !== 'html')
        return fail(t('aiSumCreateDocument'), 'type must be one of docx/pdf/md/html')
      const type: CreateDocumentType = typeRaw
      const title = String(call.input.title ?? '').trim()
      if (!title) return fail(t('aiSumCreateDocument'), 'title must not be empty')
      const content = String(call.input.content ?? '')
      if (!content.trim()) return fail(t('aiSumCreateDocument'), 'content must not be empty')
      if (type !== 'md') {
        const echo = toolEchoError(content) ?? contextMarkerError(content)
        if (echo) return fail(t('aiSumCreateDocument'), echo)
      }
      if (type === 'docx' || type === 'pdf') {
        // the new docx tab fills itself after this tool already returned, so
        // unparseable HTML must be rejected here, where the model can retry
        try {
          if (parseHtmlFragment(content, { bullet: null, ordered: null }).length === 0)
            return fail(t('aiSumCreateDocument'), 'content did not parse into any content blocks')
        } catch (e) {
          return fail(t('aiSumCreateDocument'), e instanceof Error ? e.message : String(e))
        }
      }
      let r
      try {
        r = await getDocsHost().createDocument({ type, title, content })
      } catch (error) {
        return {
          ...fail(
            t('aiSumCreateDocument'),
            error instanceof Error ? error.message : String(error),
            'host_failure',
          ),
          committed: null,
        }
      }
      if (!r.ok)
        return {
          ...fail(
            t('aiSumCreateDocument'),
            r.error ?? 'creating the document failed',
            r.errorCode ?? 'host_failure',
          ),
          committed: r.committed ?? null,
        }
      const name = `${title}.${type}`
      return {
        output: r.path
          ? `Created the new document at ${r.path}.${r.opened ? ' Opened it in a new view.' : r.opened === false ? ' The host did not open a view; the file is saved at the reported path.' : ''}`
          : `The host accepted the new document "${name}".`,
        mutated: false,
        committed: r.committed === undefined ? true : r.committed,
        summary: t('aiSumCreatedDocument', { name }),
      }
    }
    default:
      return fail(t('aiSumUnknownTool'), call.name)
  }
}

/** magic-byte sniff: the fetch handler's content-type mapping defaults unknown
 *  types to jpeg, and a webp/svg mislabeled as jpeg breaks the exported docx */
export function sniffImageMime(base64: string): 'image/png' | 'image/jpeg' | 'image/gif' | null {
  let head: string
  try {
    head = atob(base64.slice(0, 12))
  } catch {
    return null
  }
  if (head.startsWith('\x89PNG')) return 'image/png'
  if (head.startsWith('GIF8')) return 'image/gif'
  if (head.charCodeAt(0) === 0xff && head.charCodeAt(1) === 0xd8) return 'image/jpeg'
  return null
}

/** set_watermark with an image: download / decode, then hand the measured bytes to the store */
async function setPictureWatermark(
  editor: Editor,
  call: AgentToolCall,
  signal: AbortSignal | undefined,
  extras: AiDocExtras | undefined,
  readState?: DocReadState,
): Promise<ToolExecution> {
  const summary = t('aiSumSetWatermark')
  if (!extras?.watermark) return fail(summary, 'the watermark is not available here')
  const url = String(call.input.image).trim()
  if (!/^https?:\/\/|^data:image\//.test(url))
    return fail(summary, 'image must be an http(s) or data: image URL')
  const fetched = url.startsWith('data:')
    ? embeddedImageBytes(url)
    : await getDocsHost().fetchImage(url)
  if (signal?.aborted) return fail(summary, 'stopped by the user; the watermark was not set')
  if (!fetched)
    return fail(
      summary,
      'download failed (the image host refused or the link expired); do not retry this URL — pick a different result from the image search or search again',
    )
  const mime = sniffImageMime(fetched.base64)
  if (!mime) return fail(summary, 'unsupported image format (only png/jpg/gif can be embedded)')
  let natural: { width: number; height: number }
  try {
    natural = await imageSizeOf(`data:${mime};base64,${fetched.base64}`)
  } catch {
    return fail(summary, 'the image could not be decoded')
  }
  const resolved = resolvePictureWatermark(call.input, {
    base64: fetched.base64,
    mime,
    widthPx: natural.width,
    heightPx: natural.height,
  })
  if ('error' in resolved) return fail(summary, resolved.error)
  if (signal?.aborted) return fail(summary, 'stopped by the user; the watermark was not set')
  if (editedExternally(editor, readState)) return fail(summary, STALE_DOC_ERROR, 'stale_target')
  const commit = () => extras.watermark!.set(resolved.spec)
  const error = readState?.apply ? readState.apply(commit) : commit()
  if (error) return fail(summary, error)
  return {
    output: `Picture watermark set (${natural.width}x${natural.height}px source).`,
    mutated: false, // header part state, written on save
    summary,
  }
}

/** insert_picture: sized, optionally floating picture after a block */
async function insertPicture(
  editor: Editor,
  call: AgentToolCall,
  signal: AbortSignal | undefined,
  readState?: DocReadState,
): Promise<ToolExecution> {
  const summary = t('aiSumInsertPicture')
  const url = toolImageSource(call.input)
  let float: FloatSpec | undefined
  if (call.input.float !== undefined) {
    const r = resolveFloat(call.input.float)
    if ('error' in r) return fail(summary, r.error)
    float = r.float
  }
  const at = insertPosition(editor, call.input.afterBlockIndex)
  if ('error' in at) return fail(summary, at.error)
  const fetched = await readToolImage(url)
  if (signal?.aborted) return fail(summary, 'stopped by the user; the picture was not inserted')
  if (!fetched)
    return fail(
      summary,
      'download failed (the image host refused or the link expired); do not retry this URL — pick a different result from the image search or search again',
    )
  const mime = sniffImageMime(fetched.base64)
  if (!mime) return fail(summary, 'unsupported image format (only png/jpg/gif can be embedded)')
  let natural: { width: number; height: number }
  try {
    natural = await imageSizeOf(`data:${mime};base64,${fetched.base64}`)
  } catch {
    return fail(summary, 'the image could not be decoded')
  }
  if (signal?.aborted) return fail(summary, 'stopped by the user; the picture was not inserted')
  const built = pictureNode({
    base64: fetched.base64,
    mime,
    naturalWidth: natural.width,
    naturalHeight: natural.height,
    width: call.input.width,
    height: call.input.height,
    float,
    ...(typeof call.input.altText === 'string' ? { altText: call.input.altText } : {}),
  })
  if ('error' in built) return fail(summary, built.error)
  if (editedExternally(editor, readState)) return fail(summary, STALE_DOC_ERROR, 'stale_target')
  const pos = insertPosition(editor, call.input.afterBlockIndex)
  if ('error' in pos) return fail(summary, pos.error)
  editor.chain().insertContentAt(pos.pos, built.node).run()
  markDocSeen(editor, readState)
  return {
    output: `Inserted the ${float ? 'floating ' : ''}picture (${built.widthPx}x${built.heightPx}px) after block ${at.after}.`,
    mutated: true,
    summary,
  }
}

/** download a direct image URL and insert it at the cursor as a protected image block */
async function insertImageFromUrl(
  editor: Editor,
  url: string | DocsImageSource,
  maxW: number,
  signal: AbortSignal | undefined,
  labels: { failLabel: string; doneLabel: string; blockLabel: string },
  readState?: DocReadState,
  selection?: () => { from: number; to: number },
): Promise<ToolExecution> {
  const fetched = await readToolImage(url)
  // never write after the user hit stop (the download may resolve long after the abort)
  if (signal?.aborted)
    return fail(labels.failLabel, 'stopped by the user; the image was not inserted')
  if (!fetched)
    return fail(
      labels.failLabel,
      'download failed (the image host refused or the link expired); do not retry this URL — pick a different result from the image search or search again',
    )
  const mime = sniffImageMime(fetched.base64)
  if (!mime) {
    return fail(
      labels.failLabel,
      'unsupported image format (only png/jpg/gif can be embedded) — pick a different image',
    )
  }
  const dataUrl = `data:${mime};base64,${fetched.base64}`
  try {
    const natural = await imageSizeOf(dataUrl)
    if (signal?.aborted)
      return fail(labels.failLabel, 'stopped by the user; the image was not inserted')
    const scale = Math.min(1, maxW / natural.width)
    const w = Math.round(natural.width * scale)
    const h = Math.round(natural.height * scale)
    // The download can take long: user edits made meanwhile must keep the
    // freshness baseline stale, so only our own insertion may mark the doc
    // seen. Checked right before the write — there is no async gap after.
    const userEditedDuringFetch = editedExternally(editor, readState)
    if (readState && userEditedDuringFetch)
      return fail(labels.failLabel, STALE_DOC_ERROR, 'stale_target')
    const insert = () => {
      if (selection) editor.commands.setTextSelection(selection())
      return insertPictureAtSelection(editor, {
        dataUrl,
        base64: fetched.base64,
        mime,
        widthPx: w,
        heightPx: h,
        label: labels.blockLabel,
      })
    }
    const inserted = readState?.apply ? readState.apply(insert) : insert()
    if (!inserted)
      return fail(
        labels.failLabel,
        'The image could not be inserted at the captured position.',
        'execution_failure',
      )
    if (!userEditedDuringFetch) markDocSeen(editor, readState)
    return {
      output: `Inserted the image (${w}×${h}px).`,
      mutated: true,
      summary: labels.doneLabel,
    }
  } catch {
    return fail(labels.failLabel, 'the image could not be decoded')
  }
}

export function executeTool(
  editor: Editor,
  call: AgentToolCall,
  numIds: NumIds,
  track?: AiTrack,
  signal?: AbortSignal,
  frozen?: FrozenSelection | null,
  comments?: AiCommentsAccess,
  hf?: AiHeaderFooterAccess,
  pageSetup?: AiPageSetupAccess,
  extras?: AiDocExtras,
  notes?: AiNotesAccess,
  readState?: DocReadState,
): ToolExecution | Promise<ToolExecution> {
  if (
    ['replace_blocks', 'replace_selection'].includes(call.name) &&
    (call.input.preconditions !== undefined || call.input.dryRun === true)
  )
    return fail(
      'Document operation refused',
      'Replacement preconditions and previews require the live runtime/MCP entry point; no changes were made.',
    )
  const scope = frozen && frozen.doc === editor.state.doc ? frozen.scope : null
  const staleSummary = staleSummaryOf(call)
  if (staleSummary && editedExternally(editor, readState))
    return fail(staleSummary(), STALE_DOC_ERROR, 'stale_target')
  const settle = (exec: ToolExecution): ToolExecution => {
    const readsDoc = call.name === 'get_document_context' || call.name === 'read_blocks'
    if (exec.mutated || (readsDoc && !exec.isError)) markDocSeen(editor, readState)
    return exec
  }
  // Async tools take a separate Promise branch; the other sync tools keep returning
  // synchronously (doesn't break existing tests). No settle here: marking the doc
  // seen after the long download would baptize user edits made meanwhile —
  // insert_image maintains the baseline itself right at its synchronous write.
  if (
    call.name === 'insert_image' ||
    call.name === 'insert_picture' ||
    call.name === 'create_document'
  ) {
    return executeAsyncTool(editor, call, signal, readState, scope)
  }
  if (call.name === 'set_watermark' && typeof call.input.image === 'string')
    return setPictureWatermark(editor, call, signal, extras, readState)
  return settle(
    executeSyncTool(
      editor,
      call,
      numIds,
      track,
      scope,
      comments,
      hf,
      pageSetup,
      extras,
      notes,
      readState?.imageHtml,
    ),
  )
}

function executeSyncTool(
  editor: Editor,
  call: AgentToolCall,
  numIds: NumIds,
  track?: AiTrack,
  scope?: SelectionScope | null,
  comments?: AiCommentsAccess,
  hf?: AiHeaderFooterAccess,
  pageSetup?: AiPageSetupAccess,
  extras?: AiDocExtras,
  notes?: AiNotesAccess,
  imageHtml?: (html: string) => string,
): ToolExecution {
  switch (call.name) {
    case 'get_document_context':
      return {
        // the still-valid frozen scope keeps the reported selection consistent
        // with what scope:'selection' and cursor-relative inserts will act on
        output: [
          buildDocumentContext(editor, scope ?? undefined, hf?.read()),
          ...(pageSetup ? pageSetupContextLines(pageSetup.list()) : []),
        ].join('\n'),
        mutated: false,
        summary: t('aiSumReadDocContext'),
      }

    case 'generate_mindmap': {
      const summary = t('aiSumMindMap')
      const parsed = parseMindMapTree(call.input, editor.state.doc.childCount)
      if (!parsed.ok) return fail(summary, parsed.error)
      const focus = typeof call.input.focus === 'string' ? call.input.focus.trim() : ''
      const trimmedNote = parsed.trimmed
        ? ` (trimmed to the ${MIND_MAP_LIMITS.maxNodes}-node / depth-${MIND_MAP_LIMITS.maxDepth} limit)`
        : ''
      return {
        output:
          `Validated mind-map data for display: ${parsed.nodeCount} nodes, depth ${parsed.depth}` +
          `${focus ? `, focused on "${focus.slice(0, MIND_MAP_LIMITS.maxFocusChars)}"` : ''}${trimmedNote}. ` +
          'The document is unchanged. Display data is included in the tool result.',
        mutated: false,
        summary,
        display: { kind: 'mindmap', tree: parsed.tree },
      }
    }

    case 'read_blocks': {
      const range = validRange(editor, call.input.startBlockIndex, call.input.endBlockIndex)
      if (!range) return fail(t('aiSumReadBlocks'), rangeError(editor))
      const serialized = serializeRangeToHtml(editor, range.start, range.end)
      const html = imageHtml?.(serialized) ?? serialized
      const offset = Math.max(0, Math.trunc(Number(call.input.offset)) || 0)
      if (offset > 0 && offset >= html.length) {
        return fail(
          t('aiSumReadBlocks'),
          `offset ${offset} is beyond the content (${html.length} characters in total)`,
        )
      }
      const slice = html.slice(offset, offset + READ_MAX_CHARS)
      const end = offset + slice.length
      const note =
        end < html.length
          ? `\n…(truncated: ${html.length} characters in total, call read_blocks again with offset=${end} to continue)`
          : offset > 0
            ? `\n(end of range: ${html.length} characters in total)`
            : ''
      let empty = '(range is empty)'
      if (!slice) {
        let deletedBlocks = 0
        for (let i = range.start; i <= range.end; i++) {
          if (isTrackedDeleted(editor.state.doc.child(i))) deletedBlocks++
        }
        if (deletedBlocks > 0) {
          empty = `(the ${deletedBlocks} block(s) in this range are pending tracked deletions — that text is already deleted and hidden from reads; do not delete or rewrite it again)`
        }
      }
      return {
        output: slice ? slice + note : empty,
        mutated: false,
        summary: t('aiSumReadBlocksRange', { start: range.start, end: range.end }),
      }
    }

    case 'insert_content': {
      const requestedIndex = call.input.afterBlockIndex
      if (
        requestedIndex !== undefined &&
        (!Number.isInteger(requestedIndex) ||
          Number(requestedIndex) < -1 ||
          Number(requestedIndex) >= editor.state.doc.childCount)
      )
        return fail(
          t('aiSumInsertContent'),
          'invalid or out-of-range afterBlockIndex; read fresh document context',
        )
      const html = String(call.input.html ?? '')
      const echo = toolEchoError(html) ?? contextMarkerError(html)
      if (echo) return fail(t('aiSumInsertContent'), echo)
      let nodes: ReturnType<typeof parseHtmlFragment>
      try {
        nodes = parseHtmlFragment(html, numIds)
        nodes = resolveImageRefs(editor, nodes, html).nodes
      } catch (e) {
        return fail(t('aiSumInsertContent'), e instanceof Error ? e.message : String(e))
      }
      if (nodes.length === 0)
        return fail(t('aiSumInsertContent'), 'html did not parse into any content blocks')
      const count = editor.state.doc.childCount
      if (isBlankDocument(editor)) {
        // the blank template's single empty paragraph gets replaced
        replaceBlockRange(editor, 0, count - 1, nodes, track)
        return {
          output: `Inserted ${nodes.length} block(s) (the document was empty). Block indexes have changed; use get_document_context if needed.`,
          mutated: true,
          summary: t('aiSumInsertedBlocks', { count: nodes.length }),
        }
      }
      const cursorScope = call.input.afterBlockIndex === undefined
      const after = cursorScope
        ? getCursorBlockIndex(editor, scope)
        : Number(call.input.afterBlockIndex)
      if (!Number.isInteger(after)) return fail(t('aiSumInsertContent'), 'invalid afterBlockIndex')
      // -1 hits blockRangePositions' 0/0 default, i.e. insert at doc start
      insertBlocksAfter(editor, after, nodes, track)
      return {
        output: `Inserted ${nodes.length} block(s) after block ${after}. Subsequent block indexes have shifted; use get_document_context if needed.`,
        mutated: true,
        summary: t('aiSumInsertedBlocks', { count: nodes.length }),
      }
    }

    case 'replace_blocks': {
      const range = validRange(editor, call.input.startBlockIndex, call.input.endBlockIndex)
      if (!range) return fail(t('aiSumReplaceContent'), rangeError(editor))
      const html = String(call.input.html ?? '')
      const echo =
        toolEchoError(html) ??
        contextMarkerError(html) ??
        legacyProtectedEchoError(editor, range.start, range.end, html)
      if (echo) return fail(t('aiSumReplaceContent'), echo)
      let nodes: ReturnType<typeof parseHtmlFragment>
      let resolved: ReturnType<typeof resolveImageRefs>
      try {
        nodes = parseHtmlFragment(html, numIds)
        resolved = resolveImageRefs(editor, nodes, html)
      } catch (e) {
        return fail(t('aiSumReplaceContent'), e instanceof Error ? e.message : String(e))
      }
      nodes = resolved.nodes
      if (nodes.length === 0)
        return fail(t('aiSumReplaceContent'), 'html did not parse into any content blocks')
      const positions = blockRangePositions(editor, range.start, range.end)
      const inherited = inheritBlockFormatting(editor, range.start, range.end, nodes, false)
      const keptStructures = [
        ...preservedImageStructures(editor, resolved.kept),
        ...preservedTableStructures(
          editor.state.doc,
          positions.from,
          positions.to,
          inherited.map((node) => editor.schema.nodeFromJSON(node)),
        ),
        ...preservedInlineStructures(
          editor.state.doc,
          positions.from,
          positions.to,
          inherited.map((node) => editor.schema.nodeFromJSON(node)),
          call.input.deleteStructures,
        ),
      ]
      const structureError = validateStructureDeletion(
        editor.state.doc,
        positions.from,
        positions.to,
        call.input.deleteStructures,
        keptStructures,
      )
      if (structureError) return fail(t('aiSumReplaceContent'), structureError)
      const had = countInlineImages(editor, range.start, range.end)
      const keptInRange = new Set(
        resolved.kept
          .filter((k) => k.block >= range.start && k.block <= range.end)
          .map((k) => `${k.block}.${k.index}`),
      ).size
      const lost = had - keptInRange
      const structureNote = replacementStructureNote(editor, range.start, range.end, !!track)
      const omittedAnchors: number[] = []
      for (let index = range.start; index <= range.end; index++)
        if (editor.state.doc.child(index).attrs.sectionBoundary) omittedAnchors.push(index)
      replaceBlockRange(editor, range.start, range.end, nodes, track, omittedAnchors)
      const pictureNote =
        lost > 0
          ? ` ${lost} picture(s) that were in the replaced range are not in the new content and have been deleted${track ? ' (as a tracked change the user can reject)' : ' (the user can undo this)'}; tell the user if that was not intended. To keep pictures, include their exact current <img> reference tags from read_blocks in the new HTML.`
          : ''
      return {
        output: `Replaced blocks ${range.start}-${range.end} with ${nodes.length} block(s). Compatible paragraph/text formatting was inherited; this does not guarantee preservation of section boundaries, fields or other document structures.${structureNote}${pictureNote} Block indexes have changed; use get_document_context if needed.`,
        mutated: true,
        summary: t('aiSumReplacedBlocks', { start: range.start, end: range.end }),
      }
    }

    case 'replace_selection': {
      const html = String(call.input.html ?? '')
      const echo = toolEchoError(html) ?? contextMarkerError(html)
      if (echo) return fail(t('aiSumReplaceSelection'), echo)
      const range = selectionRange(editor, scope)
      if (!range) {
        return fail(
          t('aiSumReplaceSelection'),
          'nothing is selected: replace_selection needs a range selection (the <sel>…</sel> span in the context); to rewrite whole blocks use replace_blocks',
        )
      }
      const doc = editor.state.doc
      const $from = doc.resolve(range.from)
      const $to = doc.resolve(range.to)
      if (!$from.sameParent($to) || !$from.parent.isTextblock) {
        return fail(
          t('aiSumReplaceSelection'),
          'the selection spans several paragraphs or cells, or sits in protected content (image, field); replace_selection only edits text inside one paragraph/heading/list item, including a table cell — use replace_blocks for whole blocks',
        )
      }
      const liveRanges = liveInlineRanges(doc, range.from, range.to)
      if (liveRanges.length === 0) {
        return fail(
          t('aiSumReplaceSelection'),
          'the selection is a pending tracked deletion — it is already deleted; the user accepts/rejects revisions in the Review tab',
        )
      }
      let inline: ReturnType<typeof parseInlineFragment>
      let keptImages: Array<{ block: number; index: number }>
      try {
        inline = parseInlineFragment(html)
        const resolved = resolveImageRefs(editor, inline, html)
        inline = resolved.nodes
        keptImages = resolved.kept
      } catch (e) {
        return fail(t('aiSumReplaceSelection'), e instanceof Error ? e.message : String(e))
      }
      if (inline.length === 0) {
        return fail(t('aiSumReplaceSelection'), 'html did not parse into any text')
      }
      const structureError = validateStructureDeletion(
        editor.state.doc,
        range.from,
        range.to,
        call.input.deleteStructures,
        preservedImageStructures(editor, keptImages),
      )
      if (structureError) return fail(t('aiSumReplaceSelection'), structureError)
      const schema = editor.schema
      // a carried-over picture's own marks are not the model restyling the selection
      const fragmentStyled = inline.some((n) => n.type === 'text' && (n.marks ?? []).length > 0)
      const $liveFrom = doc.resolve(liveRanges[0].from)
      const anchor = $liveFrom.parent.childAfter($liveFrom.parentOffset).node
      const inherited = (anchor?.marks ?? []).filter(
        (m) =>
          !REVISION_MARK_TYPES.has(m.type.name) &&
          !['instrField', 'refField', 'ctrlCheckbox', 'docSym'].includes(m.type.name) &&
          !(fragmentStyled && FRAGMENT_MARK_TYPES.has(m.type.name)),
      )
      const nodes = inline.map((n) => {
        if (n.type !== 'text') return schema.nodeFromJSON(n)
        let marks: readonly Mark[] = inherited
        for (const m of n.marks ?? []) marks = schema.markFromJSON(m).addToSet(marks)
        return schema.text(n.text ?? '', marks)
      })
      const blockIndex = $from.index(0)
      // the echo quotes what was replaced back to the model, so a selection
      // covering a withheld span would undo the withholding right here
      const oldText = liveTextBetween(editor, range.from, range.to)
      replaceInlineRange(editor, range.from, range.to, nodes, track)
      const newText = nodes.map((n) => n.textContent).join('')
      return {
        output: `Replaced the selected text in block ${blockIndex}: "${clipText(oldText, 200)}" → "${clipText(newText, 200)}". The rest of the block is unchanged; the new text is now selected.`,
        mutated: true,
        summary: t('aiSumReplaceSelection'),
      }
    }

    case 'insert_chart': {
      const kind = String(call.input.kind ?? '') as NewChart['kind']
      if (!['bar', 'line', 'pie'].includes(kind))
        return fail(t('aiSumInsertChart'), 'kind must be one of bar/line/pie')
      const categories = Array.isArray(call.input.categories)
        ? (call.input.categories as unknown[]).map((c) => String(c ?? ''))
        : []
      const seriesIn = Array.isArray(call.input.series)
        ? (call.input.series as ReadonlyArray<{ name?: unknown; values?: unknown } | null>)
        : []
      if (!categories.length || !seriesIn.length)
        return fail(t('aiSumInsertChart'), 'categories and series must not be empty')
      const series = seriesIn.map((s, i) => {
        const values: unknown[] = Array.isArray(s?.values) ? s.values : []
        return {
          name: String(s?.name ?? `Series ${i + 1}`),
          values: categories.map((_, j) => {
            const v = values[j] ?? null
            const n = Number(v)
            return v != null && Number.isFinite(n) ? n : null
          }),
        }
      })
      const title = String(call.input.title ?? '').trim() || 'Chart title'
      const spec: NewChart = { kind, title, categories, series }
      const display: ChartDisplay = { partPath: '', kind, title, categories, series }
      const count = editor.state.doc.childCount
      const after =
        call.input.afterBlockIndex === undefined
          ? getCursorBlockIndex(editor, scope)
          : Math.min(Math.max(Number(call.input.afterBlockIndex), -1), count - 1)
      if (!Number.isInteger(after)) return fail(t('aiSumInsertChart'), 'invalid afterBlockIndex')
      const { to } = blockRangePositions(editor, after, after)
      editor
        .chain()
        .insertContentAt(to, {
          type: 'docProtected',
          attrs: {
            docxIndex: null,
            blockType: 'chart',
            label: 'Chart',
            genChart: spec,
            chartDisplay: display,
          },
        })
        .run()
      return {
        output: `Inserted a ${kind} chart "${title}" (${categories.length} categories × ${series.length} series).`,
        mutated: true,
        summary: t('aiSumInsertedChart', { title }),
      }
    }

    case 'edit_chart': {
      const idx = Number(call.input.blockIndex)
      if (!Number.isInteger(idx) || idx < 0 || idx >= editor.state.doc.childCount) {
        return fail(t('aiSumEditChart'), 'blockIndex invalid or out of range')
      }
      const node = editor.state.doc.child(idx)
      // native docx charts are passthrough blocks + chartDisplay; AI/UI-created ones are blockType 'chart'
      const display =
        node.type.name === 'docProtected' ? (node.attrs.chartDisplay as ChartDisplay | null) : null
      if (!display)
        return fail(
          t('aiSumEditChart'),
          `block ${idx} is not a chart (or the chart has no editable data cache)`,
        )
      const isNative = display.partPath !== '' // native chart part from docx: data-point structure is immutable
      const next: ChartDisplay = {
        ...display,
        categories: [...display.categories],
        series: display.series.map((s) => ({ ...s, values: [...s.values] })),
      }
      if (call.input.title !== undefined) next.title = String(call.input.title)
      if (call.input.categories !== undefined) {
        const cats = call.input.categories
        if (!Array.isArray(cats) || cats.length !== display.categories.length) {
          return fail(
            t('aiSumEditChart'),
            `categories must match the original category count (${display.categories.length})`,
          )
        }
        cats.forEach((c, i) => {
          if (c != null) next.categories[i] = String(c)
        })
      }
      const serIn = Array.isArray(call.input.series)
        ? (call.input.series as ReadonlyArray<{
            index?: unknown
            name?: unknown
            values?: unknown
          } | null>)
        : []
      for (const s of serIn) {
        const si = Number(s?.index)
        const orig = display.series[si]
        if (!Number.isInteger(si) || !orig)
          return fail(
            t('aiSumEditChart'),
            `series index ${s?.index} is invalid (${display.series.length} series in total)`,
          )
        if (s?.name !== undefined) next.series[si]!.name = String(s.name)
        if (s?.values !== undefined) {
          const values: unknown[] | null = Array.isArray(s.values) ? s.values : null
          if (!values || values.length !== orig.values.length) {
            return fail(
              t('aiSumEditChart'),
              `values of series ${si} must match the original series length (${orig.values.length})`,
            )
          }
          for (let j = 0; j < values.length; j++) {
            const v = values[j]
            if (v == null) continue
            const n = Number(v)
            if (!Number.isFinite(n))
              return fail(t('aiSumEditChart'), `value ${j} of series ${si} is not a number`)
            if (isNative && orig.values[j] == null) {
              return fail(
                t('aiSumEditChart'),
                `data point ${j} of series ${si} is empty in the original chart; empty points of a native chart cannot be written`,
              )
            }
            next.series[si]!.values[j] = n
          }
        }
      }
      // generated charts (genChart) update the spec in sync; the chart part is rebuilt from the new data on save
      const gen = node.attrs.genChart as NewChart | null
      const nextGen: NewChart | null = gen
        ? {
            ...gen,
            title: next.title ?? gen.title,
            categories: [...next.categories],
            series: next.series.map((s) => ({ name: s.name ?? '', values: [...s.values] })),
          }
        : null
      const { from } = blockRangePositions(editor, idx, idx)
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(from, undefined, {
          ...node.attrs,
          chartDisplay: next,
          genChart: nextGen,
        }),
      )
      return {
        output: `Updated the data of chart "${next.title ?? ''}" (changes are written back to the chart on save).`,
        mutated: true,
        summary: t('aiSumEditedChart', { index: idx }),
      }
    }

    case 'read_revisions':
      return {
        output: buildRevisionsContext(editor, extras?.storyRevisions?.list()),
        mutated: false,
        summary: t('aiSumReadRevisions'),
      }

    case 'accept_changes':
    case 'reject_changes': {
      const mode = call.name === 'accept_changes' ? 'accept' : 'reject'
      const summary = t(mode === 'accept' ? 'aiSumAcceptChanges' : 'aiSumRejectChanges')
      if (extras?.revisionsLocked)
        return fail(
          summary,
          'document protection enforces tracked changes (or read-only editing): revisions cannot be accepted or rejected until the protection is removed',
        )
      const sel = validateSelector(call.input)
      if ('error' in sel) return fail(summary, sel.error)
      const r = applyRevisionSelection(editor, sel, mode, extras?.storyRevisions)
      if ('error' in r) return fail(summary, r.error)
      const left = listRevisionEntries(editor.state.doc, extras?.storyRevisions?.list())
      const verb = mode === 'accept' ? 'Accepted' : 'Rejected'
      return {
        output: `${verb} ${r.entries.length} tracked change(s) (${r.entries.map((e) => e.id).join(', ')}). ${left.length ? `Still ${describePending(left)}.` : 'No tracked changes remain.'}`,
        mutated: true,
        summary,
      }
    }

    case 'read_comments': {
      if (!comments) return fail(t('aiSumReadComments'), 'comments are not available here')
      return {
        output: buildCommentsContext(editor, comments.list(), true),
        mutated: false,
        summary: t('aiSumReadComments'),
      }
    }

    case 'reply_comment': {
      if (!comments) return fail(t('aiSumReplyComment'), 'comments are not available here')
      const parentId = String(call.input.parentId ?? '').trim()
      const text = String(call.input.text ?? '').trim()
      if (!parentId || !text) {
        return fail(t('aiSumReplyComment'), 'parentId and text must not be empty')
      }
      const target = comments.list().find((c) => c.id === parentId)
      if (!target) {
        return fail(
          t('aiSumReplyComment'),
          `no comment with id ${parentId}; call read_comments for the current ids`,
        )
      }
      const rootId = target.parentId ?? target.id // replies always attach to the thread root
      if (!comments.reply(rootId, text)) {
        return fail(
          t('aiSumReplyComment'),
          'the comment anchor no longer exists in the document; the reply was not added',
        )
      }
      return {
        output: `Replied to comment ${rootId}.`,
        mutated: true, // the reply id joins the anchor marks, so the doc changed
        summary: t('aiSumReplyComment'),
      }
    }

    case 'resolve_comment': {
      if (!comments) return fail(t('aiSumResolveComment'), 'comments are not available here')
      const id = String(call.input.id ?? '').trim()
      const target = comments.list().find((c) => c.id === id)
      if (!target) {
        return fail(
          t('aiSumResolveComment'),
          `no comment with id ${id}; call read_comments for the current ids`,
        )
      }
      const rootId = target.parentId ?? target.id
      if (!comments.resolve(rootId)) {
        return fail(t('aiSumResolveComment'), `comment ${rootId} could not be resolved`)
      }
      return {
        output: `Comment ${rootId} marked as resolved.`,
        mutated: false, // app state only; the document content is untouched
        summary: t('aiSumResolveComment'),
      }
    }

    case 'insert_footnote':
    case 'insert_endnote': {
      const kind: NoteKind = call.name === 'insert_footnote' ? 'footnote' : 'endnote'
      const summary = t(kind === 'footnote' ? 'aiSumInsertFootnote' : 'aiSumInsertEndnote')
      if (!notes) return fail(summary, 'footnotes and endnotes are not available here')
      const text = String(call.input.text ?? '').trim()
      if (!text) return fail(summary, 'text must not be empty')
      const afterText =
        call.input.afterText === undefined || call.input.afterText === null
          ? undefined
          : String(call.input.afterText)
      if (afterText === '') return fail(summary, 'afterText must not be empty when given')
      const where = noteInsertPos(editor.state.doc, Number(call.input.blockIndex), afterText)
      if ('error' in where) return fail(summary, where.error)
      const id = notes.add(kind, text)
      insertNoteRef(editor, kind, id, where.pos)
      return {
        output: `Inserted ${kind} ${id} in block ${call.input.blockIndex}.`,
        mutated: true,
        summary,
      }
    }

    case 'delete_note': {
      const summary = t('aiSumDeleteNote')
      if (!notes) return fail(summary, 'footnotes and endnotes are not available here')
      const kind = call.input.kind
      if (kind !== 'footnote' && kind !== 'endnote')
        return fail(summary, 'kind must be "footnote" or "endnote"')
      const id = String(call.input.id ?? '').trim()
      if (!notes.list(kind).some((n) => n.id === id))
        return fail(summary, `no ${kind} with id ${id}; call read_notes for the current ids`)
      const locked = notes.protectedMarkBlock?.(kind, id)
      if (locked !== null && locked !== undefined) {
        return fail(
          summary,
          `the reference mark of ${kind} ${id} sits in protected block ${locked} (a paragraph the editor cannot rewrite); the note was left in place`,
        )
      }
      const removed = removeNoteRefs(editor, kind, id)
      notes.remove(kind, id)
      return {
        output: `Deleted ${kind} ${id}${removed ? ' and its reference mark' : ' (it had no reference mark in the text)'}.`,
        mutated: true,
        summary,
      }
    }

    case 'edit_note': {
      const summary = t('aiSumEditNote')
      if (!notes) return fail(summary, 'footnotes and endnotes are not available here')
      const id = String(call.input.id ?? '').trim()
      if (!id) return fail(summary, 'id must not be empty')
      const kinds: NoteKind[] =
        call.input.kind === undefined || call.input.kind === null
          ? ['footnote', 'endnote']
          : call.input.kind === 'footnote' || call.input.kind === 'endnote'
            ? [call.input.kind]
            : []
      if (kinds.length === 0) return fail(summary, 'kind must be "footnote" or "endnote"')
      const found = kinds
        .map((kind) => ({ kind, note: notes.list(kind).find((n) => n.id === id) }))
        .filter((hit): hit is { kind: NoteKind; note: NoteInfo } => hit.note !== undefined)
      if (found.length === 0)
        return fail(summary, `no note with id ${id}; call read_notes for the current ids`)
      if (found.length > 1)
        return fail(summary, `both a footnote and an endnote have id ${id}; pass kind`)
      const { kind, note } = found[0]!
      const find = typeof call.input.find === 'string' ? call.input.find : ''
      if (!find) return fail(summary, 'find must not be empty')
      const replace = call.input.replace == null ? '' : String(call.input.replace)
      const edited = editNoteText(note, find, replace, call.input.matchCase !== false)
      if (edited.count === 0)
        return fail(
          summary,
          `${kind} ${id} does not contain "${clipExcerpt(find)}"; the note was left as is`,
        )
      notes.replace(kind, id, edited.note)
      return {
        output: `Edited ${kind} ${id}: ${edited.count} occurrence${edited.count === 1 ? '' : 's'} of "${clipExcerpt(find)}" replaced.`,
        mutated: true,
        summary,
      }
    }

    case 'read_notes': {
      const summary = t('aiSumReadNotes')
      if (!notes) return fail(summary, 'footnotes and endnotes are not available here')
      return {
        output: buildNotesContext(
          editor.state.doc,
          notes.list('footnote'),
          notes.list('endnote'),
          (kind, id) => notes.protectedMarkBlock?.(kind, id) ?? null,
        ),
        mutated: false,
        summary,
      }
    }

    case 'add_comment': {
      if (!comments?.add) return fail(t('aiSumAddComment'), 'comments are not available here')
      const body = typeof call.input.comment === 'string' ? call.input.comment.trim() : ''
      if (!body) return fail(t('aiSumAddComment'), 'comment must not be empty')
      const anchor = resolveCommentAnchor(editor, {
        blockIndex: call.input.blockIndex,
        text: call.input.text,
        occurrence: call.input.occurrence,
      })
      if ('error' in anchor) return fail(t('aiSumAddComment'), anchor.error)
      const meta: CommentAuthorMeta = {}
      if (typeof call.input.author === 'string' && call.input.author.trim()) {
        meta.author = call.input.author.trim()
      }
      if (typeof call.input.initials === 'string' && call.input.initials.trim()) {
        meta.initials = call.input.initials.trim()
      }
      const id = comments.add({ from: anchor.from, to: anchor.to }, body, meta)
      if (!id) return fail(t('aiSumAddComment'), 'the range holds no text to anchor a comment to')
      return {
        output: `Added comment ${id} on block ${call.input.blockIndex}, anchored to "${clipExcerpt(anchor.excerpt)}".`,
        mutated: true,
        summary: t('aiSumAddComment'),
      }
    }

    case 'delete_comment': {
      if (!comments?.remove) return fail(t('aiSumDeleteComment'), 'comments are not available here')
      const id = String(call.input.id ?? '').trim()
      const list = comments.list()
      const target = list.find((c) => c.id === id)
      if (!target) {
        return fail(
          t('aiSumDeleteComment'),
          `no comment with id ${id}; call read_comments for the current ids`,
        )
      }
      const replies = repliesOf(list, id)
      if (replies.length > 0 && call.input.withReplies !== true) {
        return fail(
          t('aiSumDeleteComment'),
          `comment ${id} has ${replies.length} repl${replies.length === 1 ? 'y' : 'ies'} (${replies.map((r) => r.id).join(', ')}); set withReplies: true to delete the whole thread`,
        )
      }
      if (!comments.remove(id))
        return fail(t('aiSumDeleteComment'), `comment ${id} could not be deleted`)
      const gone = [id, ...replies.map((r) => r.id)]
      return {
        output: `Deleted comment${gone.length > 1 ? 's' : ''} ${gone.join(', ')}.`,
        mutated: true,
        summary: t('aiSumDeleteComment'),
      }
    }

    case 'set_header_footer': {
      const kind = String(call.input.kind ?? '')
      const summaryOf = () => t(kind === 'footer' ? 'aiSumSetFooter' : 'aiSumSetHeader')
      if (!hf) return fail(summaryOf(), 'header/footer editing is not available here')
      if (kind !== 'header' && kind !== 'footer') {
        return fail(summaryOf(), 'kind must be "header" or "footer"')
      }
      const view = call.input.view === undefined ? 'default' : String(call.input.view)
      if (view !== 'default' && view !== 'first' && view !== 'even') {
        return fail(summaryOf(), 'view must be "default", "first" or "even"')
      }
      if (typeof call.input.text !== 'string') return fail(summaryOf(), 'text must be a string')
      const section = call.input.section === undefined ? 0 : Number(call.input.section)
      const snapshot = hf.read()
      if (snapshot.multiSection && !snapshot.sections)
        return fail(
          summaryOf(),
          'This host does not expose section-addressed header/footer editing.',
          'unavailable',
        )
      const count = snapshot.sections?.length ?? 1
      if (
        (call.input.section !== undefined && !Number.isInteger(call.input.section)) ||
        !Number.isInteger(section) ||
        section < 0 ||
        section >= count
      )
        return fail(summaryOf(), `section must be 0-${count - 1}`)
      const scope = call.input.scope ?? 'section'
      if (scope !== 'section' && scope !== 'shared')
        return fail(summaryOf(), 'scope must be section or shared')
      const strip = snapshot.sections?.[section]?.strips.find(
        (s) => s.kind === kind && s.view === view,
      )
      const affected =
        scope === 'shared' ? (strip?.sharedTargets ?? [{ section, view }]) : [{ section, view }]
      // a doubly escaped newline arrives as the two characters backslash-n; a footer must not show it literally
      const text = call.input.text.includes('\n')
        ? call.input.text
        : call.input.text.replace(/\\n/g, '\n')
      if (text.length > 2000) {
        return fail(summaryOf(), 'text is too long for a header/footer (2000 characters max)')
      }
      const textboxIndex = call.input.textboxIndex
      if (
        textboxIndex !== undefined &&
        (!Number.isInteger(textboxIndex) || Number(textboxIndex) < 0)
      )
        return fail(summaryOf(), 'textboxIndex must be a non-negative integer from current context')
      const error = hf.set(
        kind,
        view,
        text,
        section,
        scope,
        textboxIndex === undefined ? undefined : Number(textboxIndex),
      )
      if (error) return fail(summaryOf(), error)
      return {
        output: `Updated section ${section} ${kind}${view !== 'default' ? ` (${view}-page variant)` : ''}${textboxIndex === undefined ? '' : ` textbox ${textboxIndex}`}; scope=${scope}; affected targets: ${JSON.stringify(affected)}. ${scope === 'section' ? 'Other sections retain their content; shared/inherited parts are isolated on save.' : 'All listed aliases of the owning part are affected.'}${view === 'even' ? ' Even/odd display is enabled document-wide.' : ''}`,
        mutated: false, // app state only, saved with the document; not part of the PM doc
        summary: summaryOf(),
      }
    }

    case 'set_page_setup': {
      const summary = t('aiSumSetPageSetup')
      if (!pageSetup) return fail(summary, 'page setup is not available here')
      const sections = pageSetup.list()
      if (sections.length === 0) return fail(summary, 'the document has no sections yet')
      const { section, blockIndex, ...fields } = call.input
      let targets: number[]
      if (section !== undefined && section !== null) {
        if (!Number.isInteger(section) || Number(section) < 0 || Number(section) >= sections.length)
          return fail(summary, `section must be 0-${sections.length - 1}`)
        targets = [Number(section)]
      } else if (blockIndex !== undefined && blockIndex !== null) {
        if (
          !Number.isInteger(blockIndex) ||
          Number(blockIndex) < 0 ||
          Number(blockIndex) >= editor.state.doc.childCount
        )
          return fail(summary, rangeError(editor))
        targets = [sectionIndexOfBlock(sections, Number(blockIndex))]
      } else {
        targets = sections.map((s) => s.index)
      }
      const lines: string[] = []
      for (const index of targets) {
        const current = pageSetup.current(index)
        if (!current) return fail(summary, `section ${index} is not available`)
        const resolved = resolvePageSetup(fields, current)
        if ('error' in resolved) return fail(summary, resolved.error)
        const error = pageSetup.set(index, resolved)
        if (error) return fail(summary, error)
      }
      for (const s of pageSetup.list()) {
        if (targets.includes(s.index))
          lines.push(
            `section ${s.index} (blocks ${s.firstBlock}-${s.lastBlock}): ${sectionLine(s)}`,
          )
      }
      return {
        output: `Page setup updated.\n${lines.join('\n')}`,
        mutated: false, // section settings live beside the document and are written on save
        summary,
      }
    }

    case 'insert_section_break': {
      const summary = t('aiSumInsertSectionBreak')
      if (!pageSetup) return fail(summary, 'section breaks are not available here')
      const after = call.input.afterBlockIndex
      const count = editor.state.doc.childCount
      if (!Number.isInteger(after) || Number(after) < -1 || Number(after) >= count)
        return fail(summary, rangeError(editor))
      const type = call.input.type === undefined ? 'nextPage' : call.input.type
      if (!SECTION_BREAK_TYPES.includes(type as SectionBreakType))
        return fail(summary, `type must be one of ${SECTION_BREAK_TYPES.join(', ')}`)
      const error = pageSetup.insertBreak(type as SectionBreakType, Number(after))
      if (error) return fail(summary, error)
      return {
        output: `Inserted a ${type} section break after block ${after}; the break is block ${Number(after) + 1} and later indexes shifted by one. Sections now:\n${pageSetup
          .list()
          .map(
            (s) =>
              `- section ${s.index} (blocks ${s.firstBlock}-${s.lastBlock}): ${sectionLine(s)}`,
          )
          .join('\n')}`,
        mutated: true,
        summary,
      }
    }

    case 'list_styles': {
      const summary = t('aiSumListStyles')
      if (!extras?.styles) return fail(summary, 'the style catalog is not available here')
      return { output: describeStyles(extras.styles.list()), mutated: false, summary }
    }

    case 'define_style': {
      const summary = t('aiSumDefineStyle')
      if (!extras?.styles) return fail(summary, 'style definitions are not available here')
      const resolved = resolveStyleDefinition(call.input, extras.styles.list())
      if ('error' in resolved) return fail(summary, resolved.error)
      const error = extras.styles.upsert(resolved.value.upsert)
      if (error) return fail(summary, error)
      const { styleId } = resolved.value.upsert
      return {
        output: `${resolved.value.existing ? 'Updated' : 'Created'} style ${styleId}; apply it with the applyStyle op { op: "applyStyle", target, styleId: "${styleId}" }.`,
        mutated: false, // styles.xml is written on save; the PM doc is untouched
        summary,
      }
    }

    case 'set_watermark': {
      const summary = t('aiSumSetWatermark')
      if (!extras?.watermark) return fail(summary, 'the watermark is not available here')
      const resolved = resolveWatermark(call.input)
      if ('error' in resolved) return fail(summary, resolved.error)
      const error = extras.watermark.set(resolved.spec)
      if (error) return fail(summary, error)
      return {
        output: resolved.spec ? `Watermark set to "${resolved.spec.text}".` : 'Watermark removed.',
        mutated: false, // header part state, written on save
        summary,
      }
    }

    case 'insert_text_box': {
      const summary = t('aiSumInsertTextBox')
      const at = insertPosition(editor, call.input.afterBlockIndex)
      if ('error' in at) return fail(summary, at.error)
      const built = textBoxNode(call.input)
      if ('error' in built) return fail(summary, built.error)
      editor.chain().insertContentAt(at.pos, built.node).run()
      return {
        output: `Inserted a ${built.widthPx}x${built.heightPx}px text box after block ${at.after} (it is block ${at.after + 1}; later indexes shifted by one).`,
        mutated: true,
        summary,
      }
    }

    case 'apply_ops': {
      const dryRun = call.input.dryRun === true
      const outcome = executeOps(editor, call.input.ops, {
        numIds,
        track,
        selection: scope,
        dryRun,
        styles: extras?.styles,
      })
      if (!outcome.ok) return fail(t('aiSumApplyCommands'), outcome.error ?? 'op execution failed')
      if (dryRun) {
        return {
          output: `Dry run: ${outcome.plan?.length ?? 0} op(s) valid, nothing applied.\n${(outcome.plan ?? []).join('\n')}`,
          mutated: false,
          summary: t('aiSumApplyCommands'),
        }
      }
      const changed = outcome.results.reduce((sum, r) => sum + r.changed, 0)
      const exactCounts = outcome.results.every(
        (result) => result.matchCount !== undefined && result.changeCount !== undefined,
      )
      const skippedDeleted = outcome.results.reduce((sum, r) => sum + (r.skippedDeleted ?? 0), 0)
      // explicit model-facing note so it stops retrying deletions of already-deleted text
      const deletedNote =
        skippedDeleted > 0
          ? `\nNote: ${skippedDeleted} matched target(s) were skipped because that text is a pending tracked deletion (already struck through). It is not current content — do not try to delete or replace it again; the user accepts/rejects revisions in the Review tab.`
          : ''
      return {
        output: outcome.summary + deletedNote,
        mutated: changed > 0,
        committed: changed > 0,
        ...(exactCounts
          ? {
              matchCount: outcome.results.reduce((sum, result) => sum + result.matchCount!, 0),
              changeCount: outcome.results.reduce((sum, result) => sum + result.changeCount!, 0),
            }
          : {}),
        summary: outcome.summary,
      }
    }

    default:
      return fail(call.name, `unknown tool: ${call.name}`)
  }
}

/** top-level index of the block containing the caret (doc end as fallback);
 *  a frozen scope wins over the live caret — it is what the prompt described */
function getCursorBlockIndex(editor: Editor, scope?: SelectionScope | null): number {
  if (scope) return Math.min(Math.max(scope.endIndex, 0), editor.state.doc.childCount - 1)
  const { from } = editor.state.selection
  let result = editor.state.doc.childCount - 1
  let index = 0
  editor.state.doc.forEach((node, offset) => {
    if (from >= offset && from <= offset + node.nodeSize) result = index
    index++
  })
  return result
}
