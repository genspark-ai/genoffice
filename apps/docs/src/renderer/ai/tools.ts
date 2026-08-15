import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { ChartDisplay, NewChart } from '@genoffice/docx-engine'
import type { AgentToolCall, AgentToolDef } from '../../shared/ipc'
import { t } from '../i18n/locale'
import { executeCommands, type Command, type CommandEnvelope } from './commands'
import {
  blockRangePositions,
  buildDocumentContext,
  insertBlocksAfter,
  isBlankDocument,
  isTrackedDeleted,
  parseHtmlFragment,
  replaceBlockRange,
  serializeRangeToHtml,
  type AiTrack,
  type NumIds,
} from './protocol'

/**
 * Local agent tools: a document-context reader plus a script-style execution
 * channel. The "execution backend" is the in-process ProseMirror doc, split
 * into three safe primitives plus the deterministic command engine.
 */

const READ_MAX_CHARS = 24_000

export const AGENT_TOOLS: AgentToolDef[] = [
  {
    name: 'get_document_context',
    description:
      'Get the latest state of the current document: block list (index|type|content preview), full-text stats (word/character counts) and the current selection. Block indexes change after modifications; call this when you need up-to-date indexes.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'read_blocks',
    description:
      'Read the full content of a block range (restricted HTML). Previews in the block list are truncated; you must read the full original text with this tool before rewriting. ' +
      'Long ranges are paged: a truncated result says which offset to continue from; concatenate the slices in order to get the full HTML.',
    inputSchema: {
      type: 'object',
      properties: {
        startBlockIndex: { type: 'integer', description: 'start block index (0-based, inclusive)' },
        endBlockIndex: { type: 'integer', description: 'end block index (inclusive)' },
        offset: {
          type: 'integer',
          description:
            'character offset to continue a truncated read (default 0); use the offset given in the previous truncation notice',
        },
      },
      required: ['startBlockIndex', 'endBlockIndex'],
    },
  },
  {
    name: 'insert_content',
    description:
      'Insert new content at a given position (restricted HTML, may contain multiple blocks). For writing/continuing/generating new content; to rewrite existing content use replace_blocks.',
    inputSchema: {
      type: 'object',
      properties: {
        html: { type: 'string', description: 'restricted HTML fragment to insert' },
        afterBlockIndex: {
          type: 'integer',
          description:
            'insert after this block index; -1 = start of document; omitted = after the block containing the cursor',
        },
      },
      required: ['html'],
    },
  },
  {
    name: 'replace_blocks',
    description:
      'Replace a block range with new content (restricted HTML). For rewriting/translating/condensing/expanding existing content; the new block count may differ from the old.',
    inputSchema: {
      type: 'object',
      properties: {
        startBlockIndex: { type: 'integer', description: 'start block index (0-based, inclusive)' },
        endBlockIndex: { type: 'integer', description: 'end block index (inclusive)' },
        html: { type: 'string', description: 'replacement restricted HTML fragment' },
      },
      required: ['startBlockIndex', 'endBlockIndex', 'html'],
    },
  },
  {
    name: 'apply_commands',
    description:
      'Execute formatting/structure/batch commands (batchUpdate style, see the command guide in the system prompt): text style, paragraph format, heading level, find & replace, delete/move blocks, list conversion, image properties.',
    inputSchema: {
      type: 'object',
      properties: {
        commands: {
          type: 'array',
          description: 'array of commands executed in order; each command is a single-key object',
          items: { type: 'object' },
        },
      },
      required: ['commands'],
    },
  },
  {
    name: 'web_search',
    description:
      'Search the web for textual information (references/data/facts). Use when you need up-to-date information or are unsure about a fact. Returns titles/links/snippets.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'search keywords' },
        maxResults: { type: 'integer', description: 'maximum number of results, default 6' },
      },
      required: ['query'],
    },
  },
  {
    name: 'image_search',
    description:
      'Search for images. Returns a list of image imageUrl entries; after picking one, insert it into the document with insert_image.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'image search keywords (English works better)' },
        maxResults: { type: 'integer', description: 'maximum number of results, default 8' },
      },
      required: ['query'],
    },
  },
  {
    name: 'insert_image',
    description:
      'Download a direct image link (an imageUrl from image_search) and insert it into the document (at the cursor / end of document).',
    inputSchema: {
      type: 'object',
      properties: {
        url: { type: 'string', description: 'direct image link' },
        maxWidthPx: { type: 'integer', description: 'maximum width (px), default 480' },
      },
      required: ['url'],
    },
  },
  {
    name: 'insert_chart',
    description:
      'Insert a chart (saved as a native Word chart). Data must be real: from the document content or web_search results — do not make up numbers.',
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: ['bar', 'line', 'pie'], description: 'chart type' },
        title: { type: 'string', description: 'chart title' },
        categories: {
          type: 'array',
          items: { type: 'string' },
          description: 'category (x axis / sector) labels',
        },
        series: {
          type: 'array',
          description:
            'data series; values has the same length as categories, use null for missing data. Pie charts use only the first series',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string' },
              values: { type: 'array', items: { type: ['number', 'null'] } },
            },
            required: ['values'],
          },
        },
        afterBlockIndex: {
          type: 'integer',
          description:
            'insert after this block index; -1 = start of document; omitted = after the block containing the cursor',
        },
      },
      required: ['kind', 'categories', 'series'],
    },
  },
  {
    name: 'edit_chart',
    description:
      'Edit the data of an existing chart in the document (title/category labels/series names/values). Chart blocks in the block list can be edited; ' +
      'the category count and the number of values per series must match the original chart (data points cannot be added or removed).',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', description: 'block index of the chart' },
        title: { type: 'string', description: 'new title (omit to keep)' },
        categories: {
          type: 'array',
          items: { type: ['string', 'null'] },
          description:
            'new category labels, same length as the original; pass null for positions to keep',
        },
        series: {
          type: 'array',
          description: 'series to change',
          items: {
            type: 'object',
            properties: {
              index: { type: 'integer', description: 'series index (0-based)' },
              name: { type: 'string', description: 'new series name (omit to keep)' },
              values: {
                type: 'array',
                items: { type: ['number', 'null'] },
                description:
                  'new values, same length as the original series; pass null for positions to keep',
              },
            },
            required: ['index'],
          },
        },
      },
      required: ['blockIndex'],
    },
  },
]

export interface ToolExecution {
  /** result text fed back to the model */
  output: string
  isError?: boolean
  /** true when the tool changed the document */
  mutated: boolean
  /** short human-readable label for the chat activity chip */
  summary: string
}

const fail = (summary: string, output: string): ToolExecution => ({
  output,
  isError: true,
  mutated: false,
  summary,
})

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

/** Doc as last seen by the AI pipeline (context build / read / own write); a differing doc means the user edited in between. */
const docBaseline = new WeakMap<Editor, ProseMirrorNode>()

export function markDocSeen(editor: Editor): void {
  docBaseline.set(editor, editor.state.doc)
}

function editedExternally(editor: Editor): boolean {
  const seen = docBaseline.get(editor)
  return seen !== undefined && seen !== editor.state.doc
}

/** tools addressing the document by block index: refused after an external edit until the model re-reads */
const INDEX_WRITE_SUMMARIES: Record<string, () => string> = {
  insert_content: () => t('aiSumInsertContent'),
  replace_blocks: () => t('aiSumReplaceContent'),
  apply_commands: () => t('aiSumApplyCommands'),
  insert_chart: () => t('aiSumInsertChart'),
  edit_chart: () => t('aiSumEditChart'),
}

const STALE_DOC_ERROR =
  'The document was edited by the user since it was last read; block indexes may be stale. ' +
  'Call get_document_context (or read_blocks) to get the current state, then retry.'

function rangeError(editor: Editor): string {
  return `block index invalid or out of range (the document has ${editor.state.doc.childCount} blocks); call get_document_context for fresh indexes`
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
): Promise<ToolExecution> {
  switch (call.name) {
    case 'web_search': {
      const query = String(call.input.query ?? '').trim()
      if (!query) return fail(t('aiSumWebSearch'), 'query must not be empty')
      const r = await window.desktop.webSearch(query, Number(call.input.maxResults) || 6)
      // a backend failure must not read as "no results" — the model would fabricate conclusions
      if (r.method === 'error') {
        return fail(
          t('aiSumWebSearch'),
          `web search failed (service error, not an empty result — you may retry): ${r.error ?? 'unknown error'}`,
        )
      }
      const lines: string[] = []
      if (r.answer) lines.push(`Direct answer: ${r.answer}\n`)
      r.results.forEach((it, i) =>
        lines.push(`${i + 1}. ${it.title}\n   ${it.url}\n   ${it.snippet}`),
      )
      return {
        output: lines.join('\n') || '(no results)',
        mutated: false,
        summary: t('aiSumWebSearchDone', { query, count: r.results.length }),
      }
    }
    case 'image_search': {
      const query = String(call.input.query ?? '').trim()
      if (!query) return fail(t('aiSumImageSearch'), 'query must not be empty')
      const r = await window.desktop.imageSearch(query, Number(call.input.maxResults) || 8)
      // a backend failure must not read as an empty gallery — the model would fabricate image choices
      if (r.method === 'error') {
        return fail(
          t('aiSumImageSearch'),
          `image search failed (service error, not an empty result — you may retry): ${r.error ?? 'unknown error'}`,
        )
      }
      const lines = r.images.map(
        (im, i) =>
          `${i + 1}. ${im.title || '(untitled)'} [${im.width ?? '?'}x${im.height ?? '?'}]\n   ${im.imageUrl}`,
      )
      return {
        output: lines.join('\n') || '(no images)',
        mutated: false,
        summary: t('aiSumImageSearchDone', { query, count: r.images.length }),
      }
    }
    case 'insert_image': {
      const url = String(call.input.url ?? '')
      if (!/^https?:\/\//.test(url)) return fail(t('aiSumInsertImage'), 'invalid url')
      const fetched = await window.desktop.fetchImage(url)
      // never write after the user hit stop (the download may resolve long after the abort)
      if (signal?.aborted)
        return fail(t('aiSumInsertImage'), 'stopped by the user; the image was not inserted')
      if (!fetched)
        return fail(t('aiSumInsertImage'), 'download failed (the image may not be accessible)')
      const dataUrl = `data:${fetched.mime};base64,${fetched.base64}`
      const maxW = Number(call.input.maxWidthPx) || 480
      try {
        const natural = await imageSizeOf(dataUrl)
        if (signal?.aborted)
          return fail(t('aiSumInsertImage'), 'stopped by the user; the image was not inserted')
        const scale = Math.min(1, maxW / natural.width)
        const w = Math.round(natural.width * scale)
        const h = Math.round(natural.height * scale)
        // The download can take long: user edits made meanwhile must keep the
        // freshness baseline stale, so only our own insertion may mark the doc
        // seen. Checked right before the write — there is no async gap after.
        const userEditedDuringFetch = editedExternally(editor)
        editor
          .chain()
          .focus()
          .insertContent({
            type: 'docProtected',
            attrs: {
              docxIndex: null,
              blockType: 'image',
              label: 'Image (web)',
              imageDataUrl: dataUrl,
              imageWidthPx: w,
              imageHeightPx: h,
              genImage: { base64: fetched.base64, mime: fetched.mime, widthPx: w, heightPx: h },
            },
          })
          .run()
        if (!userEditedDuringFetch) markDocSeen(editor)
        return {
          output: `Inserted the image (${w}×${h}px).`,
          mutated: true,
          summary: t('aiSumInsertWebImage'),
        }
      } catch {
        return fail(t('aiSumInsertImage'), 'the image could not be decoded')
      }
    }
    default:
      return fail(t('aiSumUnknownTool'), call.name)
  }
}

export function executeTool(
  editor: Editor,
  call: AgentToolCall,
  numIds: NumIds,
  track?: AiTrack,
  signal?: AbortSignal,
): ToolExecution | Promise<ToolExecution> {
  const staleSummary = INDEX_WRITE_SUMMARIES[call.name]
  if (staleSummary && editedExternally(editor)) return fail(staleSummary(), STALE_DOC_ERROR)
  const settle = (exec: ToolExecution): ToolExecution => {
    const readsDoc = call.name === 'get_document_context' || call.name === 'read_blocks'
    if (exec.mutated || (readsDoc && !exec.isError)) markDocSeen(editor)
    return exec
  }
  // Async tools take a separate Promise branch; the other sync tools keep returning
  // synchronously (doesn't break existing tests). No settle here: marking the doc
  // seen after the long download would baptize user edits made meanwhile —
  // insert_image maintains the baseline itself right at its synchronous write.
  if (call.name === 'web_search' || call.name === 'image_search' || call.name === 'insert_image') {
    return executeAsyncTool(editor, call, signal)
  }
  return settle(executeSyncTool(editor, call, numIds, track))
}

function executeSyncTool(
  editor: Editor,
  call: AgentToolCall,
  numIds: NumIds,
  track?: AiTrack,
): ToolExecution {
  switch (call.name) {
    case 'get_document_context':
      return {
        output: buildDocumentContext(editor),
        mutated: false,
        summary: t('aiSumReadDocContext'),
      }

    case 'read_blocks': {
      const range = validRange(editor, call.input.startBlockIndex, call.input.endBlockIndex)
      if (!range) return fail(t('aiSumReadBlocks'), rangeError(editor))
      const html = serializeRangeToHtml(editor, range.start, range.end)
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
      const html = String(call.input.html ?? '')
      const echo = toolEchoError(html)
      if (echo) return fail(t('aiSumInsertContent'), echo)
      let nodes: ReturnType<typeof parseHtmlFragment>
      try {
        nodes = parseHtmlFragment(html, numIds)
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
        ? getCursorBlockIndex(editor)
        : Math.min(Math.max(Number(call.input.afterBlockIndex), -1), count - 1)
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
      const echo = toolEchoError(html)
      if (echo) return fail(t('aiSumReplaceContent'), echo)
      let nodes: ReturnType<typeof parseHtmlFragment>
      try {
        nodes = parseHtmlFragment(html, numIds)
      } catch (e) {
        return fail(t('aiSumReplaceContent'), e instanceof Error ? e.message : String(e))
      }
      if (nodes.length === 0)
        return fail(t('aiSumReplaceContent'), 'html did not parse into any content blocks')
      replaceBlockRange(editor, range.start, range.end, nodes, track)
      return {
        output: `Replaced blocks ${range.start}-${range.end} with ${nodes.length} block(s). Block indexes have changed; use get_document_context if needed.`,
        mutated: true,
        summary: t('aiSumReplacedBlocks', { start: range.start, end: range.end }),
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
          ? getCursorBlockIndex(editor)
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

    case 'apply_commands': {
      const commands = call.input.commands
      if (!Array.isArray(commands) || commands.length === 0) {
        return fail(t('aiSumApplyCommands'), 'commands must be a non-empty array')
      }
      const envelope: CommandEnvelope = { commands: commands as Command[] }
      const outcome = executeCommands(editor, envelope, { numIds, track })
      if (!outcome.ok)
        return fail(t('aiSumApplyCommands'), outcome.error ?? 'command execution failed')
      const changed = outcome.results.reduce((sum, r) => sum + r.changed, 0)
      const skippedDeleted = outcome.results.reduce((sum, r) => sum + (r.skippedDeleted ?? 0), 0)
      // explicit model-facing note so it stops retrying deletions of already-deleted text
      const deletedNote =
        skippedDeleted > 0
          ? `\nNote: ${skippedDeleted} matched target(s) were skipped because that text is a pending tracked deletion (already struck through). It is not current content — do not try to delete or replace it again; the user accepts/rejects revisions in the Review tab.`
          : ''
      return {
        output: outcome.summary + deletedNote,
        mutated: changed > 0,
        summary: outcome.summary,
      }
    }

    default:
      return fail(call.name, `unknown tool: ${call.name}`)
  }
}

/** top-level index of the block containing the caret (doc end as fallback) */
function getCursorBlockIndex(editor: Editor): number {
  const { from } = editor.state.selection
  let result = editor.state.doc.childCount - 1
  let index = 0
  editor.state.doc.forEach((node, offset) => {
    if (from >= offset && from <= offset + node.nodeSize) result = index
    index++
  })
  return result
}
