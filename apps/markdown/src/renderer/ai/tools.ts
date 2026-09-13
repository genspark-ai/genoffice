import type { Editor, JSONContent } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { AgentToolCall, AgentToolDef, ToolExecution } from '@genoffice/agent-core'
import {
  OP_SPECS,
  blockIndexRange,
  buildOpsGuide,
  isBlankDoc,
  parseMarkdownToNodes,
  runOps,
  usesBlockIndexes,
  validateOps,
  type FrontmatterAccess,
  type MdOp,
} from '../editor/ops'
import { t } from '../i18n/locale'
import {
  DraftLanding,
  type AiDocWriter,
  type DocWriteResult,
  type WritePosition,
} from './doc-writer'

export type { FrontmatterAccess } from '../editor/ops'
export { blockIndexRange } from '../editor/ops'

const CONTEXT_MAX_CHARS = 8000
const PREVIEW_CHARS = 60
const READ_PAGE_CHARS = 24000
const SELECTION_MAX_CHARS = 4000

const INDEX_CHANGE_NOTICE =
  'Block indexes may have changed; call get_document_context before further index-based edits.'
const STALE_DOC_ERROR =
  'The document changed since you last saw it (the user edited it). Call get_document_context to refresh before editing.'

// ── staleness guard: index-addressed writes are refused after user edits ──

const docBaseline = new WeakMap<Editor, PmNode>()

export function markDocSeen(editor: Editor): void {
  docBaseline.set(editor, editor.state.doc)
}

function editedExternally(editor: Editor): boolean {
  const seen = docBaseline.get(editor)
  return seen !== undefined && seen !== editor.state.doc
}

// ── document skeleton / serialization helpers ──

function blockPreview(node: PmNode): string {
  const text = node.textContent.replace(/\s+/g, ' ').trim()
  return text.length > PREVIEW_CHARS ? `${text.slice(0, PREVIEW_CHARS)}…` : text
}

function blockLabel(node: PmNode): string {
  if (node.type.name === 'heading') return `h${node.attrs.level}`
  return node.type.name
}

/** Serialize a range of top-level blocks back to markdown */
function serializeBlocks(editor: Editor, from: number, to: number): string {
  const content: JSONContent[] = []
  editor.state.doc.forEach((node, _offset, index) => {
    if (index >= from && index <= to) content.push(node.toJSON() as JSONContent)
  })
  return editor.markdown?.serialize({ type: 'doc', content }) ?? ''
}

function selectionMarkdown(editor: Editor): string {
  const { from, to } = editor.state.selection
  if (from === to) return ''
  const text = editor.state.doc.textBetween(from, to, '\n')
  return text.length > SELECTION_MAX_CHARS ? `${text.slice(0, SELECTION_MAX_CHARS)}…` : text
}

/** Per-turn context: numbered block skeleton + selection, same shape as the docs agent */
export function buildDocContext(editor: Editor): string {
  const doc = editor.state.doc
  const blockCount = doc.childCount
  if (isBlankDoc(doc)) {
    return ['## Document state', 'The document is currently blank.'].join('\n')
  }
  const lines: string[] = ['## Document state', `${blockCount} top-level blocks:`, '']
  let used = 0
  for (let i = 0; i < blockCount; i++) {
    const node = doc.child(i)
    const line = `${i} | ${blockLabel(node)} | ${blockPreview(node)}`
    used += line.length + 1
    if (used > CONTEXT_MAX_CHARS) {
      lines.push(`… (${blockCount - i} more blocks; use read_blocks to view them)`)
      break
    }
    lines.push(line)
  }
  // report by range, not by text: a selected image/table has no text but the
  // model still needs to know which block the request targets
  const { from: selFrom, to: selTo } = editor.state.selection
  if (selFrom < selTo) {
    const { startIndex, endIndex } = blockIndexRange(doc, selFrom, selTo)
    const where =
      startIndex === endIndex ? `block ${startIndex}` : `blocks ${startIndex}-${endIndex}`
    const selection =
      selectionMarkdown(editor) ||
      `(a non-text block is selected: ${blockLabel(doc.child(startIndex))})`
    lines.push('', `## User selection (${where})`, selection)
  }
  return lines.join('\n')
}

// ── tool definitions ──

export const AGENT_TOOLS: AgentToolDef[] = [
  {
    name: 'get_document_context',
    description:
      'Refresh the document overview: a numbered list of top-level blocks (index | type | preview) plus the current selection. Call this before index-based edits when in doubt.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'read_blocks',
    description:
      'Read a range of top-level blocks as markdown. Long output is paged; a notice tells you the offset to continue from.',
    inputSchema: {
      type: 'object',
      properties: {
        startIndex: { type: 'integer', description: '0-based index of the first block' },
        endIndex: { type: 'integer', description: '0-based index of the last block (inclusive)' },
        offset: {
          type: 'integer',
          description: 'Character offset to continue a previously truncated read',
        },
      },
      required: ['startIndex', 'endIndex'],
    },
  },
  {
    name: 'apply_ops',
    description: `Edit the document with a batch of ops — the same operations the editor's own toolbar performs. Ops run in order and stop at the first failure; the result lists what each op did.\n\n${buildOpsGuide()}`,
    inputSchema: {
      type: 'object',
      properties: {
        ops: {
          type: 'array',
          description: 'Ordered list of op objects (see the op catalogue)',
          items: { type: 'object' },
        },
      },
      required: ['ops'],
    },
  },
  {
    name: 'write_document',
    description:
      '[For long new content: a whole document, a chapter, a full report, article or translation] Hands the writing to the system writer, which streams markdown straight into the document while the user watches; you never write the text yourself. Give a concrete plan (title, section outline with the key points of each, tone, target length) and put every fact, figure, name and quote the text must use into context — the writer sees only the plan and context, not the conversation. Omit afterIndex on a blank document; on a document with content, pass afterIndex to insert after that block, or replaceDocument=true when the user asked to rewrite everything. Short additions (a paragraph or two) use apply_ops insertContent instead.',
    inputSchema: {
      type: 'object',
      properties: {
        plan: {
          type: 'string',
          description: 'title, section outline with key points per section, tone, target length',
        },
        title: { type: 'string', description: 'document/chapter title, when there is one' },
        context: {
          type: 'string',
          description:
            'reference material: facts, figures, quotes, sources gathered from the conversation and web_search',
        },
        afterIndex: {
          type: 'integer',
          description:
            'insert after this block index (-1 = document start); omitted = whole document',
        },
        replaceDocument: {
          type: 'boolean',
          description:
            'true replaces all existing content (only when the user asked for a rewrite)',
        },
      },
      required: ['plan'],
    },
  },
  {
    name: 'image_search',
    description:
      'Search for images. Returns a list of imageUrl entries; after picking one, insert it into the document with insert_image.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search keywords' },
        maxResults: { type: 'integer', description: 'Maximum number of results, default 8' },
      },
      required: ['query'],
    },
  },
  {
    name: 'insert_image',
    description:
      'Download a direct image link (an imageUrl from image_search) and insert it into the document after a top-level block. The image file is saved next to the document, so the document must have been saved to disk at least once.',
    inputSchema: {
      type: 'object',
      properties: {
        url: { type: 'string', description: 'Direct image URL (http/https)' },
        afterIndex: {
          type: 'integer',
          description:
            '0-based block index to insert after; -1 = document start. Defaults to the end of the document.',
        },
        alt: { type: 'string', description: 'Alt text describing the image' },
      },
      required: ['url'],
    },
  },
  {
    name: 'generate_image',
    description:
      'Generate an illustration with AI from a text prompt and insert it into the document after a top-level block. For illustration/diagram-style art that image_search cannot find, or when the user asks to generate/draw a picture. The document must have been saved to disk at least once.',
    inputSchema: {
      type: 'object',
      properties: {
        prompt: {
          type: 'string',
          description: 'Image description in English, concrete and visual',
        },
        aspectRatio: {
          type: 'string',
          description: 'Optional aspect ratio like "16:9", "1:1", "4:3"',
        },
        afterIndex: {
          type: 'integer',
          description:
            '0-based block index to insert after; -1 = document start. Defaults to the end of the document.',
        },
        alt: { type: 'string', description: 'Alt text describing the image' },
      },
      required: ['prompt'],
    },
  },
  {
    name: 'read_frontmatter',
    description:
      'Read the document properties: the raw YAML frontmatter block at the top of the file (title, tags, date, …), without the --- fences. Write it back with the setFrontmatter op.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
]

// ── executor ──

function fail(output: string, summary: string): ToolExecution {
  return { output, isError: true, summary }
}

function clampIndex(value: unknown, max: number): number | null {
  const n = Number(value)
  if (!Number.isInteger(n) || n < 0 || n > max) return null
  return n
}

/** Activity-chip label: the op's own label when exactly one op ran, else the count */
function opsSummary(ops: MdOp[], applied: number): string {
  if (ops.length === 1 && applied === 1) return t(OP_SPECS[ops[0]!.op].labelKey)
  return t('aiToolApplyOpsDone', { n: applied })
}

function applyOps(editor: Editor, input: unknown, fm?: FrontmatterAccess): ToolExecution {
  const label = t('aiToolApplyOps')
  const parsed = validateOps(input)
  if ('error' in parsed) return fail(parsed.error, label)
  if (usesBlockIndexes(parsed.ops) && editedExternally(editor)) return fail(STALE_DOC_ERROR, label)
  const r = runOps(editor, parsed.ops, { source: 'ai', frontmatter: fm })
  const lines = r.results.map((res, i) =>
    res.ok
      ? `ops[${i}] ${parsed.ops[i]!.op}: ${res.message}`
      : `ops[${i}] ${parsed.ops[i]!.op} FAILED: ${res.error}`,
  )
  const skipped = parsed.ops.length - r.results.length
  if (skipped > 0) lines.push(`${skipped} later op(s) were not executed.`)
  if (r.blocksChanged) lines.push(INDEX_CHANGE_NOTICE)
  if (r.applied > 0) markDocSeen(editor)
  const failed = r.applied < parsed.ops.length
  return {
    output: lines.join('\n'),
    isError: failed || undefined,
    mutated: r.applied > 0,
    summary: opsSummary(parsed.ops, r.applied),
  }
}

// ── image insertion (insert_image / generate_image) ──

/** magic-byte sniff: the fetch handler's content-type mapping defaults unknown
 *  types to jpeg, which would save webp/svg bytes under a lying .jpg name */
function sniffImageExt(base64: string): string | null {
  let head: string
  try {
    head = atob(base64.slice(0, 12))
  } catch {
    return null
  }
  if (head.startsWith('\x89PNG')) return 'png'
  if (head.startsWith('GIF8')) return 'gif'
  if (head.charCodeAt(0) === 0xff && head.charCodeAt(1) === 0xd8) return 'jpg'
  return null
}

/** download a direct image URL, persist it beside the document, insert an image block */
async function insertImageFromUrl(
  editor: Editor,
  url: string,
  input: { afterIndex?: unknown; alt?: unknown },
  signal: AbortSignal | undefined,
  labels: { fail: string; done: string },
): Promise<ToolExecution> {
  const fetched = await window.markdownApi.fetchImage(url)
  // never write after the user hit stop (the download may resolve long after the abort)
  if (signal?.aborted) return fail('stopped by the user; the image was not inserted', labels.fail)
  if (!fetched) return fail('download failed (the image may not be accessible)', labels.fail)
  const ext = sniffImageExt(fetched.base64)
  if (!ext) {
    return fail(
      'unsupported image format (only png/jpg/gif can be embedded) — pick a different image',
      labels.fail,
    )
  }
  const rel = await window.markdownApi.saveImage({ base64: fetched.base64, ext })
  if (signal?.aborted) return fail('stopped by the user; the image was not inserted', labels.fail)
  if (!rel) {
    return fail(
      'the document has no saved location yet, so there is nowhere to store the image file — ask the user to save the document first, then retry',
      labels.fail,
    )
  }
  // downloads can take long: user edits made meanwhile must keep the freshness
  // baseline stale, so only our own insertion may mark the doc seen
  const userEditedDuringFetch = editedExternally(editor)
  const maxIndex = editor.state.doc.childCount - 1
  // typeof guard: Number(null) is 0, which would silently mean "after block 0"
  const afterRaw = input.afterIndex
  const after =
    typeof afterRaw === 'number' && Number.isInteger(afterRaw)
      ? Math.min(Math.max(afterRaw, -1), maxIndex)
      : maxIndex
  const r = runOps(editor, [{ op: 'insertImage', after, src: rel, alt: String(input.alt ?? '') }], {
    source: 'ai',
  })
  const res = r.results[0]!
  if (!res.ok) return fail(res.error, labels.fail)
  if (!userEditedDuringFetch) markDocSeen(editor)
  return {
    output: `Inserted the image (saved as ${rel}). ${INDEX_CHANGE_NOTICE}`,
    mutated: true,
    summary: labels.done,
  }
}

async function writeDocument(
  editor: Editor,
  call: AgentToolCall,
  signal: AbortSignal | undefined,
  writer: AiDocWriter | undefined,
): Promise<ToolExecution> {
  const label = t('aiToolWriteDoc')
  const plan = String(call.input.plan ?? '').trim()
  if (!plan) return fail('plan must not be empty', label)
  if (!writer) return fail('document writing is not available here', label)
  if (editedExternally(editor)) return fail(STALE_DOC_ERROR, label)
  const doc = editor.state.doc
  const afterRaw = call.input.afterIndex
  let position: WritePosition
  if (afterRaw !== undefined && afterRaw !== null) {
    if (!Number.isInteger(afterRaw) || Number(afterRaw) < -1 || Number(afterRaw) >= doc.childCount)
      return fail(`afterIndex out of range; the document has ${doc.childCount} blocks.`, label)
    position = { kind: 'after', index: Number(afterRaw) }
  } else if (isBlankDoc(doc) || call.input.replaceDocument === true) {
    position = { kind: 'whole' }
  } else {
    return fail(
      'the document is not blank: pass afterIndex to insert the new content after a block, or replaceDocument=true when the user asked to rewrite the whole document',
      label,
    )
  }
  const str = (v: unknown) => (v === undefined || v === null ? undefined : String(v))
  const draft = new DraftLanding(editor, position)
  let result: DocWriteResult
  let rendered: string | null
  try {
    result = await writer.write(
      { plan, title: str(call.input.title), context: str(call.input.context) },
      (markdown) => draft.update(markdown),
      signal,
    )
  } finally {
    rendered = draft.finish()
  }
  if (editor.isDestroyed) return fail('the document was closed', label)
  if (!result.ok || !result.markdown?.trim()) {
    return fail(
      `The writer produced nothing (${result.error ?? 'no output'}); the document is unchanged. Tell the user briefly and offer to try again.`,
      t('aiToolWriteDocFailed'),
    )
  }
  // a kept partial whose tail no longer parses lands what the user saw rendered
  let parses: boolean
  try {
    parses = parseMarkdownToNodes(editor, result.markdown).length > 0
  } catch {
    parses = false
  }
  const markdown = !parses && result.truncated && rendered ? rendered : result.markdown
  // an explicit rewrite replaces everything; text the user typed into a formerly blank
  // document while the draft streamed is kept, the content goes into the draft's slot
  const op: MdOp =
    position.kind === 'whole' &&
    !isBlankDoc(editor.state.doc) &&
    call.input.replaceDocument === true
      ? {
          op: 'replaceBlocks',
          target: { start: 0, end: editor.state.doc.childCount - 1 },
          markdown,
        }
      : { op: 'insertContent', after: draft.indexBefore(), markdown }
  const r = runOps(editor, [op], { source: 'ai' })
  const res = r.results[0]!
  if (!res.ok) return fail(res.error, t('aiToolWriteDocFailed'))
  markDocSeen(editor)
  const note = result.truncated
    ? ' The stream ended early, so the content is INCOMPLETE (the user chose to keep it): the tail is missing. Say so and offer to finish the missing sections with write_document (afterIndex at the end) or apply_ops insertContent.'
    : ''
  return {
    output: `Content written by the system (${result.markdown.length} chars). ${res.message} ${INDEX_CHANGE_NOTICE}${note}\nReply with one or two sentences describing what was written; do not paste the content.`,
    mutated: true,
    summary: result.truncated ? t('aiToolWriteDocPartial') : label,
  }
}

export function executeTool(
  editor: Editor,
  call: AgentToolCall,
  signal?: AbortSignal,
  fm?: FrontmatterAccess,
  writer?: AiDocWriter,
): ToolExecution | Promise<ToolExecution> {
  const doc = editor.state.doc
  const maxIndex = doc.childCount - 1

  switch (call.name) {
    case 'write_document':
      return writeDocument(editor, call, signal, writer)

    case 'read_frontmatter': {
      if (!fm) return fail('frontmatter is not available', t('aiToolReadFm'))
      const inner = fm.read()
      return {
        output: inner || '(the document has no frontmatter)',
        mutated: false,
        summary: t('aiToolReadFm'),
      }
    }

    case 'get_document_context': {
      markDocSeen(editor)
      return {
        output: buildDocContext(editor),
        mutated: false,
        summary: t('aiToolReadDoc'),
      }
    }

    case 'read_blocks': {
      const start = clampIndex(call.input.startIndex, maxIndex)
      const end = clampIndex(call.input.endIndex, maxIndex)
      if (start === null || end === null || start > end) {
        return fail(
          `Invalid block range; the document has ${doc.childCount} blocks.`,
          t('aiToolReadBlocks'),
        )
      }
      const full = serializeBlocks(editor, start, end)
      const offset = Math.max(0, Number(call.input.offset) || 0)
      const page = full.slice(offset, offset + READ_PAGE_CHARS)
      const truncated = offset + READ_PAGE_CHARS < full.length
      const notice = truncated
        ? `\n\n[truncated — continue with offset=${offset + READ_PAGE_CHARS}]`
        : ''
      return {
        output: page + notice,
        mutated: false,
        summary: t('aiToolReadBlocks'),
      }
    }

    case 'apply_ops':
      return applyOps(editor, call.input.ops, fm)

    case 'image_search': {
      const query = String(call.input.query ?? '').trim()
      if (!query) return fail('query must not be empty', t('aiToolImageSearch'))
      return window.markdownApi
        .imageSearch(query, Number(call.input.maxResults) || 8)
        .then((r): ToolExecution => {
          // a backend failure must not read as an empty gallery — the model would fabricate image choices
          if (r.method === 'error') {
            return fail(
              `image search failed (service error, not an empty result — you may retry): ${r.error ?? 'unknown error'}`,
              t('aiToolImageSearch'),
            )
          }
          const lines = r.images.map(
            (im, i) =>
              `${i + 1}. ${im.title || '(untitled)'} [${im.width ?? '?'}x${im.height ?? '?'}]\n   ${im.imageUrl}`,
          )
          return {
            output: lines.join('\n') || '(no images)',
            mutated: false,
            summary: t('aiToolImageSearchDone', { query, count: r.images.length }),
          }
        })
    }

    case 'insert_image': {
      if (editedExternally(editor)) return fail(STALE_DOC_ERROR, t('aiToolInsertImage'))
      const url = String(call.input.url ?? '')
      if (!/^https?:\/\//.test(url)) return fail('invalid url', t('aiToolInsertImage'))
      return insertImageFromUrl(editor, url, call.input, signal, {
        fail: t('aiToolInsertImage'),
        done: t('aiToolInsertImageDone'),
      })
    }

    case 'generate_image': {
      if (editedExternally(editor)) return fail(STALE_DOC_ERROR, t('aiToolGenImage'))
      const prompt = String(call.input.prompt ?? '').trim()
      if (!prompt) return fail('prompt must not be empty', t('aiToolGenImage'))
      const aspectRatio = String(call.input.aspectRatio ?? '').trim()
      return window.markdownApi
        .aiGenerateImage({ prompt, ...(aspectRatio ? { aspectRatio } : {}) })
        .then((generated): ToolExecution | Promise<ToolExecution> => {
          if (signal?.aborted) {
            return fail('stopped by the user; the image was not inserted', t('aiToolGenImage'))
          }
          if (!generated.url) {
            return fail(generated.error ?? 'image generation failed', t('aiToolGenImage'))
          }
          return insertImageFromUrl(editor, generated.url, call.input, signal, {
            fail: t('aiToolGenImage'),
            done: t('aiToolGenImageDone'),
          })
        })
    }

    default:
      return fail(`Unknown tool: ${call.name}`, call.name)
  }
}
