import type { Editor, JSONContent } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { AgentToolCall, AgentToolDef, ToolExecution } from '@genoffice/agent-core'
import { markAiRange, markAiRanges } from '../editor/aiHighlight'
import { stripLegacyFencedDivs } from '../markdown/docText'
import { t } from '../i18n/locale'

/** GFM marks the AI may apply to matched text (style_matches) */
const STYLABLE_MARKS = ['bold', 'italic', 'strike', 'code'] as const
type StylableMark = (typeof STYLABLE_MARKS)[number]

/** raw-text access to the YAML properties block (inner text, no --- fences) */
export interface FrontmatterAccess {
  read(): string
  /** replace the whole inner YAML; empty string removes the block */
  write(inner: string): void
}

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

/** top-level block index range covered by [from, to] (shared with the edit queue) */
export function blockIndexRange(
  doc: PmNode,
  from: number,
  to: number,
): { startIndex: number; endIndex: number } {
  let startIndex = -1
  let endIndex = -1
  let index = 0
  doc.forEach((node, offset) => {
    if (offset + node.nodeSize > from && offset < to) {
      if (startIndex === -1) startIndex = index
      endIndex = index
    }
    index++
  })
  if (startIndex === -1) {
    startIndex = doc.childCount - 1
    endIndex = startIndex
  }
  return { startIndex, endIndex }
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
    name: 'insert_content',
    description:
      'Insert new markdown content after a top-level block. Use afterIndex -1 to insert at the very beginning of the document. On a blank document this replaces the empty paragraph.',
    inputSchema: {
      type: 'object',
      properties: {
        afterIndex: {
          type: 'integer',
          description: '0-based block index to insert after; -1 = document start',
        },
        markdown: { type: 'string', description: 'Markdown content to insert' },
      },
      required: ['afterIndex', 'markdown'],
    },
  },
  {
    name: 'replace_blocks',
    description:
      'Replace a range of top-level blocks (inclusive) with new markdown content. Use this for rewrites, formatting changes and deletions (empty markdown deletes the range).',
    inputSchema: {
      type: 'object',
      properties: {
        startIndex: { type: 'integer', description: '0-based index of the first block' },
        endIndex: { type: 'integer', description: '0-based index of the last block (inclusive)' },
        markdown: { type: 'string', description: 'Replacement markdown; empty string deletes' },
      },
      required: ['startIndex', 'endIndex', 'markdown'],
    },
  },
  {
    name: 'replace_text',
    description:
      'Replace every occurrence of an exact text within one block, keeping the block structure and the surrounding formatting. Prefer this over replace_blocks for small in-place fixes (a word, a number, a phrase). The match is plain text (no markdown syntax) and never crosses paragraph or table-cell boundaries.',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', description: '0-based index of the block to edit' },
        find: { type: 'string', description: 'Exact text to find (case-sensitive plain text)' },
        replace: { type: 'string', description: 'Replacement plain text; empty string deletes' },
      },
      required: ['blockIndex', 'find', 'replace'],
    },
  },
  {
    name: 'style_matches',
    description:
      'Apply or remove an inline style on every occurrence of an exact text within a block range — character-level formatting like bolding each "TODO". The match is plain text and never crosses paragraph or table-cell boundaries.',
    inputSchema: {
      type: 'object',
      properties: {
        find: { type: 'string', description: 'Exact text to find (case-sensitive plain text)' },
        style: {
          type: 'string',
          enum: [...STYLABLE_MARKS],
          description: 'Inline style to toggle on the matches',
        },
        remove: { type: 'boolean', description: 'true removes the style instead of applying it' },
        startIndex: {
          type: 'integer',
          description: '0-based index of the first block to search; defaults to 0',
        },
        endIndex: {
          type: 'integer',
          description: '0-based index of the last block (inclusive); defaults to the last block',
        },
      },
      required: ['find', 'style'],
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
      'Generate an illustration with AI from a text prompt and insert it into the document after a top-level block. For illustration/diagram-style art that image_search cannot find, or when the user asks to generate/draw a picture. Requires Genspark login with cloud tools enabled, and the document must have been saved to disk at least once.',
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
      'Read the document properties: the raw YAML frontmatter block at the top of the file (title, tags, date, …), without the --- fences.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'set_frontmatter',
    description:
      'Replace the whole YAML frontmatter block. Read it first and keep the keys you are not changing. Pass the inner YAML only (no --- fences); an empty string removes the block. The text is stored verbatim — emit valid YAML.',
    inputSchema: {
      type: 'object',
      properties: {
        yaml: { type: 'string', description: 'Full inner YAML text; empty string removes' },
      },
      required: ['yaml'],
    },
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

/** Byte offsets of each top-level block: [startPos, endPos] in the current doc */
function blockRange(doc: PmNode, from: number, to: number): { from: number; to: number } {
  let pos = 0
  let start = 0
  let end = 0
  for (let i = 0; i <= to; i++) {
    const child = doc.child(i)
    if (i === from) start = pos
    pos += child.nodeSize
    if (i === to) end = pos
  }
  return { from: start, to: end }
}

function parseMarkdownToNodes(editor: Editor, markdown: string): PmNode[] {
  // model output guard: `:::` fenced divs are not GFM and would land as
  // literal text — strip the fences and keep the body (same as file open).
  // Raw HTML needs no guard: parse runs it through the schema, so semantic
  // tags degrade to their GFM equivalents (<b>→bold, <img>→image) and
  // anything the schema cannot represent loses its styling, keeping text.
  const json = editor.markdown?.parse(stripLegacyFencedDivs(markdown))
  const content = json?.content ?? []
  return content.map((c) => editor.schema.nodeFromJSON(c))
}

/** exactly one empty paragraph — a lone image/math/table with no text is content, not blank */
function isBlankDoc(doc: PmNode): boolean {
  if (doc.childCount !== 1) return doc.childCount === 0
  const first = doc.firstChild!
  if (first.type.name !== 'paragraph') return false
  let hasLeaf = false
  first.descendants((node) => {
    if (!node.isText) hasLeaf = true
    return !hasLeaf
  })
  return !hasLeaf && first.textContent.trim() === ''
}

// ── plain-text matching inside blocks (replace_text / style_matches) ──

/** flattened text of [from, to) with a PM position per character; -1 marks
 *  atom placeholders and block seams so a match can never cross them */
function textIndexOf(doc: PmNode, from: number, to: number): { text: string; pos: number[] } {
  let text = ''
  const pos: number[] = []
  doc.nodesBetween(from, to, (node, p) => {
    if (node.isText && node.text) {
      const start = Math.max(from, p)
      const end = Math.min(to, p + node.text.length)
      for (let i = start; i < end; i++) {
        text += node.text[i - p]
        pos.push(i)
      }
    } else if (node.isLeaf) {
      text += ' '
      pos.push(-1)
    } else if (node.isBlock && text.length > 0 && !text.endsWith('\n')) {
      text += '\n'
      pos.push(-1)
    }
    return true
  })
  return { text, pos }
}

const MAX_MATCHES = 200

/** non-overlapping PM ranges of every plain-text occurrence of `find` */
function findMatches(
  index: { text: string; pos: number[] },
  find: string,
): Array<{ from: number; to: number }> {
  const ranges: Array<{ from: number; to: number }> = []
  let at = index.text.indexOf(find)
  while (at !== -1 && ranges.length < MAX_MATCHES) {
    const first = index.pos[at]!
    const last = index.pos[at + find.length - 1]!
    // contiguous inline text only (no atoms or block seams inside the match)
    let contiguous = first !== -1 && last - first === find.length - 1
    for (let i = at; contiguous && i < at + find.length; i++) {
      if (index.pos[i] === -1) contiguous = false
    }
    if (contiguous) ranges.push({ from: first, to: last + 1 })
    at = index.text.indexOf(find, at + find.length)
  }
  return ranges
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
  const doc = editor.state.doc
  const maxIndex = doc.childCount - 1
  // typeof guard: Number(null) is 0, which would silently mean "after block 0"
  const afterRaw = input.afterIndex
  const after =
    typeof afterRaw === 'number' && Number.isInteger(afterRaw)
      ? Math.min(Math.max(afterRaw, -1), maxIndex)
      : maxIndex
  const node = editor.schema.nodes.image!.create({
    src: rel,
    alt: String(input.alt ?? '').trim() || null,
  })
  let tr = editor.state.tr
  if (isBlankDoc(doc)) {
    tr = tr.replaceWith(0, doc.content.size, node)
    tr = markAiRange(tr, 0, tr.doc.content.size)
  } else {
    const pos = after === -1 ? 0 : blockRange(doc, after, after).to
    tr = tr.insert(pos, node)
    tr = markAiRange(tr, pos, pos + node.nodeSize)
  }
  editor.view.dispatch(tr)
  if (!userEditedDuringFetch) markDocSeen(editor)
  return {
    output: `Inserted the image (saved as ${rel}). ${INDEX_CHANGE_NOTICE}`,
    mutated: true,
    summary: labels.done,
  }
}

export function executeTool(
  editor: Editor,
  call: AgentToolCall,
  signal?: AbortSignal,
  fm?: FrontmatterAccess,
): ToolExecution | Promise<ToolExecution> {
  const doc = editor.state.doc
  const maxIndex = doc.childCount - 1

  switch (call.name) {
    case 'read_frontmatter': {
      if (!fm) return fail('frontmatter is not available', t('aiToolReadFm'))
      const inner = fm.read()
      return {
        output: inner || '(the document has no frontmatter)',
        mutated: false,
        summary: t('aiToolReadFm'),
      }
    }

    case 'set_frontmatter': {
      if (!fm) return fail('frontmatter is not available', t('aiToolSetFm'))
      const inner = String(call.input.yaml ?? '')
      fm.write(inner.trim() ? inner.replace(/\n+$/, '') : '')
      return {
        output: inner.trim() ? 'Frontmatter updated.' : 'Frontmatter removed.',
        mutated: true,
        summary: t('aiToolSetFm'),
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

    case 'insert_content': {
      if (editedExternally(editor)) return fail(STALE_DOC_ERROR, t('aiToolInsert'))
      const markdown = String(call.input.markdown ?? '')
      if (!markdown.trim()) return fail('markdown must not be empty', t('aiToolInsert'))
      const after = Number(call.input.afterIndex)
      if (!Number.isInteger(after) || after < -1 || after > maxIndex) {
        return fail(
          `afterIndex out of range; the document has ${doc.childCount} blocks.`,
          t('aiToolInsert'),
        )
      }
      let nodes: PmNode[]
      try {
        nodes = parseMarkdownToNodes(editor, markdown)
      } catch (err) {
        return fail(
          `markdown parse failed: ${err instanceof Error ? err.message : String(err)}`,
          t('aiToolInsert'),
        )
      }
      if (nodes.length === 0) return fail('markdown parsed to no content', t('aiToolInsert'))

      let tr = editor.state.tr
      if (isBlankDoc(doc)) {
        tr = tr.replaceWith(0, doc.content.size, nodes)
        tr = markAiRange(tr, 0, tr.doc.content.size)
      } else {
        const pos = after === -1 ? 0 : blockRange(doc, after, after).to
        const insertedSize = nodes.reduce((s, n) => s + n.nodeSize, 0)
        tr = tr.insert(pos, nodes)
        tr = markAiRange(tr, pos, pos + insertedSize)
      }
      editor.view.dispatch(tr)
      markDocSeen(editor)
      return {
        output: `Inserted ${nodes.length} block(s). ${INDEX_CHANGE_NOTICE}`,
        mutated: true,
        summary: t('aiToolInsertDone', { n: nodes.length }),
      }
    }

    case 'replace_blocks': {
      if (editedExternally(editor)) return fail(STALE_DOC_ERROR, t('aiToolReplace'))
      const start = clampIndex(call.input.startIndex, maxIndex)
      const end = clampIndex(call.input.endIndex, maxIndex)
      if (start === null || end === null || start > end) {
        return fail(
          `Invalid block range; the document has ${doc.childCount} blocks.`,
          t('aiToolReplace'),
        )
      }
      const markdown = String(call.input.markdown ?? '')
      let nodes: PmNode[]
      try {
        nodes = parseMarkdownToNodes(editor, markdown)
      } catch (err) {
        return fail(
          `markdown parse failed: ${err instanceof Error ? err.message : String(err)}`,
          t('aiToolReplace'),
        )
      }
      const { from, to } = blockRange(doc, start, end)
      let tr = editor.state.tr
      if (nodes.length === 0) {
        // deleting every block is not allowed by the schema — leave one empty paragraph
        if (start === 0 && end === maxIndex) {
          tr = tr.replaceWith(from, to, editor.schema.nodes.paragraph!.create())
        } else {
          tr = tr.delete(from, to)
        }
      } else {
        const insertedSize = nodes.reduce((s, n) => s + n.nodeSize, 0)
        tr = tr.replaceWith(from, to, nodes)
        tr = markAiRange(tr, from, from + insertedSize)
      }
      editor.view.dispatch(tr)
      markDocSeen(editor)
      return {
        output: `Replaced blocks ${start}-${end} with ${nodes.length} block(s). ${INDEX_CHANGE_NOTICE}`,
        mutated: true,
        summary: t('aiToolReplaceDone', { n: end - start + 1 }),
      }
    }

    case 'replace_text': {
      if (editedExternally(editor)) return fail(STALE_DOC_ERROR, t('aiToolReplaceText'))
      const blockIndex = clampIndex(call.input.blockIndex, maxIndex)
      if (blockIndex === null) {
        return fail(
          `blockIndex out of range; the document has ${doc.childCount} blocks.`,
          t('aiToolReplaceText'),
        )
      }
      const find = String(call.input.find ?? '')
      if (!find) return fail('find must not be empty', t('aiToolReplaceText'))
      const replace = String(call.input.replace ?? '')
      const { from, to } = blockRange(doc, blockIndex, blockIndex)
      const matches = findMatches(textIndexOf(doc, from, to), find)
      if (matches.length === 0) {
        return fail(
          `"${find}" was not found in block ${blockIndex}. The match is exact plain text (check spacing and punctuation); read the block with read_blocks to see its current text.`,
          t('aiToolReplaceText'),
        )
      }
      let tr = editor.state.tr
      // bottom-up so earlier matches keep their positions while editing
      for (let i = matches.length - 1; i >= 0; i--) {
        const m = matches[i]!
        if (replace) tr = tr.insertText(replace, m.from, m.to)
        else tr = tr.delete(m.from, m.to)
      }
      if (replace) {
        // final highlight positions: each match shifts by the deltas of the matches before it
        const delta = replace.length - find.length
        tr = markAiRanges(
          tr,
          matches.map((m, i) => ({
            from: m.from + delta * i,
            to: m.from + delta * i + replace.length,
          })),
        )
      }
      editor.view.dispatch(tr)
      markDocSeen(editor)
      return {
        output: `Replaced ${matches.length} occurrence(s) in block ${blockIndex}.`,
        mutated: true,
        summary: t('aiToolReplaceTextDone', { n: matches.length }),
      }
    }

    case 'style_matches': {
      if (editedExternally(editor)) return fail(STALE_DOC_ERROR, t('aiToolStyleText'))
      const find = String(call.input.find ?? '')
      if (!find) return fail('find must not be empty', t('aiToolStyleText'))
      const style = String(call.input.style ?? '') as StylableMark
      const markType = STYLABLE_MARKS.includes(style) ? editor.schema.marks[style] : undefined
      if (!markType) {
        return fail(`style must be one of: ${STYLABLE_MARKS.join(', ')}`, t('aiToolStyleText'))
      }
      const start = clampIndex(call.input.startIndex ?? 0, maxIndex)
      const end = clampIndex(call.input.endIndex ?? maxIndex, maxIndex)
      if (start === null || end === null || start > end) {
        return fail(
          `Invalid block range; the document has ${doc.childCount} blocks.`,
          t('aiToolStyleText'),
        )
      }
      const { from, to } = blockRange(doc, start, end)
      const matches = findMatches(textIndexOf(doc, from, to), find)
      if (matches.length === 0) {
        return fail(
          `"${find}" was not found in blocks ${start}-${end}. The match is exact plain text; read the range with read_blocks to see its current text.`,
          t('aiToolStyleText'),
        )
      }
      const remove = call.input.remove === true
      let tr = editor.state.tr
      for (const m of matches) {
        tr = remove
          ? tr.removeMark(m.from, m.to, markType)
          : tr.addMark(m.from, m.to, markType.create())
      }
      tr = markAiRanges(tr, matches)
      editor.view.dispatch(tr)
      markDocSeen(editor)
      return {
        output: `${remove ? 'Removed' : 'Applied'} ${style} on ${matches.length} match(es) in blocks ${start}-${end}.`,
        mutated: true,
        summary: t('aiToolStyleTextDone', { n: matches.length }),
      }
    }

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
