import type { Editor, JSONContent } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { AgentToolCall, AgentToolDef, ToolExecution } from '@genoffice/agent-core'
import { markAiRange } from '../editor/aiHighlight'
import { t } from '../i18n/locale'

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
  const isBlank =
    blockCount === 0 || (blockCount === 1 && doc.firstChild!.textContent.trim() === '')
  if (isBlank) {
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
  const selection = selectionMarkdown(editor)
  if (selection) lines.push('', '## User selection', selection)
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
  const json = editor.markdown?.parse(markdown)
  const content = json?.content ?? []
  return content.map((c) => editor.schema.nodeFromJSON(c))
}

function isBlankDoc(doc: PmNode): boolean {
  return doc.childCount === 1 && doc.firstChild!.textContent.trim() === ''
}

export function executeTool(editor: Editor, call: AgentToolCall): ToolExecution {
  const doc = editor.state.doc
  const maxIndex = doc.childCount - 1

  switch (call.name) {
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

    default:
      return fail(`Unknown tool: ${call.name}`, call.name)
  }
}
