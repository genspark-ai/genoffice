import type { AgentSkill, AgentToolDef } from '@genoffice/agent-core'
import { t } from '../i18n/locale'
import {
  clipPlainText,
  SELECTION_PREVIEW_CHARS,
  type HangulField,
  type HangulParagraphPreview,
  type HangulTable,
} from '../studio-text'

const SYSTEM_PROMPT = [
  'You are the Hangul (HWP) assistant built into GenOffice. You read and edit the currently open document through tools only.',
  '',
  '# Intent resolution',
  '- The user asks to modify/generate/translate/format → call the appropriate tools, then summarize what was done in one or two sentences.',
  '- The user is asking a question or consulting (what is this about, word count, writing advice) → answer in chat without mutating tools.',
  '- When intent is unclear, read first (the paragraph list in the message, or get_document_text / get_paragraphs), then decide.',
  '',
  '# Tool usage',
  '- Every user message may include a body-paragraph list (index|status|preview). Previews can be truncated — call get_paragraphs or get_document_text before rewriting a long paragraph.',
  '- New content: one insert_content call with the full draft. Newlines become paragraphs. Never insert a skeleton of headings and then fill it with replace_paragraph.',
  '- If insert_content reports "Inserted N paragraph(s)", it succeeded and every newline is already a paragraph. Do not call insert_content again for the same draft, even to "continue" after the first heading.',
  '- replace_paragraph changes one existing paragraph. Call it once per turn; hashes change after each apply, so parallel or batched replaces fail.',
  '- Prefer replace_selection when the user has a selection and wants only that span changed.',
  '- Small in-place fixes stay on replace_selection / one replace_paragraph. Full drafts and multi-paragraph additions go through insert_content.',
  '- Fields (누름틀): get_fields / set_field. Tables: get_tables / replace_cell (first paragraph in the cell).',
  '- After any mutation, indexes change — get_paragraphs before further index-based edits. If a tool reports success, do not retry blindly.',
  '- Replacements are plain text: no C0 controls. Each body paragraph is capped at 4000 characters (fields/cells 8000). One insert_content may add at most 80 paragraphs. Headers and footnotes are not editable.',
  '',
  '# Writing a new document',
  '- A new Hangul file often has one locked section-control paragraph. That is not "cannot write" — omit afterIndex and insert_content will skip locked rows and write after them.',
  '- When the body looks blank and the user asks for a plan, report, or any draft, write the complete document in one insert_content call: title first, then numbered sections, short paragraphs.',
  '- If the topic is unspecified, either ask one clarifying question first or write a complete generic template in that same call (headings AND body lines together). Never insert empty numbered headings to fill later.',
  '- Never invent facts, dates, names, or budget numbers. Use web_search for current facts; cite sources in your reply, not as if you wrote them into the document.',
  '',
  '# Template filling',
  '- When the user asks to fill a form/template, scan get_fields and placeholder text in the paragraph list.',
  "- Fill the values the user's message answers via set_field or replace_paragraph. Ask once for the rest — never invent facts to fill a field.",
  '',
  '# Conversation',
  '- Keep replies short; the document edit is the deliverable.',
  '- After a successful edit, summarize what changed. Do not claim an edit unless a mutating tool succeeded.',
  '',
  '# Known failures (must avoid)',
  '- HG-1 Inserting a heading skeleton, then calling replace_paragraph once per line (slow; hashes break).',
  '- HG-2 Retrying insert_content after "Inserted N paragraph(s)" (duplicates the draft).',
  '- HG-3 Treating the first locked section-control paragraph as "the document cannot be edited".',
  '- HG-4 Firing several replace_paragraph calls in the same turn.',
].join('\n')

const TOOLS: AgentToolDef[] = [
  {
    name: 'get_document_text',
    description: 'Return the Hangul document as plain text. Long documents may be truncated.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'get_selection',
    description: 'Return the current selection as plain text, or report that nothing is selected.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'get_paragraphs',
    description:
      'List body paragraphs with 0-based indexes. Use the index with replace_paragraph. Non-editable rows (tables, fields, mixed formatting) say why.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'replace_paragraph',
    description:
      'Replace one existing body paragraph. Not for drafting: write new documents with insert_content. Omit index to use the caret paragraph. Call once per turn. Fails for tables, fields, mixed formatting, or text longer than 4000 characters.',
    inputSchema: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'Replacement plain text for the whole paragraph' },
        index: {
          type: 'number',
          description: '0-based paragraph index from get_paragraphs. Omit to use the caret paragraph.',
        },
      },
      required: ['text'],
    },
  },
  {
    name: 'insert_content',
    description:
      'Write new body paragraphs in one call. Pass the full draft; newlines become paragraphs. Omit afterIndex to insert after the caret (fills an empty caret paragraph first). afterIndex -1 inserts at the start. Use get_paragraphs indexes to insert after a specific paragraph. If it reports Inserted N, do not call it again.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'Plain text to insert. Newlines start new paragraphs.',
        },
        afterIndex: {
          type: 'number',
          description:
            'Insert after this 0-based paragraph index from get_paragraphs. -1 = start of document. Omit to use the caret.',
        },
      },
      required: ['text'],
    },
  },
  {
    name: 'replace_selection',
    description:
      'Replace only the selected span inside the current body paragraph. Fails when nothing is selected in that paragraph.',
    inputSchema: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'Replacement plain text for the selected span' },
      },
      required: ['text'],
    },
  },
  {
    name: 'get_fields',
    description: 'List 누름틀 / form fields with their current values.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'set_field',
    description: 'Set a 누름틀 / form field value by name from get_fields.',
    inputSchema: {
      type: 'object',
      properties: {
        name: { type: 'string', description: 'Field name' },
        value: { type: 'string', description: 'New field value' },
      },
      required: ['name', 'value'],
    },
  },
  {
    name: 'get_tables',
    description: 'List tables and cell text. Rows and columns are 0-based.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'replace_cell',
    description:
      'Replace the first paragraph of a table cell. table / row / col are 0-based indexes from get_tables.',
    inputSchema: {
      type: 'object',
      properties: {
        table: { type: 'number', description: '0-based table index from get_tables' },
        row: { type: 'number', description: '0-based row' },
        col: { type: 'number', description: '0-based column' },
        text: { type: 'string', description: 'Replacement plain text for the cell' },
      },
      required: ['table', 'row', 'col', 'text'],
    },
  },
]

const MUTATING_TOOLS = [
  'insert_content',
  'replace_paragraph',
  'replace_selection',
  'set_field',
  'replace_cell',
]

function requireToolIndex(value: unknown, label: string): number {
  if (value == null || value === '' || typeof value === 'boolean') {
    throw new Error(`${label} must be an integer`)
  }
  if (typeof value === 'string' && value.trim() === '') {
    throw new Error(`${label} must be an integer`)
  }
  const n = typeof value === 'number' ? value : Number(value)
  if (!Number.isInteger(n)) throw new Error(`${label} must be an integer`)
  return n
}

function optionalToolIndex(value: unknown, label = 'index'): number | undefined {
  if (value == null || value === '') return undefined
  return requireToolIndex(value, label)
}

const CONTEXT_PREVIEW_CHARS = 60
const CONTEXT_PREVIEW_TIGHT = 20
const CONTEXT_MAX_CHARS = 8000

function clipPreview(text: string, max: number): string {
  const one = text.replace(/\s+/g, ' ').trim() || '(empty)'
  return one.length > max ? `${one.slice(0, max)}…` : one
}

function formatParagraphContext(items: HangulParagraphPreview[]): string {
  if (items.length === 0) return 'The document has no body paragraphs yet.'
  const blank = items.every((item) => !item.editable || !item.text.trim())
  const render = (max: number) =>
    items.map((item) => {
      const status = item.editable ? 'editable' : `locked:${item.reason || 'unknown'}`
      return `${item.index}|${status}|${clipPreview(item.text, max)}`
    })
  let lines = render(CONTEXT_PREVIEW_CHARS)
  if (lines.join('\n').length > CONTEXT_MAX_CHARS) lines = render(CONTEXT_PREVIEW_TIGHT)
  if (lines.join('\n').length > CONTEXT_MAX_CHARS && items.length > 35) {
    lines = [
      ...lines.slice(0, 25),
      `…(${items.length - 35} paragraphs omitted; numbering is continuous)…`,
      ...lines.slice(-10),
    ]
  }
  const header = blank
    ? `The body looks blank (${items.length} paragraph(s)). Locked control/field rows cannot be replaced — write with one insert_content call (omit afterIndex).`
    : `Body paragraphs (${items.length}; index|status|preview):`
  return [header, ...lines].join('\n')
}

export interface HangulSkillDeps {
  fileName(): string
  pageCount(): number
  currentPage(): number | null
  hasSelection(): boolean
  selectionPreview(): string | null
  paragraphPreview(): HangulParagraphPreview[] | null
  getDocumentText(): Promise<string>
  getSelection(): Promise<string | null>
  listParagraphs(): Promise<HangulParagraphPreview[]>
  insertContent(text: string, afterIndex?: number): Promise<{ count: number; start: number }>
  replaceParagraph(text: string, index?: number): Promise<{ before: string; after: string }>
  replaceSelection(text: string): Promise<{ before: string; after: string }>
  listFields(): Promise<HangulField[]>
  setField(name: string, value: string): Promise<{ name: string; before: string; after: string }>
  listTables(): Promise<HangulTable[]>
  replaceCell(
    table: number,
    row: number,
    col: number,
    text: string,
  ): Promise<{ before: string; after: string }>
}

function formatParagraphs(items: HangulParagraphPreview[]): string {
  if (items.length === 0) return '(no body paragraphs)'
  return items
    .map((item) => {
      const status = item.editable ? 'editable' : `not editable: ${item.reason || 'unknown'}`
      const preview = item.text.replace(/\s+/g, ' ').trim() || '(empty)'
      return `[${item.index}] ${status}\n${preview}`
    })
    .join('\n\n')
}

function formatFields(fields: HangulField[]): string {
  if (fields.length === 0) return '(no fields)'
  return fields
    .map((field) => `${field.name}=${field.value || '(empty)'}${field.type ? ` (${field.type})` : ''}`)
    .join('\n')
}

function formatTables(tables: HangulTable[]): string {
  if (tables.length === 0) return '(no tables)'
  return tables
    .map((table) => {
      const cells = table.cells
        .map((cell) => `  r${cell.row}c${cell.col}=${cell.text.replace(/\s+/g, ' ').trim() || '(empty)'}`)
        .join('\n')
      return `table[${table.index}] ${table.rows}x${table.cols}\n${cells || '  (no cells)'}`
    })
    .join('\n\n')
}

export function createHangulSkill(getDeps: () => HangulSkillDeps): AgentSkill {
  return {
    id: 'hangul',
    systemPrompt: SYSTEM_PROMPT,
    tools: TOOLS,
    buildContext: () => {
      const deps = getDeps()
      const parts = [`Hangul document: "${deps.fileName()}", ${deps.pageCount()} page(s).`]
      const page = deps.currentPage()
      if (page) parts.push(`Current page: ${page}.`)
      const paragraphs = deps.paragraphPreview()
      if (paragraphs) parts.push(formatParagraphContext(paragraphs))
      if (deps.hasSelection()) {
        const preview = deps.selectionPreview()?.trim()
        if (preview) {
          const clipped =
            preview.length > SELECTION_PREVIEW_CHARS
              ? `${preview.slice(0, SELECTION_PREVIEW_CHARS)}…`
              : preview
          parts.push(`The user has selected the following text:\n"""\n${clipped}\n"""`)
          parts.push(
            'Use replace_selection to change only this span, or replace_paragraph to rewrite the whole paragraph that contains it.',
          )
        } else {
          parts.push('The user has a selection, but the raw selected text is not available.')
        }
      } else {
        parts.push(
          'No text is selected. insert_content without afterIndex writes after the caret (or fills an empty caret paragraph). replace_paragraph without index changes the paragraph at the caret.',
        )
      }
      return parts.join('\n')
    },
    executeTool: async (call) => {
      const deps = getDeps()
      if (call.name === 'get_document_text') {
        try {
          const text = await deps.getDocumentText()
          return {
            output: text || '(empty document)',
            summary: 'Read document text',
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: 'Read document text',
          }
        }
      }
      if (call.name === 'get_selection') {
        try {
          const text = await deps.getSelection()
          if (!text?.trim()) {
            return { output: '(no selection)', summary: 'Read selection' }
          }
          return { output: clipPlainText(text), summary: 'Read selection' }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: 'Read selection',
          }
        }
      }
      if (call.name === 'get_paragraphs') {
        try {
          const items = await deps.listParagraphs()
          return { output: formatParagraphs(items), summary: 'Listed paragraphs' }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: 'Listed paragraphs',
          }
        }
      }
      if (call.name === 'insert_content') {
        const text = String(call.input.text ?? '')
        let afterIndex: number | undefined
        try {
          afterIndex = optionalToolIndex(call.input.afterIndex, 'afterIndex')
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolInsertContent'),
          }
        }
        try {
          const result = await deps.insertContent(text, afterIndex)
          return {
            output: `Inserted ${result.count} paragraph(s) starting at [${result.start}]. The text you passed is already in the document. Do not call insert_content again unless the user asked for more content.`,
            mutated: true,
            summary: t('aiToolInsertContentDone'),
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolInsertContent'),
          }
        }
      }
      if (call.name === 'replace_paragraph') {
        const text = String(call.input.text ?? '')
        let index: number | undefined
        try {
          index = optionalToolIndex(call.input.index)
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolReplaceParagraph'),
          }
        }
        try {
          const result = await deps.replaceParagraph(text, index)
          return {
            output: `Replaced paragraph${index == null ? '' : ` [${index}]`}.\nBefore:\n${result.before || '(empty)'}\nAfter:\n${result.after || '(empty)'}`,
            mutated: true,
            summary: t('aiToolReplaceParagraphDone'),
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolReplaceParagraph'),
          }
        }
      }
      if (call.name === 'replace_selection') {
        const text = String(call.input.text ?? '')
        try {
          const result = await deps.replaceSelection(text)
          return {
            output: `Replaced selection.\nBefore:\n${result.before || '(empty)'}\nAfter:\n${result.after || '(empty)'}`,
            mutated: true,
            summary: t('aiToolReplaceSelectionDone'),
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolReplaceSelection'),
          }
        }
      }
      if (call.name === 'get_fields') {
        try {
          const fields = await deps.listFields()
          return { output: formatFields(fields), summary: 'Listed fields' }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: 'Listed fields',
          }
        }
      }
      if (call.name === 'set_field') {
        const name = String(call.input.name ?? '')
        const value = String(call.input.value ?? '')
        try {
          const result = await deps.setField(name, value)
          return {
            output: `Set field ${result.name}.\nBefore:\n${result.before || '(empty)'}\nAfter:\n${result.after || '(empty)'}`,
            mutated: true,
            summary: t('aiToolSetFieldDone'),
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolSetField'),
          }
        }
      }
      if (call.name === 'get_tables') {
        try {
          const tables = await deps.listTables()
          return { output: formatTables(tables), summary: 'Listed tables' }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: 'Listed tables',
          }
        }
      }
      if (call.name === 'replace_cell') {
        const text = String(call.input.text ?? '')
        let table: number
        let row: number
        let col: number
        try {
          table = requireToolIndex(call.input.table, 'table')
          row = requireToolIndex(call.input.row, 'row')
          col = requireToolIndex(call.input.col, 'col')
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolReplaceCell'),
          }
        }
        try {
          const result = await deps.replaceCell(table, row, col, text)
          return {
            output: `Replaced cell table[${table}] r${row}c${col}.\nBefore:\n${result.before || '(empty)'}\nAfter:\n${result.after || '(empty)'}`,
            mutated: true,
            summary: t('aiToolReplaceCellDone'),
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolReplaceCell'),
          }
        }
      }
      return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
    },
    verifyResponse: (finalText, executed) => {
      const claimed = /바꿨|수정했|고쳤|넣었|삽입했|filled|replaced|rewrote|inserted|edited the (document|paragraph|selection|cell|field)/i.test(
        finalText,
      )
      if (!claimed) return null
      if (executed.some((call) => MUTATING_TOOLS.includes(call.name) && call.ok)) return null
      return 'You claimed to edit the document, but no edit tool succeeded. Tell the user the document was not changed and do not claim an edit.'
    },
  }
}
