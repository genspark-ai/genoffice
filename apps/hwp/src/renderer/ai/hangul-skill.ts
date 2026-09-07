import type { AgentSkill, AgentToolDef } from '@genoffice/agent-core'
import { t } from '../i18n/locale'
import {
  clipPlainText,
  SELECTION_PREVIEW_CHARS,
  type HangulField,
  type HangulParagraphPreview,
  type HangulTable,
} from '../studio-text'

const SYSTEM_PROMPT = `You are GenOffice's Hangul (HWP) assistant. You read and edit the currently open Hangul document.

# Tools
- get_document_text: full document as plain text (may be truncated for long files)
- get_selection: currently selected text, if any
- get_paragraphs: numbered body paragraphs, including which ones are editable
- insert_content: insert new body paragraphs (plain text; newlines become new paragraphs)
- replace_paragraph: replace one body paragraph (caret, or index from get_paragraphs)
- replace_selection: replace only the selected span inside the current body paragraph
- get_fields / set_field: 누름틀 (click-here / form fields)
- get_tables / replace_cell: table cells (first paragraph in the cell)
- web_search: up-to-date facts beyond the document

# Editing
- Prefer replace_selection when the user has a selection and wants only that span changed.
- replace_paragraph replaces one whole body paragraph. Use index after get_paragraphs to edit a paragraph that does not have the caret. Call it once per paragraph; hashes change after each apply.
- Replacements are plain text: no C0 controls. insert_content may include newlines (one paragraph per line). Each body paragraph is capped at 4000 characters. Fields and cells allow up to 8000.
- insert_content writes new paragraphs. On a blank document, omit afterIndex so the empty caret paragraph is filled first. afterIndex -1 inserts at the start. You cannot lift the 4000-character body cap. Headers and footnotes are not editable.
- Tables, fields, mixed character formatting, headers, and footnotes cannot go through replace_paragraph. Use set_field or replace_cell instead when those tools apply.
- After a successful edit, summarize what changed. Do not claim an edit unless a mutating tool succeeded.

Read with tools before answering. You may use web_search for facts outside the document; cite sources in your reply, not as if you wrote them into the document.`

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
      'Replace one body paragraph with plain text. Omit index to use the paragraph that has the caret. Fails for tables, fields, mixed formatting, or text longer than 4000 characters.',
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
      'Insert new body paragraphs. Newlines become separate paragraphs. Omit afterIndex to insert after the caret (fills an empty caret paragraph first). afterIndex -1 inserts at the start. Use get_paragraphs indexes to insert after a specific paragraph.',
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

export interface HangulSkillDeps {
  fileName(): string
  pageCount(): number
  currentPage(): number | null
  hasSelection(): boolean
  selectionPreview(): string | null
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
            output: `Inserted ${result.count} paragraph(s) starting at [${result.start}].`,
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
