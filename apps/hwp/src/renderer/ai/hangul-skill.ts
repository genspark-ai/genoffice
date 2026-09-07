import type { AgentSkill, AgentToolDef } from '@genoffice/agent-core'
import { t } from '../i18n/locale'
import {
  clipPlainText,
  SELECTION_PREVIEW_CHARS,
  type HangulField,
  type HangulFormatSpec,
  type HangulPageSetupSpec,
  type HangulPaper,
  type HangulParagraphPreview,
  type HangulTable,
  type HangulTableEditAction,
  type HangulTableEditSpec,
  type HangulTableStyleSpec,
  type HangulVAlign,
} from '../studio-text'

const SYSTEM_PROMPT = [
  'You are the Hangul (HWP) assistant built into GenOffice. You read and edit the currently open document through tools only.',
  '',
  '# Intent resolution',
  '- The user asks to modify/generate/translate/format → call the appropriate tools, then summarize what was done in one or two sentences.',
  '- The user is asking a question or consulting (what is this about, word count, writing advice) → answer in chat without mutating tools.',
  '- When intent is unclear, read the body paragraph list in the message (or get_paragraphs / get_document_text), then decide.',
  '',
  '# Tool usage',
  '- Every user message carries the latest "body paragraph list" (index|status|preview). Previews may be truncated — get_paragraphs or get_document_text before rewriting a long paragraph.',
  '- After any mutation, indexes change — get_paragraphs before further index-based edits. If a tool reports success, do not retry it.',
  '- New drafts: one insert_content call with the full text. Newlines become paragraphs. Cap: 80 paragraphs per call, 4000 characters per paragraph (fields/cells 8000). If the 80-paragraph cap rejects the call, insert only the remainder with afterIndex set to the last written paragraph.',
  '- If insert_content reports "Inserted N paragraph(s)", it succeeded. Do not call it again for the same draft.',
  '- replace_paragraph changes one existing paragraph, once per turn (hashes change after each apply). Prefer replace_selection when the user has a selection and wants only that span changed.',
  '- Small in-place fixes stay on replace_selection / one replace_paragraph. Multi-paragraph additions go through insert_content. Omit afterIndex / index to use the caret.',
  '- Fields (누름틀): get_fields / set_field. Existing tables: get_tables / replace_cell (one leftover cell). New tables: insert_table with cells[][] already filled — never insert an empty table and then call replace_cell once per cell. Never fake a table with tabs, markdown pipes, or ASCII.',
  '- Table structure (add/delete a row or column, merge or split cells): edit_table on an existing table from get_tables. Do not rebuild the table with insert_table just to change rows or merge a header.',
  '- Cell chrome (fill, vertical align, border) and table width in mm: style_table. Text bold/size still uses apply_format with table + row.',
  '- Page setup (portrait/landscape, A4/A3 paper, margins in mm, 1–4 columns): set_page. Do not tell the user to open 쪽 설정.',
  '- Formatting (bold/italic/underline/strikethrough, fontSize in points, color hex, font name, align, lineSpacing, indentLeft/Right/FirstLine in points, bullet/number lists): apply_format on existing paragraphs, or on table cells with table + row from get_tables (omit col to format the whole row — e.g. header / first row bold 13pt). Do not mix table/row with index/indexes. Skip locked body rows. After insert_content, format the title at the starting index from that tool result — not index 0 when that row is locked. Do not rewrite a paragraph just to change style. Do not tell the user to format a table row by hand. Do not tell the user formatting was applied unless apply_format succeeded.',
  '- Replacements are plain text: no C0 controls. Headers and footnotes are not editable.',
  '',
  '# Writing a new document',
  '- A new file often has one locked section-control paragraph. That is not "cannot write" — omit afterIndex; insert_content skips locked rows.',
  '- When the body is blank and the user wants a plan, report, or any draft, write the complete document in one insert_content call: title first, then numbered sections, short paragraphs. Then apply_format / insert_table as needed. Do not put HTML or markdown tables in insert_content.',
  '- If the topic is unspecified, ask one clarifying question or write a complete generic template (headings AND body together). Never insert empty numbered headings to fill later.',
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
  '- HG-1 Inserting a heading skeleton, then calling replace_paragraph once per line.',
  '- HG-2 Retrying insert_content after "Inserted N paragraph(s)".',
  '- HG-3 Treating the first locked section-control paragraph as "the document cannot be edited".',
  '- HG-4 Firing several replace_paragraph calls in the same turn.',
  '- HG-5 Drawing a table with tabs, markdown pipes, or ASCII lines instead of insert_table.',
  '- HG-6 Rewriting a paragraph with replace_paragraph only to change bold/align/list/size.',
  '- HG-7 Inserting an empty table, then firing many replace_cell calls. Pass cells[][] on insert_table.',
  '- HG-8 Telling the user table-cell format is impossible, or formatting a body paragraph instead of apply_format with table + row.',
  '- HG-9 Rebuilding a table with insert_table to add/delete a row or merge cells. Use edit_table.',
  '- HG-10 Telling the user to set paper, margins, columns, or cell fill by hand. Use set_page / style_table.',
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
      'Write new body paragraphs in one call. Pass the full draft; newlines become paragraphs. Omit afterIndex to insert after the caret (fills an empty caret paragraph first). afterIndex -1 inserts at the start. Use get_paragraphs indexes to insert after a specific paragraph. If it reports Inserted N, do not resend the same draft. A rejected 80-paragraph cap is not success — send only the leftover lines with afterIndex.',
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
  {
    name: 'insert_table',
    description:
      'Insert a new table. rows/cols are required. Pass cells as a row-major string[][] whenever you know the contents — do not insert a blank table and fill it with replace_cell. Omit afterIndex to insert after the caret. afterIndex -1 inserts at the start.',
    inputSchema: {
      type: 'object',
      properties: {
        rows: { type: 'number', description: 'Number of rows (1-20)' },
        cols: { type: 'number', description: 'Number of columns (1-10)' },
        cells: {
          type: 'array',
          description: 'Optional row-major cell text. Missing or empty cells stay blank.',
          items: { type: 'array', items: { type: 'string' } },
        },
        afterIndex: {
          type: 'number',
          description:
            'Insert after this 0-based paragraph index from get_paragraphs. -1 = start of document. Omit to use the caret.',
        },
      },
      required: ['rows', 'cols'],
    },
  },
  {
    name: 'apply_format',
    description:
      'Set character or paragraph formatting on existing body text, or on table cells. Body: omit index to format the caret paragraph (or the current selection span for character styles); use indexes for several paragraphs. Table cells: pass table + row from get_tables (omit col to format the whole row). Do not mix table/row with index/indexes. Do not rewrite text just to change style.',
    inputSchema: {
      type: 'object',
      properties: {
        index: {
          type: 'number',
          description: '0-based paragraph index from get_paragraphs. Omit to use the caret.',
        },
        indexes: {
          type: 'array',
          items: { type: 'number' },
          description: 'Format several paragraphs. At most 40. Overrides index.',
        },
        table: {
          type: 'number',
          description: '0-based table index from get_tables. With row, formats cells instead of a body paragraph.',
        },
        row: {
          type: 'number',
          description: '0-based table row. Required with table. Omit col to format every cell in the row.',
        },
        col: {
          type: 'number',
          description: '0-based table column. Omit to format the whole row.',
        },
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        underline: { type: 'boolean' },
        strikethrough: { type: 'boolean' },
        fontSize: { type: 'number', description: 'Font size in points (8-72)' },
        color: {
          type: 'string',
          description: 'Text color as 3- or 6-digit hex, with or without # (e.g. FF0000)',
        },
        font: { type: 'string', description: 'Font family name (e.g. 맑은 고딕, Pretendard)' },
        lineSpacing: {
          type: 'number',
          description: 'Line spacing as a multiplier (1.5) or percent (150)',
        },
        indentLeft: { type: 'number', description: 'Left indent in points' },
        indentRight: { type: 'number', description: 'Right indent in points' },
        indentFirstLine: {
          type: 'number',
          description: 'First-line indent in points; negative = hanging indent',
        },
        align: {
          type: 'string',
          enum: ['left', 'center', 'right', 'justify'],
          description: 'Paragraph alignment',
        },
        list: {
          type: 'string',
          enum: ['none', 'bullet', 'number'],
          description: 'Bullet or numbered list, or none to clear',
        },
      },
    },
  },
  {
    name: 'edit_table',
    description:
      'Change an existing table structure. insert_row / insert_column / delete_row / delete_column / merge / split. table is from get_tables. after defaults true (below / right). merge needs row, col, endRow, endCol. split needs row, col and optional splitRows / splitCols (default 2x1). Do not insert a new table to change structure.',
    inputSchema: {
      type: 'object',
      properties: {
        action: {
          type: 'string',
          enum: ['insert_row', 'insert_column', 'delete_row', 'delete_column', 'merge', 'split'],
        },
        table: { type: 'number', description: '0-based table index from get_tables' },
        row: { type: 'number', description: '0-based row' },
        col: { type: 'number', description: '0-based column' },
        after: {
          type: 'boolean',
          description: 'For insert_row / insert_column: true = below/right (default), false = above/left',
        },
        endRow: { type: 'number', description: 'Merge end row (inclusive)' },
        endCol: { type: 'number', description: 'Merge end column (inclusive)' },
        splitRows: { type: 'number', description: 'Split into this many rows (1-10). Default 2.' },
        splitCols: { type: 'number', description: 'Split into this many columns (1-10). Default 1.' },
      },
      required: ['action', 'table'],
    },
  },
  {
    name: 'style_table',
    description:
      'Set table cell chrome or table width. fill is hex, valign is top/center/bottom, border is hex or false to clear, width is millimeters for the whole table. Omit row to style every cell; omit col to style the whole row. Not for bold/font size — that is apply_format.',
    inputSchema: {
      type: 'object',
      properties: {
        table: { type: 'number', description: '0-based table index from get_tables' },
        row: { type: 'number', description: '0-based row. Omit to style every cell (or only set width).' },
        col: { type: 'number', description: '0-based column. Omit to style the whole row.' },
        fill: { type: 'string', description: 'Cell fill as 3- or 6-digit hex' },
        valign: { type: 'string', enum: ['top', 'center', 'bottom'] },
        border: {
          description: 'Border color hex, or false to clear borders',
        },
        width: { type: 'number', description: 'Table width in millimeters (20-300)' },
      },
      required: ['table'],
    },
  },
  {
    name: 'set_page',
    description:
      'Set page paper, orientation, margins in millimeters, or column count (1-4). Omit unused fields. Uses the first section.',
    inputSchema: {
      type: 'object',
      properties: {
        orientation: { type: 'string', enum: ['portrait', 'landscape'] },
        paper: { type: 'string', enum: ['A4', 'A3', 'B4', 'B5', 'Letter', 'Legal'] },
        marginTop: { type: 'number', description: 'Top margin in mm (0-80)' },
        marginBottom: { type: 'number', description: 'Bottom margin in mm (0-80)' },
        marginLeft: { type: 'number', description: 'Left margin in mm (0-80)' },
        marginRight: { type: 'number', description: 'Right margin in mm (0-80)' },
        columns: { type: 'number', description: 'Number of columns (1-4)' },
        columnSpacing: { type: 'number', description: 'Column spacing in mm (0-80)' },
      },
    },
  },
]

const MUTATING_TOOLS = [
  'insert_content',
  'replace_paragraph',
  'replace_selection',
  'set_field',
  'replace_cell',
  'insert_table',
  'apply_format',
  'edit_table',
  'style_table',
  'set_page',
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
  if (items.length === 0) return 'The body is currently blank.'
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
    ? `The body is currently blank (${items.length} paragraph(s); index|status|preview):`
    : `Body paragraph list (${items.length}; index|status|preview):`
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
  insertTable(
    rows: number,
    cols: number,
    cells?: string[][],
    afterIndex?: number,
  ): Promise<{ table: number; rows: number; cols: number; unfilled: string[] }>
  applyFormat(
    format: HangulFormatSpec,
    index?: number,
    indexes?: number[],
    cell?: { table: number; row: number; col?: number },
  ): Promise<{ indexes: number[]; applied: string[]; table?: number; row?: number }>
  editTable(spec: HangulTableEditSpec): Promise<{ table: number; action: HangulTableEditAction; detail: string }>
  styleTable(spec: HangulTableStyleSpec): Promise<{ table: number; applied: string[] }>
  setPage(spec: HangulPageSetupSpec): Promise<{ applied: string[] }>
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
          parts.push(`Content selected by the user:\n"""\n${clipped}\n"""`)
        } else {
          parts.push('The user has a selection, but the selected text is not available.')
        }
      } else {
        parts.push('No text is selected.')
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
            output: `Inserted ${result.count} paragraph(s) starting at [${result.start}]. The text you passed is already in the document. Do not call insert_content again unless the user asked for more content. Title/heading styles: apply_format on [${result.start}] (skip locked rows such as index 0).`,
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
      if (call.name === 'insert_table') {
        let rows: number
        let cols: number
        let afterIndex: number | undefined
        try {
          rows = requireToolIndex(call.input.rows, 'rows')
          cols = requireToolIndex(call.input.cols, 'cols')
          afterIndex = optionalToolIndex(call.input.afterIndex, 'afterIndex')
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolInsertTable'),
          }
        }
        try {
          const cells = Array.isArray(call.input.cells) ? (call.input.cells as string[][]) : undefined
          const result = await deps.insertTable(rows, cols, cells, afterIndex)
          const leftover = !cells
            ? ' Table is empty because cells[][] was omitted. Insert again only if the user asked for another table; otherwise leave it and tell the user the grid is blank. Do not fire one replace_cell per cell.'
            : result.unfilled.length === 0
              ? ' Cells were filled in this call. Do not call replace_cell unless a leftover cell is listed.'
              : result.unfilled.length > 3
                ? ` ${result.unfilled.length} cells stayed empty. Do not fire one replace_cell per cell. Tell the user the table is there but those cells are blank.`
                : ` Unfilled cells: ${result.unfilled.join(', ')}. Fill only those leftover cells with replace_cell.`
          return {
            output: `Inserted table[${result.table}] ${result.rows}x${result.cols}.${leftover} Do not insert another table unless the user asked for more than one.`,
            mutated: true,
            summary: t('aiToolInsertTableDone'),
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolInsertTable'),
          }
        }
      }
      if (call.name === 'apply_format') {
        let index: number | undefined
        let indexes: number[] | undefined
        let cell: { table: number; row: number; col?: number } | undefined
        try {
          index = optionalToolIndex(call.input.index)
          if (call.input.indexes != null) {
            if (!Array.isArray(call.input.indexes)) throw new Error('indexes must be an array')
            indexes = call.input.indexes.map((value, i) => requireToolIndex(value, `indexes[${i}]`))
          }
          const table = optionalToolIndex(call.input.table, 'table')
          const row = optionalToolIndex(call.input.row, 'row')
          const col = optionalToolIndex(call.input.col, 'col')
          if (table != null || row != null || col != null) {
            if (table == null) throw new Error('table must be an integer')
            if (row == null) throw new Error('row must be an integer')
            if (index != null || indexes != null) {
              throw new Error('do not mix table/row with index/indexes')
            }
            cell = { table, row, col }
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolApplyFormat'),
          }
        }
        const format: HangulFormatSpec = {}
        if (typeof call.input.bold === 'boolean') format.bold = call.input.bold
        if (typeof call.input.italic === 'boolean') format.italic = call.input.italic
        if (typeof call.input.underline === 'boolean') format.underline = call.input.underline
        if (typeof call.input.strikethrough === 'boolean') format.strikethrough = call.input.strikethrough
        if (call.input.fontSize != null) format.fontSize = Number(call.input.fontSize)
        if (typeof call.input.color === 'string') format.color = call.input.color
        if (typeof call.input.font === 'string') format.font = call.input.font
        if (call.input.lineSpacing != null) format.lineSpacing = Number(call.input.lineSpacing)
        if (call.input.indentLeft != null) format.indentLeft = Number(call.input.indentLeft)
        if (call.input.indentRight != null) format.indentRight = Number(call.input.indentRight)
        if (call.input.indentFirstLine != null) format.indentFirstLine = Number(call.input.indentFirstLine)
        if (typeof call.input.align === 'string') {
          format.align = call.input.align as HangulFormatSpec['align']
        }
        if (typeof call.input.list === 'string') {
          format.list = call.input.list as HangulFormatSpec['list']
        }
        try {
          const result = await deps.applyFormat(format, index, indexes, cell)
          const target =
            result.table != null && result.row != null
              ? `table[${result.table}] row ${result.row} cell(s) [${result.indexes.join(', ')}]`
              : `paragraph(s) [${result.indexes.join(', ')}]`
          return {
            output: `Applied ${result.applied.join(', ') || 'format'} to ${target}.`,
            mutated: true,
            summary: t('aiToolApplyFormatDone'),
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolApplyFormat'),
          }
        }
      }
      if (call.name === 'edit_table') {
        const actions = new Set<HangulTableEditAction>([
          'insert_row',
          'insert_column',
          'delete_row',
          'delete_column',
          'merge',
          'split',
        ])
        try {
          const action = String(call.input.action ?? '') as HangulTableEditAction
          if (!actions.has(action)) throw new Error('action must be insert_row, insert_column, delete_row, delete_column, merge, or split')
          const spec: HangulTableEditSpec = {
            action,
            table: requireToolIndex(call.input.table, 'table'),
            row: optionalToolIndex(call.input.row, 'row'),
            col: optionalToolIndex(call.input.col, 'col'),
            after: typeof call.input.after === 'boolean' ? call.input.after : undefined,
            endRow: optionalToolIndex(call.input.endRow, 'endRow'),
            endCol: optionalToolIndex(call.input.endCol, 'endCol'),
            splitRows: optionalToolIndex(call.input.splitRows, 'splitRows'),
            splitCols: optionalToolIndex(call.input.splitCols, 'splitCols'),
          }
          const result = await deps.editTable(spec)
          return {
            output: `Edited table[${result.table}]: ${result.action} (${result.detail}). Indexes may have changed — get_tables before further table edits.`,
            mutated: true,
            summary: t('aiToolEditTableDone'),
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolEditTable'),
          }
        }
      }
      if (call.name === 'style_table') {
        try {
          const spec: HangulTableStyleSpec = {
            table: requireToolIndex(call.input.table, 'table'),
            row: optionalToolIndex(call.input.row, 'row'),
            col: optionalToolIndex(call.input.col, 'col'),
          }
          if (typeof call.input.fill === 'string') spec.fill = call.input.fill
          if (typeof call.input.valign === 'string') spec.valign = call.input.valign as HangulVAlign
          if (call.input.border === false) spec.border = false
          else if (typeof call.input.border === 'string') spec.border = call.input.border
          if (call.input.width != null) spec.width = Number(call.input.width)
          const result = await deps.styleTable(spec)
          return {
            output: `Styled table[${result.table}]: ${result.applied.join(', ')}.`,
            mutated: true,
            summary: t('aiToolStyleTableDone'),
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolStyleTable'),
          }
        }
      }
      if (call.name === 'set_page') {
        try {
          const spec: HangulPageSetupSpec = {}
          if (typeof call.input.orientation === 'string') {
            spec.orientation = call.input.orientation as HangulPageSetupSpec['orientation']
          }
          if (typeof call.input.paper === 'string') spec.paper = call.input.paper as HangulPaper
          if (call.input.marginTop != null) spec.marginTop = Number(call.input.marginTop)
          if (call.input.marginBottom != null) spec.marginBottom = Number(call.input.marginBottom)
          if (call.input.marginLeft != null) spec.marginLeft = Number(call.input.marginLeft)
          if (call.input.marginRight != null) spec.marginRight = Number(call.input.marginRight)
          if (call.input.columns != null) spec.columns = Number(call.input.columns)
          if (call.input.columnSpacing != null) spec.columnSpacing = Number(call.input.columnSpacing)
          const result = await deps.setPage(spec)
          return {
            output: `Set page: ${result.applied.join(', ')}.`,
            mutated: true,
            summary: t('aiToolSetPageDone'),
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: t('aiToolSetPage'),
          }
        }
      }
      return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
    },
    verifyResponse: (finalText, executed) => {
      const formatOk = executed.some((call) => call.name === 'apply_format' && call.ok)
      const claimedTextFormat =
        /굵게|기울임|밑줄|취소선|가운데|서식|fontSize|apply_format|\b\d+\s*pt\b/i.test(finalText)
      const claimedTableStyle =
        /(표|칸|셀|행|헤더).{0,10}배경|배경색|테두리|세로\s*정렬|표\s*너비|style_table|fillColor|cell fill/i.test(
          finalText,
        )
      const claimedFormat =
        /서식|굵게|기울임|밑줄|취소선|가운데|정렬했|지정했|formatted|fontSize|apply_format|\b\d+\s*pt\b/i.test(
          finalText,
        )
      // "지정했습니다" is also how the model reports cell fill. That is style_table, not apply_format.
      if (claimedFormat && !formatOk && !(claimedTableStyle && !claimedTextFormat)) {
        const claimedTableCell =
          /표\s*(의\s*)?(첫|헤더|머리)|table\s*(header|first)\s*row|첫\s*행|셀.*(굵|서식|pt)|first row.*(bold|13)/i.test(
            finalText,
          )
        if (executed.some((call) => call.name === 'apply_format' && !call.ok)) {
          return claimedTableCell
            ? 'apply_format failed. For a table row or cell, call apply_format with table and row from get_tables (omit col for the whole row). Do not claim formatting succeeded.'
            : 'apply_format failed. Call get_paragraphs and apply_format on the first editable paragraph — the title is not index 0 when that row is locked. ' +
              'Do not claim formatting succeeded.'
        }
        if (claimedTableCell) {
          return (
            'You claimed to format a table row or cell, but apply_format did not run. Call apply_format with table and row from get_tables (omit col to format the whole row). ' +
            'Do not format a body paragraph instead, and do not tell the user to do it by hand.'
          )
        }
        return (
          'You claimed to format text, but apply_format did not run. Call apply_format on the first written paragraph from insert_content (the starting index in that tool result), then describe only what actually changed.'
        )
      }
      const claimedTableEdit =
        /행을\s*(추가|넣|지)|열을\s*(추가|넣|지)|칸을\s*(추가|지)|셀을\s*(합|나누)|insert_row|delete_row|merge/i.test(
          finalText,
        )
      if (claimedTableEdit && !executed.some((call) => call.name === 'edit_table' && call.ok)) {
        return 'You claimed to change table structure, but edit_table did not succeed. Call edit_table; do not rebuild the table with insert_table and do not tell the user to do it by hand.'
      }
      if (claimedTableStyle && !executed.some((call) => call.name === 'style_table' && call.ok)) {
        return 'You claimed to change table fill/border/width, but style_table did not succeed. Call style_table. Bold/size still uses apply_format.'
      }
      const claimedPage =
        /가로\s*용지|세로\s*용지|쪽\s*설정|여백을|2단|landscape|portrait|set_page/i.test(finalText)
      if (claimedPage && !executed.some((call) => call.name === 'set_page' && call.ok)) {
        return 'You claimed to change page setup, but set_page did not succeed. Call set_page. Do not tell the user to open 쪽 설정.'
      }
      const claimed = /바꿨|수정했|고쳤|넣었|삽입했|filled|replaced|rewrote|inserted|formatted|edited the (document|paragraph|selection|cell|field|table)/i.test(
        finalText,
      )
      if (!claimed) return null
      if (executed.some((call) => MUTATING_TOOLS.includes(call.name) && call.ok)) return null
      if (executed.some((call) => MUTATING_TOOLS.includes(call.name) && !call.ok)) {
        return (
          'An edit tool failed after it may already have written. Call get_paragraphs and describe only what is actually in the document. ' +
          'Do not tell the user the document is unchanged unless the list still has the pre-edit text.'
        )
      }
      return 'You claimed to edit the document, but no edit tool ran. Tell the user the document was not changed and do not claim an edit.'
    },
  }
}
