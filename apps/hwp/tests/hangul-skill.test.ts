import { describe, expect, it } from 'vitest'
import { SELECTION_PREVIEW_CHARS } from '../src/renderer/studio-text'
import { createHangulSkill, type HangulSkillDeps } from '../src/renderer/ai/hangul-skill'

function deps(partial: Partial<HangulSkillDeps> = {}): HangulSkillDeps {
  return {
    fileName: () => 'memo.hwp',
    pageCount: () => 2,
    currentPage: () => null,
    hasSelection: () => false,
    selectionPreview: () => null,
    paragraphPreview: () => null,
    getDocumentText: async () => '문서 본문',
    getSelection: async () => null,
    replaceParagraph: async (text) => ({ before: 'old', after: text }),
    replaceSelection: async (text) => ({ before: 'sel', after: text }),
    listParagraphs: async () => [
      { index: 0, editable: true, reason: null, section: 0, paragraph: 0, text: '첫 문단' },
    ],
    insertContent: async (_text, afterIndex) => ({ count: 1, start: afterIndex == null ? 0 : afterIndex + 1 }),
    listFields: async () => [{ name: '기안자', value: '', type: 'clickhere' }],
    setField: async (name, value) => ({ name, before: '', after: value }),
    listTables: async () => [
      {
        index: 0,
        section: 0,
        paragraph: 1,
        control: 0,
        rows: 1,
        cols: 1,
        cells: [{ index: 0, row: 0, col: 0, text: '칸' }],
      },
    ],
    replaceCell: async (_table, _row, _col, text) => ({ before: '칸', after: text }),
    insertTable: async (rows, cols) => ({ table: 0, rows, cols, unfilled: [] }),
    applyFormat: async (_format, index, indexes) => ({
      indexes: indexes ?? [index ?? 0],
      applied: ['bold=true'],
    }),
    ...partial,
  }
}

describe('createHangulSkill', () => {
  it('describes a read-only Hangul document in context', () => {
    const skill = createHangulSkill(() => deps())
    const ctx = skill.buildContext?.() ?? ''
    expect(ctx).toContain('memo.hwp')
    expect(ctx).toContain('2 page')
    expect(ctx).toContain('No text is selected')
    expect(ctx).toContain('replace_paragraph')
    expect(ctx).not.toContain('Current page:')
    expect(skill.systemPrompt).toMatch(/replace_paragraph/)
    expect(skill.systemPrompt).toMatch(/insert_content/)
    expect(skill.systemPrompt).toMatch(/one insert_content/)
    expect(skill.systemPrompt).toMatch(/# Intent resolution/)
    expect(skill.systemPrompt).toMatch(/HG-2/)
    expect(skill.systemPrompt).toMatch(/insert_table/)
    expect(skill.systemPrompt).toMatch(/apply_format/)
    expect(skill.systemPrompt).toMatch(/HG-5/)
    expect(skill.systemPrompt).not.toMatch(/no editing tools/i)
    expect(skill.systemPrompt).not.toMatch(/You cannot create new body paragraphs/)
  })

  it('includes a paragraph skeleton like Docs block lists', () => {
    const skill = createHangulSkill(() =>
      deps({
        paragraphPreview: () => [
          { index: 0, editable: false, reason: 'control', section: 0, paragraph: 0, text: '' },
          { index: 1, editable: true, reason: null, section: 0, paragraph: 1, text: '프로젝트 계획서' },
        ],
      }),
    )
    const ctx = skill.buildContext?.() ?? ''
    expect(ctx).toContain('Body paragraphs (2; index|status|preview):')
    expect(ctx).toContain('0|locked:control|(empty)')
    expect(ctx).toContain('1|editable|프로젝트 계획서')
  })

  it('marks a locked-only body as blank so drafts use insert_content', () => {
    const skill = createHangulSkill(() =>
      deps({
        paragraphPreview: () => [
          { index: 0, editable: false, reason: 'control', section: 0, paragraph: 0, text: '' },
        ],
      }),
    )
    const ctx = skill.buildContext?.() ?? ''
    expect(ctx).toContain('The body looks blank')
    expect(ctx).toContain('one insert_content')
    expect(ctx).toContain('0|locked:control|(empty)')
  })

  it('includes the current page when known', () => {
    const skill = createHangulSkill(() => deps({ currentPage: () => 2 }))
    expect(skill.buildContext?.()).toContain('Current page: 2')
  })

  it('includes a selection preview when present', () => {
    const skill = createHangulSkill(() =>
      deps({ hasSelection: () => true, selectionPreview: () => '선택 문장' }),
    )
    const ctx = skill.buildContext?.() ?? ''
    expect(ctx).toContain('선택 문장')
    expect(ctx).toContain('replace_selection')
  })

  it('reads document and selection through tools', async () => {
    const skill = createHangulSkill(() =>
      deps({
        hasSelection: () => true,
        selectionPreview: () => 'sel',
        getSelection: async () => 'sel',
      }),
    )
    const doc = await skill.executeTool({ id: '1', name: 'get_document_text', input: {} })
    expect(doc.output).toBe('문서 본문')
    expect(doc.isError).toBeUndefined()
    const sel = await skill.executeTool({ id: '2', name: 'get_selection', input: {} })
    expect(sel.output).toBe('sel')
  })

  it('returns the empty-document contract', async () => {
    const skill = createHangulSkill(() => deps({ getDocumentText: async () => '' }))
    const doc = await skill.executeTool({ id: '1', name: 'get_document_text', input: {} })
    expect(doc.output).toBe('(empty document)')
    expect(doc.isError).toBeUndefined()
  })

  it('rejects unknown tools', async () => {
    const skill = createHangulSkill(() => deps())
    const result = await skill.executeTool({ id: '9', name: 'apply_text', input: {} })
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Unknown tool')
  })

  it('replaces the current paragraph', async () => {
    const skill = createHangulSkill(() =>
      deps({
        replaceParagraph: async (text) => ({ before: '안녕', after: text }),
      }),
    )
    const result = await skill.executeTool({
      id: '3',
      name: 'replace_paragraph',
      input: { text: '안녕하세요' },
    })
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('안녕')
    expect(result.output).toContain('안녕하세요')
  })

  it('accepts a string paragraph index from the model', async () => {
    let seen: number | undefined
    const skill = createHangulSkill(() =>
      deps({
        replaceParagraph: async (text, index) => {
          seen = index
          return { before: 'old', after: text }
        },
      }),
    )
    const result = await skill.executeTool({
      id: '3',
      name: 'replace_paragraph',
      input: { text: '새', index: '1' },
    })
    expect(result.isError).toBeUndefined()
    expect(seen).toBe(1)
    expect(result.output).toContain('[1]')
  })

  it('inserts paragraphs and treats a blank afterIndex as the caret', async () => {
    let seen: number | undefined
    const skill = createHangulSkill(() =>
      deps({
        insertContent: async (text, afterIndex) => {
          seen = afterIndex
          expect(text).toBe('안녕\n세상')
          return { count: 2, start: 1 }
        },
      }),
    )
    const result = await skill.executeTool({
      id: '9',
      name: 'insert_content',
      input: { text: '안녕\n세상', afterIndex: '' },
    })
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(seen).toBeUndefined()
    expect(result.output).toContain('2 paragraph')
    expect(result.output).toContain('[1]')
  })

  it('rejects a blank afterIndex that is only whitespace', async () => {
    let called = false
    const skill = createHangulSkill(() =>
      deps({
        insertContent: async () => {
          called = true
          return { count: 1, start: 0 }
        },
      }),
    )
    const result = await skill.executeTool({
      id: '9',
      name: 'insert_content',
      input: { text: '안녕', afterIndex: '  ' },
    })
    expect(result.isError).toBe(true)
    expect(result.output).toMatch(/must be an integer/)
    expect(called).toBe(false)
  })

  it('rejects a null or blank cell index instead of writing cell 0', async () => {
    let called = false
    const skill = createHangulSkill(() =>
      deps({
        replaceCell: async () => {
          called = true
          return { before: '칸', after: 'x' }
        },
      }),
    )
    for (const input of [
      { table: null, row: 0, col: 0, text: 'x' },
      { table: 0, row: '', col: 0, text: 'x' },
      { table: 0, row: 0, col: '  ', text: 'x' },
    ]) {
      const result = await skill.executeTool({ id: '8', name: 'replace_cell', input })
      expect(result.isError).toBe(true)
      expect(result.output).toMatch(/must be an integer/)
    }
    expect(called).toBe(false)
  })

  it('surfaces replace_paragraph failures', async () => {
    const skill = createHangulSkill(() =>
      deps({
        replaceParagraph: async () => {
          throw new Error('paragraph is not editable')
        },
      }),
    )
    const result = await skill.executeTool({
      id: '3',
      name: 'replace_paragraph',
      input: { text: 'x' },
    })
    expect(result.isError).toBe(true)
    expect(result.output).toBe('paragraph is not editable')
  })

  it('corrects a claimed edit that never ran', () => {
    const skill = createHangulSkill(() => deps())
    expect(skill.verifyResponse?.('문단을 바꿨습니다.', [])).toMatch(/was not changed/i)
    expect(
      skill.verifyResponse?.('문단을 바꿨습니다.', [{ name: 'replace_paragraph', ok: true }]),
    ).toBeNull()
    expect(
      skill.verifyResponse?.('선택한 부분을 바꿨습니다.', [{ name: 'replace_selection', ok: true }]),
    ).toBeNull()
    expect(skill.verifyResponse?.('문단을 넣었습니다.', [])).toMatch(/was not changed/i)
    expect(
      skill.verifyResponse?.('Inserted two paragraphs.', [{ name: 'insert_content', ok: true }]),
    ).toBeNull()
    expect(
      skill.verifyResponse?.('문단을 바꿨습니다.', [{ name: 'replace_paragraph', ok: false }]),
    ).toMatch(/get_paragraphs/i)
    expect(
      skill.verifyResponse?.('문단을 바꿨습니다.', [{ name: 'replace_paragraph', ok: false }]),
    ).not.toMatch(/was not changed/i)
    expect(
      skill.verifyResponse?.('표를 넣었습니다.', [{ name: 'insert_table', ok: true }]),
    ).toBeNull()
    expect(
      skill.verifyResponse?.('제목을 굵게 했습니다.', [{ name: 'apply_format', ok: true }]),
    ).toBeNull()
  })

  it('replaces a selection and a field', async () => {
    const skill = createHangulSkill(() => deps())
    const sel = await skill.executeTool({
      id: '4',
      name: 'replace_selection',
      input: { text: '새 선택' },
    })
    expect(sel.mutated).toBe(true)
    expect(sel.output).toContain('새 선택')
    const field = await skill.executeTool({
      id: '5',
      name: 'set_field',
      input: { name: '기안자', value: '홍길동' },
    })
    expect(field.mutated).toBe(true)
    expect(field.output).toContain('홍길동')
  })

  it('lists paragraphs and tables', async () => {
    const skill = createHangulSkill(() => deps())
    const paras = await skill.executeTool({ id: '6', name: 'get_paragraphs', input: {} })
    expect(paras.output).toContain('[0]')
    expect(paras.output).toContain('첫 문단')
    const tables = await skill.executeTool({ id: '7', name: 'get_tables', input: {} })
    expect(tables.output).toContain('table[0]')
    const cell = await skill.executeTool({
      id: '8',
      name: 'replace_cell',
      input: { table: 0, row: 0, col: 0, text: '새 칸' },
    })
    expect(cell.mutated).toBe(true)
    expect(cell.output).toContain('새 칸')
  })

  it('inserts a table and applies format', async () => {
    let tableArgs: unknown[] | null = null
    let formatArgs: unknown[] | null = null
    const skill = createHangulSkill(() =>
      deps({
        insertTable: async (rows, cols, cells, afterIndex) => {
          tableArgs = [rows, cols, cells, afterIndex]
          return { table: 1, rows, cols, unfilled: [] }
        },
        applyFormat: async (format, index, indexes) => {
          formatArgs = [format, index, indexes]
          return { indexes: indexes ?? [index ?? 0], applied: ['bold=true', 'align=center', 'color=FF0000'] }
        },
      }),
    )
    const table = await skill.executeTool({
      id: '10',
      name: 'insert_table',
      input: { rows: 2, cols: 3, cells: [['a', 'b', 'c']], afterIndex: '' },
    })
    expect(table.isError).toBeUndefined()
    expect(table.mutated).toBe(true)
    expect(table.output).toContain('table[1]')
    expect(table.output).toContain('2x3')
    expect(tableArgs).toEqual([2, 3, [['a', 'b', 'c']], undefined])
    const format = await skill.executeTool({
      id: '11',
      name: 'apply_format',
      input: { index: '0', bold: true, align: 'center', color: 'FF0000', lineSpacing: 1.5 },
    })
    expect(format.isError).toBeUndefined()
    expect(format.mutated).toBe(true)
    expect(format.output).toContain('bold=true')
    expect(formatArgs?.[1]).toBe(0)
    expect(formatArgs?.[0]).toMatchObject({ bold: true, align: 'center', color: 'FF0000', lineSpacing: 1.5 })
  })

  it('rejects a blank table size instead of writing a 0x0 table', async () => {
    let called = false
    const skill = createHangulSkill(() =>
      deps({
        insertTable: async () => {
          called = true
          return { table: 0, rows: 1, cols: 1, unfilled: [] }
        },
      }),
    )
    const result = await skill.executeTool({
      id: '10',
      name: 'insert_table',
      input: { rows: '', cols: 2 },
    })
    expect(result.isError).toBe(true)
    expect(result.output).toMatch(/must be an integer/)
    expect(called).toBe(false)
  })

  it('clips a long selection preview in context', () => {
    const preview = '가'.repeat(SELECTION_PREVIEW_CHARS + 20)
    const skill = createHangulSkill(() =>
      deps({ hasSelection: () => true, selectionPreview: () => preview }),
    )
    const ctx = skill.buildContext?.() ?? ''
    expect(ctx).toContain('가'.repeat(SELECTION_PREVIEW_CHARS))
    expect(ctx).toContain('…')
    expect(ctx).not.toContain(preview)
  })

  it('surfaces getDocumentText failures', async () => {
    const skill = createHangulSkill(() =>
      deps({
        getDocumentText: async () => {
          throw new Error('plain text unavailable')
        },
      }),
    )
    const doc = await skill.executeTool({ id: '1', name: 'get_document_text', input: {} })
    expect(doc.isError).toBe(true)
    expect(doc.output).toBe('plain text unavailable')
  })
})
