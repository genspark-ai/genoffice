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
    getDocumentText: async () => '문서 본문',
    getSelection: async () => null,
    replaceParagraph: async (text) => ({ before: 'old', after: text }),
    replaceSelection: async (text) => ({ before: 'sel', after: text }),
    listParagraphs: async () => [
      { index: 0, editable: true, reason: null, section: 0, paragraph: 0, text: '첫 문단' },
    ],
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
    expect(skill.systemPrompt).not.toMatch(/no editing tools/i)
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
