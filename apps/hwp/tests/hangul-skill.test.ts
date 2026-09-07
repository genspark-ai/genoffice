import { describe, expect, it } from 'vitest'
import { SELECTION_PREVIEW_CHARS } from '../src/renderer/studio-text'
import { createHangulSkill, type HangulSkillDeps } from '../src/renderer/ai/hangul-skill'

function deps(partial: Partial<HangulSkillDeps> = {}): HangulSkillDeps {
  return {
    fileName: () => 'memo.hwp',
    pageCount: () => 2,
    hasSelection: () => false,
    selectionPreview: () => null,
    getDocumentText: async () => '문서 본문',
    getSelection: async () => null,
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
    expect(ctx).toContain('Editing tools are not available')
    expect(skill.systemPrompt).toMatch(/no editing tools/i)
  })

  it('includes a selection preview when present', () => {
    const skill = createHangulSkill(() =>
      deps({ hasSelection: () => true, selectionPreview: () => '선택 문장' }),
    )
    expect(skill.buildContext?.()).toContain('선택 문장')
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
