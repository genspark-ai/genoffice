import { describe, expect, it } from 'vitest'
import {
  LOAD_SKILL_TOOL,
  buildInstructionsPrompt,
  buildRulesPrompt,
  buildSkillsPrompt,
  coerceScope,
  parseSkillMarkdown,
  scopeApplies,
  serializeSkillMarkdown,
  skillBodyForTool,
  skillsForSurface,
  type UserSkill,
} from '../src/instructions'

function skill(over: Partial<UserSkill> = {}): UserSkill {
  return {
    id: 'brand-voice',
    name: 'Brand voice',
    description: 'How to phrase customer-facing copy',
    scopes: ['global'],
    body: '1. Use active voice.\n2. Never say "synergy".',
    enabled: true,
    ...over,
  }
}

describe('coerceScope', () => {
  it('maps the words users actually type onto the canonical scopes', () => {
    expect(coerceScope('slides')).toBe('pptx')
    expect(coerceScope('slide')).toBe('pptx')
    expect(coerceScope('ppt')).toBe('pptx')
    expect(coerceScope('word')).toBe('docx')
    expect(coerceScope('excel')).toBe('sheets')
    expect(coerceScope('xlsx')).toBe('sheets')
    expect(coerceScope('PDF')).toBe('pdf')
  })

  it('falls back to global rather than dropping an unrecognised scope', () => {
    expect(coerceScope('nonsense')).toBe('global')
    expect(coerceScope(undefined)).toBe('global')
  })
})

describe('parseSkillMarkdown', () => {
  it('reads front matter and keeps the body', () => {
    const parsed = parseSkillMarkdown(
      `---
name: Deck review
description: Checks a deck before export
scopes: pptx, docx
enabled: true
---

## Steps
1. Check contrast.`,
      'fallback',
    )
    expect(parsed.name).toBe('Deck review')
    expect(parsed.description).toBe('Checks a deck before export')
    expect(parsed.scopes).toEqual(['pptx', 'docx'])
    expect(parsed.enabled).toBe(true)
    expect(parsed.body.startsWith('## Steps')).toBe(true)
  })

  it('accepts a bare markdown file, inferring name and description', () => {
    const parsed = parseSkillMarkdown('# Invoice layout\n\nHow we lay out invoices.\n', 'invoice')
    expect(parsed.name).toBe('Invoice layout')
    expect(parsed.description).toBe('How we lay out invoices.')
    // nothing declared a scope, so it applies everywhere
    expect(parsed.scopes).toEqual(['global'])
    expect(parsed.enabled).toBe(true)
  })

  it('honours enabled: false and a bracketed scope list', () => {
    const parsed = parseSkillMarkdown(
      '---\nname: X\nscopes: [sheets, "slides"]\nenabled: false\n---\nbody',
      'x',
    )
    expect(parsed.scopes).toEqual(['sheets', 'pptx'])
    expect(parsed.enabled).toBe(false)
  })

  it('round-trips through serialize', () => {
    const original = skill({ scopes: ['docx', 'sheets'], enabled: false })
    const back = parseSkillMarkdown(serializeSkillMarkdown(original), original.id)
    expect(back).toEqual(original)
  })
})

describe('scope selection', () => {
  it('global applies to every surface', () => {
    expect(scopeApplies(['global'], 'docx')).toBe(true)
    expect(scopeApplies(['global'], 'sheets')).toBe(true)
  })

  it('a scoped skill stays out of the other surfaces', () => {
    expect(scopeApplies(['pptx'], 'pptx')).toBe(true)
    expect(scopeApplies(['pptx'], 'docx')).toBe(false)
  })

  it('disabled skills are filtered out even when in scope', () => {
    const list = [skill({ id: 'on' }), skill({ id: 'off', enabled: false })]
    expect(skillsForSurface(list, 'docx').map((s) => s.id)).toEqual(['on'])
  })
})

describe('prompt assembly', () => {
  it('puts the global rules before the surface-specific ones', () => {
    const prompt = buildRulesPrompt({ global: 'Be terse.', docx: 'Use Heading 1.' }, 'docx')
    expect(prompt.indexOf('Be terse.')).toBeLessThan(prompt.indexOf('Use Heading 1.'))
  })

  it('leaves out another surface’s rules', () => {
    const prompt = buildRulesPrompt({ sheets: 'Freeze the header row.' }, 'docx')
    expect(prompt).toBe('')
  })

  it('advertises only titles, not bodies — the point of on-demand loading', () => {
    const prompt = buildSkillsPrompt([skill()], 'docx')
    expect(prompt).toContain('brand-voice')
    expect(prompt).toContain('Brand voice')
    expect(prompt).toContain(LOAD_SKILL_TOOL)
    expect(prompt).not.toContain('synergy')
  })

  it('emits nothing when no skill applies, so no empty heading reaches the prompt', () => {
    expect(buildSkillsPrompt([skill({ scopes: ['sheets'] })], 'docx')).toBe('')
    expect(buildInstructionsPrompt({}, [], 'docx')).toBe('')
  })
})

describe('skillBodyForTool', () => {
  it('returns the body once the agent asks for it', () => {
    expect(skillBodyForTool([skill()], 'docx', 'brand-voice')).toContain('synergy')
  })

  it('refuses a skill that is out of scope and lists what is available', () => {
    const out = skillBodyForTool([skill({ id: 'deck', scopes: ['pptx'] })], 'docx', 'deck')
    expect(out).toContain('No skill "deck"')
  })
})
