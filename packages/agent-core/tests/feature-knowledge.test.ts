import { describe, expect, it } from 'vitest'
import {
  describeKnowledge,
  filterKnowledgeEntries,
  knowledgeGlobMatches,
  knowledgeIn,
  validateKnowledge,
} from '../src/feature-knowledge'
import type { FeatureRecord } from '../src/features-skill'

function record(id: string, knowledge?: FeatureRecord['knowledge']): FeatureRecord {
  return {
    id,
    name: id,
    description: 'x',
    kind: knowledge ? 'knowledge' : '',
    ...(knowledge ? { knowledge } : {}),
    createdAt: 'x',
    updatedAt: 'x',
  }
}

describe('knowledge contract', () => {
  it('accepts a folder and a file, with an optional label', () => {
    expect(validateKnowledge({ source: 'folder', path: '/Users/me/notes' }).ok).toBe(true)
    expect(
      validateKnowledge({ source: 'file', path: '/Users/me/notes/style.md', label: 'House style' })
        .ok,
    ).toBe(true)
  })

  it('refuses a relative path, a home shortcut and an address', () => {
    for (const path of ['notes/style.md', '~/notes', 'https://example.com/style.md']) {
      const result = validateKnowledge({ source: 'folder', path })
      expect(result.ok).toBe(false)
      if (!result.ok) expect(result.errors[0]!.field).toBe('knowledge.path')
    }
  })

  it('refuses an invented source, an unknown field and a long label', () => {
    expect(validateKnowledge({ source: 'database', path: '/tmp/x' }).ok).toBe(false)
    const extra = validateKnowledge({ source: 'file', path: '/tmp/x', url: 'https://x' })
    expect(extra.ok).toBe(false)
    if (!extra.ok) expect(extra.errors[0]!.field).toBe('knowledge.url')
    expect(validateKnowledge({ source: 'file', path: '/tmp/x', label: 'x'.repeat(61) }).ok).toBe(
      false,
    )
    expect(validateKnowledge(null).ok).toBe(false)
    expect(validateKnowledge({ source: 'file' }).ok).toBe(false)
  })

  it('accepts bounded include/exclude name rules and echoes them trimmed', () => {
    const result = validateKnowledge({
      source: 'folder',
      path: '/Users/me/notes',
      include: ['*.md', ' report-* '],
      exclude: ['*-draft.md'],
    })
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.knowledge.include).toEqual(['*.md', 'report-*'])
      expect(result.knowledge.exclude).toEqual(['*-draft.md'])
    }
  })

  it('refuses malformed file rules', () => {
    for (const bad of [
      { include: [] },
      { include: ['*.md', ''] },
      { include: ['a/b*.md'] },
      { include: 'not-an-array' },
      { exclude: [Array(130).fill('x').join('')] },
    ]) {
      const result = validateKnowledge({ source: 'folder', path: '/tmp/x', ...bad })
      expect(result.ok).toBe(false)
    }
  })

  it('matches name globs: * spans, ? is one, case-insensitive, names never paths', () => {
    expect(knowledgeGlobMatches('*.md', 'style.md')).toBe(true)
    expect(knowledgeGlobMatches('*.MD', 'style.md')).toBe(true)
    expect(knowledgeGlobMatches('report-?', 'report-a.md')).toBe(false)
    expect(knowledgeGlobMatches('report-?.md', 'report-a.md')).toBe(true)
    expect(knowledgeGlobMatches('*draft*', 'Q3-draft.md')).toBe(true)
    expect(knowledgeGlobMatches('style.md', 'notes/style.md')).toBe(false)
  })

  it('filters entries by the rules: include when declared, exclude always', () => {
    const knowledge = {
      source: 'folder' as const,
      path: '/tmp/x',
      include: ['*.md'],
      exclude: ['*-draft.md'],
    }
    const kept = filterKnowledgeEntries(knowledge, [
      { name: 'style.md' },
      { name: 'Q3-draft.md' },
      { name: 'notes.txt' },
    ])
    expect(kept.map((entry) => entry.name)).toEqual(['style.md'])
    const unfiltered = filterKnowledgeEntries({ source: 'folder', path: '/tmp/x' }, [
      { name: 'style.md' },
      { name: 'notes.txt' },
    ])
    expect(unfiltered).toHaveLength(2)
  })

  it('accepts the index mode and refuses an invented one', () => {
    const names = validateKnowledge({ source: 'folder', path: '/tmp/x', index: 'names' })
    expect(names.ok).toBe(true)
    if (names.ok) expect(names.knowledge.index).toBe('names')
    const bad = validateKnowledge({ source: 'folder', path: '/tmp/x', index: 'embedding' })
    expect(bad.ok).toBe(false)
    if (!bad.ok) expect(bad.errors[0]!.message).toContain('full-text, names')
  })

  it('shows a names-only source as such in the assistant context', () => {
    const features = [record('archive', { source: 'folder', path: '/tmp/a', index: 'names' })]
    expect(describeKnowledge(features)).toEqual(['archive | archive | folder (names only)'])
  })

  it('shows the file rules in the assistant context line', () => {
    const features = [
      record('notes', {
        source: 'folder',
        path: '/tmp/notes',
        include: ['*.md'],
        exclude: ['*-draft.md'],
      }),
    ]
    expect(describeKnowledge(features)).toEqual([
      'notes | notes | folder (only *.md, not *-draft.md)',
    ])
  })

  it('lists the sources a registry holds, without their contents', () => {
    const features = [
      record('house-style', { source: 'file', path: '/tmp/style.md', label: 'House style' }),
      record('notes', { source: 'folder', path: '/tmp/notes' }),
      record('plain'),
    ]
    expect(knowledgeIn(features).map((f) => f.id)).toEqual(['house-style', 'notes'])
    expect(describeKnowledge(features)).toEqual([
      'house-style | House style | file',
      'notes | notes | folder',
    ])
  })
})
