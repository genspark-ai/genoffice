import { describe, expect, it } from 'vitest'
import {
  createKnowledgeDocument,
  KB_CREATE_MAX_CHARS,
  parseKbCreateRequest,
  type KbCreateDeps,
} from '../src/main/knowledge-base'

function makeDeps() {
  const written: Array<[string, string]> = []
  const converted: Array<[string, string]> = []
  const recent: string[] = []
  const deps: KbCreateDeps = {
    dir: () => '/save',
    uniquePath: (dir, name) => `${dir}/${name}`,
    sanitizeBase: (title) => title.replace(/[/:]/g, '_') || 'Untitled',
    writeText: async (path, text) => void written.push([path, text]),
    markdownToDocx: async (md, out) => void converted.push([md, out]),
    recordRecent: (path) => void recent.push(path),
  }
  return { deps, written, converted, recent }
}

describe('parseKbCreateRequest', () => {
  it('accepts md and docx and keeps the title as given', () => {
    expect(parseKbCreateRequest({ title: 'a', format: 'md', content: 'x' })).toEqual({
      title: 'a',
      format: 'md',
      content: 'x',
    })
    expect(parseKbCreateRequest({ format: 'docx', content: 'x' }).title).toBe('')
  })

  it('rejects other formats, blank and oversized content', () => {
    expect(() => parseKbCreateRequest({ format: 'pdf', content: 'x' })).toThrow(/format/)
    expect(() => parseKbCreateRequest({ format: 'md', content: ' \n' })).toThrow(/empty/)
    expect(() =>
      parseKbCreateRequest({ format: 'md', content: 'x'.repeat(KB_CREATE_MAX_CHARS + 1) }),
    ).toThrow(/long/)
    expect(() => parseKbCreateRequest(null)).toThrow(/format/)
  })
})

describe('createKnowledgeDocument', () => {
  it('writes markdown as text into the save folder and records it as recent', async () => {
    const { deps, written, converted, recent } = makeDeps()
    const result = await createKnowledgeDocument(deps, {
      title: 'Q3/notes',
      format: 'md',
      content: '# hi',
    })
    expect(result).toEqual({ path: '/save/Q3_notes.md' })
    expect(written).toEqual([['/save/Q3_notes.md', '# hi']])
    expect(converted).toEqual([])
    expect(recent).toEqual(['/save/Q3_notes.md'])
  })

  it('routes docx through the markdown converter', async () => {
    const { deps, written, converted, recent } = makeDeps()
    await createKnowledgeDocument(deps, { title: 'memo', format: 'docx', content: '# hi' })
    expect(converted).toEqual([['# hi', '/save/memo.docx']])
    expect(written).toEqual([])
    expect(recent).toEqual(['/save/memo.docx'])
  })

  it('does not record a file the converter failed to write', async () => {
    const { deps, recent } = makeDeps()
    deps.markdownToDocx = async () => {
      throw new Error('cli failed')
    }
    await expect(
      createKnowledgeDocument(deps, { title: 'memo', format: 'docx', content: '# hi' }),
    ).rejects.toThrow('cli failed')
    expect(recent).toEqual([])
  })
})
