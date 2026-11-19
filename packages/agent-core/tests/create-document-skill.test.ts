import { describe, expect, it } from 'vitest'
import {
  CREATE_DOCUMENT_MAX_CHARS,
  createDocumentSkill,
  type CreateDocumentRequest,
} from '../src/create-document-skill'
import { createKnowledgeBaseSkill, kbCiteHref } from '../src/knowledge-base-skill'

function makeSkill(create?: (r: CreateDocumentRequest) => Promise<{ path: string }>) {
  const requests: CreateDocumentRequest[] = []
  const skill = createDocumentSkill({
    create: async (r) => {
      requests.push(r)
      return create ? create(r) : { path: `/Users/me/GenOffice/${r.title}.${r.format}` }
    },
  })
  return { skill, requests }
}

const call = (input: Record<string, unknown>) => ({ id: 't1', name: 'create_document', input })

describe('createDocumentSkill', () => {
  it('writes through deps and returns the path with its cite link', async () => {
    const { skill, requests } = makeSkill()
    const result = await skill.executeTool(
      call({ title: 'Q3 summary', format: 'docx', content: '# Q3\n\ntext' }),
    )
    expect(result.isError).toBeUndefined()
    expect(requests).toEqual([{ title: 'Q3 summary', format: 'docx', content: '# Q3\n\ntext' }])
    expect(result.output).toContain('/Users/me/GenOffice/Q3 summary.docx')
    expect(result.output).toContain(kbCiteHref('/Users/me/GenOffice/Q3 summary.docx'))
    expect(skill.createdPaths()).toEqual(['/Users/me/GenOffice/Q3 summary.docx'])
  })

  it('rejects a bad format, empty content and oversized content before writing', async () => {
    const { skill, requests } = makeSkill()
    expect(
      (await skill.executeTool(call({ title: 'a', format: 'pdf', content: 'x' }))).isError,
    ).toBe(true)
    expect(
      (await skill.executeTool(call({ title: 'a', format: 'md', content: '  ' }))).isError,
    ).toBe(true)
    const huge = 'x'.repeat(CREATE_DOCUMENT_MAX_CHARS + 1)
    expect(
      (await skill.executeTool(call({ title: 'a', format: 'md', content: huge }))).isError,
    ).toBe(true)
    expect(requests).toEqual([])
  })

  it('defaults a missing title and reports a write failure as a tool error', async () => {
    const { skill, requests } = makeSkill(async () => {
      throw new Error('disk full')
    })
    const result = await skill.executeTool(call({ format: 'md', content: 'x' }))
    expect(requests[0]!.title).toBe('Untitled')
    expect(result.isError).toBe(true)
    expect(result.output).toContain('disk full')
    expect(skill.createdPaths()).toEqual([])
  })

  it('verifyResponse pushes back on a saved claim without a successful call', () => {
    const { skill } = makeSkill()
    expect(skill.verifyResponse!('I have saved the summary as report.docx.', [])).toContain(
      'create_document',
    )
    expect(
      skill.verifyResponse!('I have saved the summary.', [{ name: 'create_document', ok: true }]),
    ).toBeNull()
    expect(
      skill.verifyResponse!('The handbook was created in 2019.', [
        { name: 'read_knowledge_file', ok: true },
      ]),
    ).toBeNull()
    expect(skill.verifyResponse!('The document has been saved as notes.md.', [])).toContain(
      'create_document',
    )
    expect(skill.verifyResponse!('The lease runs until 2027.', [])).toBeNull()
    for (const fact of [
      'The document was created by HR and the file was generated monthly.',
      'I created a table from the handbook figures below.',
      'The report says the backup is saved nightly.',
    ]) {
      expect(skill.verifyResponse!(fact, [])).toBeNull()
    }
    expect(skill.verifyResponse!('I saved it to the GenOffice folder.', [])).toContain(
      'create_document',
    )
    expect(skill.verifyResponse!('Saved as q3-summary.docx in your folder.', [])).toContain(
      'create_document',
    )
  })

  it('the knowledge-base citation guard accepts a path the document skill wrote', async () => {
    const { skill } = makeSkill()
    const kb = createKnowledgeBaseSkill({
      listFiles: () => [],
      search: async () => [],
      read: async () => ({ text: '', totalChars: 0, offset: 0 }),
      alsoCitable: () => skill.createdPaths(),
    })
    const href = kbCiteHref('/Users/me/GenOffice/notes.md')
    expect(kb.verifyResponse!(`Saved: [notes.md](${href})`, [])).not.toBeNull()
    await skill.executeTool(call({ title: 'notes', format: 'md', content: '- a' }))
    expect(kb.verifyResponse!(`Saved: [notes.md](${href})`, [])).toBeNull()
  })
})
