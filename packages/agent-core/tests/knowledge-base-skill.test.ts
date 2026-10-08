import { describe, expect, it } from 'vitest'
import { createKnowledgeBaseSkill, type KnowledgeBaseDeps } from '../src/knowledge-base-skill'

const CORPUS = [
  {
    path: '/Users/me/合同/租赁合同.docx',
    folder: '合同',
    name: '租赁合同.docx',
    ext: 'docx',
    sizeBytes: 2048,
  },
  {
    path: '/Users/me/工作/employee handbook.pdf',
    folder: '工作',
    name: 'employee handbook.pdf',
    ext: 'pdf',
    sizeBytes: 734_003_200,
  },
]

function makeDeps(overrides: Partial<KnowledgeBaseDeps> = {}): KnowledgeBaseDeps & {
  searches: Array<[string, number]>
  reads: Array<[string, number]>
} {
  const searches: Array<[string, number]> = []
  const reads: Array<[string, number]> = []
  return {
    searches,
    reads,
    listFiles: () => CORPUS,
    search: async (query, limit) => {
      searches.push([query, limit])
      return overrides.search
        ? overrides.search(query, limit)
        : [
            {
              path: CORPUS[0]!.path,
              name: CORPUS[0]!.name,
              ext: 'docx',
              folder: '合同',
              snippet: '试用期考核办法',
            },
          ]
    },
    read: async (path, offset) => {
      reads.push([path, offset])
      return overrides.read
        ? overrides.read(path, offset)
        : { text: '正文', totalChars: 100, offset }
    },
  }
}

describe('createKnowledgeBaseSkill', () => {
  it('builds the corpus listing with folder, name and the absolute paths for citations', async () => {
    const skill = createKnowledgeBaseSkill(makeDeps())
    const context = skill.buildContext!()
    expect(context).toContain('0 | 合同 | 租赁合同.docx | .docx | 2KB')
    expect(context).toContain('/Users/me/合同/租赁合同.docx')
    expect(context).toContain('/Users/me/工作/employee handbook.pdf')
  })

  it('an empty corpus contributes no context', () => {
    const deps = makeDeps()
    deps.listFiles = () => []
    const skill = createKnowledgeBaseSkill(deps)
    expect(skill.buildContext!()).toBe('')
  })

  it('a large corpus degrades to the name listing without the path block', () => {
    const deps = makeDeps()
    const big = Array.from({ length: 200 }, (_, i) => ({
      path: `/Users/me/folder-${i}/a-very-long-file-name-that-keeps-going-${i}.docx`,
      folder: `folder-${i}`,
      name: `a-very-long-file-name-that-keeps-going-${i}.docx`,
      ext: 'docx',
      sizeBytes: 1024,
    }))
    deps.listFiles = () => big
    const context = createKnowledgeBaseSkill(deps).buildContext!()
    // degraded: names only — no absolute-path block (search results carry paths)
    expect(context).not.toContain('Absolute paths')
    expect(context).not.toContain('/Users/me/')
    expect(context).toContain('0 | folder-0')
    expect(context.length).toBeLessThan(20_000)
  })

  it('search forwards the query and formats hits with folder, path and snippet', async () => {
    const deps = makeDeps()
    const skill = createKnowledgeBaseSkill(deps)
    const result = await skill.executeTool!({
      id: 't1',
      name: 'search_knowledge_base',
      input: { query: '试用期', limit: 5 },
    })
    expect(deps.searches).toEqual([['试用期', 5]])
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('合同/租赁合同.docx')
    expect(result.output).toContain('path: /Users/me/合同/租赁合同.docx')
    expect(result.output).toContain('试用期考核办法')
  })

  it('search teaches the synonym retry instead of dead-ending', async () => {
    const deps = makeDeps({ search: async () => [] })
    const skill = createKnowledgeBaseSkill(deps)
    const result = await skill.executeTool!({
      id: 't2',
      name: 'search_knowledge_base',
      input: { query: '差旅报销' },
    })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('Retry with synonyms')
  })

  it('read resolves the index against the corpus and pages with offsets', async () => {
    const deps = makeDeps()
    const skill = createKnowledgeBaseSkill(deps)
    const result = await skill.executeTool!({
      id: 't3',
      name: 'read_knowledge_file',
      input: { index: 0, offset: 40 },
    })
    expect(deps.reads).toEqual([['/Users/me/合同/租赁合同.docx', 40]])
    expect(result.output).toContain('characters 40..42 of 100')
    expect(result.output).toContain('(continue with offset)')
  })

  it('read rejects an out-of-range index as a tool error', async () => {
    const skill = createKnowledgeBaseSkill(makeDeps())
    const result = await skill.executeTool!({
      id: 't4',
      name: 'read_knowledge_file',
      input: { index: 9 },
    })
    expect(result.isError).toBe(true)
  })

  it('an empty search query is a tool error, not a corpus sweep', async () => {
    const deps = makeDeps()
    const skill = createKnowledgeBaseSkill(deps)
    const result = await skill.executeTool!({
      id: 't5',
      name: 'search_knowledge_base',
      input: { query: '   ' },
    })
    expect(result.isError).toBe(true)
    expect(deps.searches).toHaveLength(0)
  })

  it('verifyResponse accepts corpus citations and rejects invented ones', () => {
    const skill = createKnowledgeBaseSkill(makeDeps())
    const good =
      '来源见 [租赁合同.docx](filenav:///Users/me/%E5%90%88%E5%90%8C/%E7%A7%9F%E8%B5%81%E5%90%88%E5%90%8C.docx)'
    expect(skill.verifyResponse!(good, [])).toBeNull()
    const bad = '来源见 [编造.docx](filenav:///tmp/编造.docx)'
    const correction = skill.verifyResponse!(bad, [])
    expect(correction).toContain('not knowledge-base paths')
    expect(correction).toContain('filenav:///tmp/编造.docx')
  })

  it('verifyResponse lets answers without citations through', () => {
    const skill = createKnowledgeBaseSkill(makeDeps())
    expect(skill.verifyResponse!('知识库里没有相关内容。', [])).toBeNull()
  })
})
