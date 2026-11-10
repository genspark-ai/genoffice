import { describe, expect, it } from 'vitest'
import {
  createKnowledgeBaseSkill,
  kbCiteHref,
  type KnowledgeBaseDeps,
} from '../src/knowledge-base-skill'

const CORPUS = [
  {
    path: '/Users/me/\u5408\u540c/\u79df\u8d41\u5408\u540c.docx',
    folder: '\u5408\u540c',
    name: '\u79df\u8d41\u5408\u540c.docx',
    ext: 'docx',
    sizeBytes: 2048,
  },
  {
    path: '/Users/me/\u5de5\u4f5c/employee handbook.pdf',
    folder: '\u5de5\u4f5c',
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
              folder: '\u5408\u540c',
              snippet: '\u8bd5\u7528\u671f\u8003\u6838\u529e\u6cd5',
            },
          ]
    },
    read: async (path, offset) => {
      reads.push([path, offset])
      return overrides.read
        ? overrides.read(path, offset)
        : { text: '\u6b63\u6587', totalChars: 100, offset }
    },
  }
}

describe('createKnowledgeBaseSkill', () => {
  it('builds the corpus listing with folder, name and the absolute paths for citations', async () => {
    const skill = createKnowledgeBaseSkill(makeDeps())
    const context = skill.buildContext!()
    expect(context).toContain('0 | \u5408\u540c | \u79df\u8d41\u5408\u540c.docx | .docx | 2KB')
    expect(context).toContain('/Users/me/\u5408\u540c/\u79df\u8d41\u5408\u540c.docx')
    expect(context).toContain('/Users/me/\u5de5\u4f5c/employee handbook.pdf')
    expect(context).toContain(
      'cite: filenav:///Users/me/%E5%B7%A5%E4%BD%9C/employee%20handbook.pdf',
    )
  })

  it('hands the model ready-made cite links that survive spaces, parens and Windows paths', () => {
    expect(kbCiteHref('/Users/me/q (2).xlsx')).toBe('filenav:///Users/me/q%20%282%29.xlsx')
    expect(kbCiteHref('C:\\Users\\me\\report.docx')).toBe(
      'filenav://C%3A%5CUsers%5Cme%5Creport.docx',
    )
    expect(
      decodeURIComponent(kbCiteHref('C:\\Users\\me\\a b.docx').slice('filenav://'.length)),
    ).toBe('C:\\Users\\me\\a b.docx')
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
      input: { query: '\u8bd5\u7528\u671f', limit: 5 },
    })
    expect(deps.searches).toEqual([['\u8bd5\u7528\u671f', 5]])
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('\u5408\u540c/\u79df\u8d41\u5408\u540c.docx')
    expect(result.output).toContain('path: /Users/me/\u5408\u540c/\u79df\u8d41\u5408\u540c.docx')
    expect(result.output).toContain('\u8bd5\u7528\u671f\u8003\u6838\u529e\u6cd5')
  })

  it('search teaches the synonym retry instead of dead-ending', async () => {
    const deps = makeDeps({ search: async () => [] })
    const skill = createKnowledgeBaseSkill(deps)
    const result = await skill.executeTool!({
      id: 't2',
      name: 'search_knowledge_base',
      input: { query: '\u5dee\u65c5\u62a5\u9500' },
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
    expect(deps.reads).toEqual([['/Users/me/\u5408\u540c/\u79df\u8d41\u5408\u540c.docx', 40]])
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
      '\u6765\u6e90\u89c1 [\u79df\u8d41\u5408\u540c.docx](filenav:///Users/me/%E5%90%88%E5%90%8C/%E7%A7%9F%E8%B5%81%E5%90%88%E5%90%8C.docx)'
    expect(skill.verifyResponse!(good, [])).toBeNull()
    // an English path: the english-comments gate reads a CJK '//' URL as a comment
    const bad = 'cited [fake.docx](filenav:///tmp/fake.docx)'
    const correction = skill.verifyResponse!(bad, [])
    expect(correction).toContain('not knowledge-base paths')
    expect(correction).toContain('filenav:///tmp/fake.docx')
  })

  it('verifyResponse lets answers without citations through', () => {
    const skill = createKnowledgeBaseSkill(makeDeps())
    expect(
      skill.verifyResponse!(
        '\u77e5\u8bc6\u5e93\u91cc\u6ca1\u6709\u76f8\u5173\u5185\u5bb9\u3002',
        [],
      ),
    ).toBeNull()
  })
})

describe('createKnowledgeBaseSkill with no starred files', () => {
  it('an empty corpus mounts no tools and no prompt section; starring brings them back', () => {
    let files = [] as typeof CORPUS
    const skill = createKnowledgeBaseSkill({ ...makeDeps(), listFiles: () => files })
    expect(skill.tools).toEqual([])
    expect(skill.systemPrompt).toBe('')
    files = CORPUS
    expect(skill.tools.map((t) => t.name)).toEqual(['read_knowledge_file', 'search_knowledge_base'])
    expect(skill.systemPrompt).toContain('## Knowledge base')
  })

  it('answers both tools without calling the backend', async () => {
    const deps = makeDeps()
    const skill = createKnowledgeBaseSkill({ ...deps, listFiles: () => [] })
    const searched = await skill.executeTool!({
      id: '1',
      name: 'search_knowledge_base',
      input: { query: 'x' },
    })
    expect(searched).toMatchObject({ isError: true })
    expect(searched.output).toContain('empty')
    const read = await skill.executeTool!({
      id: '2',
      name: 'read_knowledge_file',
      input: { index: 0 },
    })
    expect(read.output).toContain('empty')
    expect(deps.searches).toEqual([])
    expect(deps.reads).toEqual([])
  })
})
