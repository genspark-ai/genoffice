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
  searches: Array<[string, number, unknown]>
  reads: Array<[string, number, unknown]>
} {
  const searches: Array<[string, number, unknown]> = []
  const reads: Array<[string, number, unknown]> = []
  return {
    searches,
    reads,
    listFiles: () => CORPUS,
    search: async (query, limit, opts) => {
      searches.push([query, limit, opts])
      return overrides.search
        ? overrides.search(query, limit, opts)
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
    read: async (path, offset, opts) => {
      reads.push([path, offset, opts])
      return overrides.read
        ? overrides.read(path, offset, opts)
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
    expect(deps.searches).toEqual([['\u8bd5\u7528\u671f', 5, undefined]])
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
    expect(deps.reads).toEqual([
      ['/Users/me/\u5408\u540c/\u79df\u8d41\u5408\u540c.docx', 40, undefined],
    ])
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

  it('read accepts a search-result path (the indexed corpus reaches past the listing)', async () => {
    const deps = makeDeps()
    const skill = createKnowledgeBaseSkill(deps)
    const result = await skill.executeTool!({
      id: 't6',
      name: 'read_knowledge_file',
      input: { path: '/Users/me/work/notes.docx', offset: 0 },
    })
    expect(result.isError).toBeFalsy()
    expect(deps.reads).toEqual([['/Users/me/work/notes.docx', 0, undefined]])
    expect(result.output).toContain('characters 0..2 of 100')
  })

  it('verifyResponse whitelists what this session searched out, still rejecting the invented', async () => {
    const deps = makeDeps({
      search: async () => [
        {
          path: '/Users/me/work/notes.docx',
          name: 'notes.docx',
          ext: 'docx',
          folder: 'work',
          snippet: 'hit',
        },
      ],
    })
    const skill = createKnowledgeBaseSkill(deps)
    await skill.executeTool!({
      id: 't7',
      name: 'search_knowledge_base',
      input: { query: 'notes' },
    })
    // an indexed-but-unlisted path found through search cites cleanly
    const good = 'see [notes.docx](filenav:///Users/me/work/notes.docx)'
    expect(skill.verifyResponse!(good, [])).toBeNull()
    // a path nobody showed the model is still refused
    const bad = 'cited [fake.docx](filenav:///tmp/fake.docx)'
    expect(skill.verifyResponse!(bad, [])).toContain('not knowledge-base paths')
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

describe('createKnowledgeBaseSkill with a round scope', () => {
  const SCOPED = [CORPUS[0]!.path, '/Users/me/work/notes.docx']
  const makeScoped = (
    overrides: Partial<KnowledgeBaseDeps> = {},
    scope: readonly string[] = SCOPED,
  ) => {
    const calls = {
      searches: [] as Array<[string, number, { scopePaths?: readonly string[] } | undefined]>,
      reads: [] as Array<[string, number, { scopePaths?: readonly string[] } | undefined]>,
    }
    const deps: KnowledgeBaseDeps = {
      listFiles: overrides.listFiles ?? (() => CORPUS),
      search: async (query, limit, opts) => {
        calls.searches.push([query, limit, opts])
        return (
          overrides.search?.(query, limit, opts) ?? [
            {
              path: CORPUS[0]!.path,
              name: CORPUS[0]!.name,
              ext: 'docx',
              folder: '\u5408\u540c',
              snippet: '\u8bd5\u7528\u671f\u8003\u6838\u529e\u6cd5',
            },
          ]
        )
      },
      read: async (path, offset, opts) => {
        calls.reads.push([path, offset, opts])
        return (
          overrides.read?.(path, offset, opts) ?? { text: '\u6b63\u6587', totalChars: 100, offset }
        )
      },
      scopePaths: () => scope,
      ...(overrides.alsoCitable ? { alsoCitable: overrides.alsoCitable } : {}),
    }
    return { deps, calls }
  }

  it('the system prompt replaces the whole-index language with the selected list', () => {
    const skill = createKnowledgeBaseSkill(makeScoped().deps)
    expect(skill.systemPrompt).toContain('## Knowledge base')
    expect(skill.systemPrompt).toContain('Corpus scope for this conversation')
    expect(skill.systemPrompt).toContain('selected 2 files')
    expect(skill.systemPrompt).toContain(CORPUS[0]!.path)
    expect(skill.systemPrompt).toContain('/Users/me/work/notes.docx')
    // without a scope the prompt is exactly the unscoped one
    const unscoped = createKnowledgeBaseSkill(makeDeps())
    expect(unscoped.systemPrompt).not.toContain('Corpus scope')
  })

  it('the context listing only shows the selected files', () => {
    const skill = createKnowledgeBaseSkill(makeScoped().deps)
    const context = skill.buildContext!()
    // the scoped file keeps its listing; the out-of-scope one drops out
    expect(context).toContain(CORPUS[0]!.name)
    expect(context).not.toContain('employee handbook.pdf')
  })

  it('an empty scope behaves exactly like no scope at all', () => {
    const { deps } = makeScoped({}, [])
    const skill = createKnowledgeBaseSkill(deps)
    expect(skill.systemPrompt).not.toContain('Corpus scope')
    const context = skill.buildContext!()
    expect(context).toContain('employee handbook.pdf')
  })

  it('a scope keeps the tools mounted even when the listing window is empty', () => {
    const { deps } = makeScoped({ listFiles: () => [] })
    const skill = createKnowledgeBaseSkill(deps)
    expect(skill.tools.map((t) => t.name)).toEqual(['read_knowledge_file', 'search_knowledge_base'])
    // and the tools answer rather than refusing with "empty"
    expect(skill.systemPrompt).toContain('## Knowledge base')
  })

  it('search and read carry the scope on their calls, and drop it when none is attached', async () => {
    const { deps, calls } = makeScoped()
    const skill = createKnowledgeBaseSkill(deps)
    await skill.executeTool!({ id: 's1', name: 'search_knowledge_base', input: { query: '合同' } })
    await skill.executeTool!({
      id: 'r1',
      name: 'read_knowledge_file',
      input: { path: '/Users/me/work/notes.docx', offset: 0 },
    })
    await skill.executeTool!({ id: 'r2', name: 'read_knowledge_file', input: { index: 0 } })
    expect(calls.searches[0]![2]).toEqual({ scopePaths: SCOPED })
    expect(calls.reads[0]![2]).toEqual({ scopePaths: SCOPED })
    expect(calls.reads[1]![2]).toEqual({ scopePaths: SCOPED })

    const plain = makeDeps()
    const unscoped = createKnowledgeBaseSkill(plain)
    await unscoped.executeTool!({
      id: 's2',
      name: 'search_knowledge_base',
      input: { query: '合同' },
    })
    await unscoped.executeTool!({ id: 'r3', name: 'read_knowledge_file', input: { index: 0 } })
    expect(plain.searches[0]![2]).toBeUndefined()
  })

  it('verifyResponse still accepts an in-scope citation and refuses out-of-scope ones', () => {
    const { deps } = makeScoped()
    const skill = createKnowledgeBaseSkill(deps)
    // in scope (listed window ∩ scope)
    const inScope =
      '\u89c1 [\u79df\u8d41\u5408\u540c.docx](filenav:///Users/me/%E5%90%88%E5%90%8C/%E7%A7%9F%E8%B5%81%E5%90%88%E5%90%8C.docx)'
    expect(skill.verifyResponse!(inScope, [])).toBeNull()
    // listed in the corpus window but outside the selection: still refused
    const outOfScope =
      'cited [employee handbook.pdf](filenav:///Users/me/%E5%B7%A5%E4%BD%9C/employee%20handbook.pdf)'
    expect(skill.verifyResponse!(outOfScope, [])).toContain('not knowledge-base paths')
    // invented path, as ever
    const invented = 'cited [fake.docx](filenav:///tmp/fake.docx)'
    expect(skill.verifyResponse!(invented, [])).toContain('not knowledge-base paths')
  })

  it('verifyResponse refuses a search-result path from before the scope was attached', async () => {
    // a session that searched unscoped first, then had a selection attached
    const { deps } = makeScoped({
      search: async () => [
        {
          path: '/Users/me/work/other.docx',
          name: 'other.docx',
          ext: 'docx',
          folder: 'work',
          snippet: 'hit',
        },
      ],
    })
    const skill = createKnowledgeBaseSkill(deps)
    await skill.executeTool!({ id: 's1', name: 'search_knowledge_base', input: { query: 'other' } })
    const stale = 'see [other.docx](filenav:///Users/me/work/other.docx)'
    expect(skill.verifyResponse!(stale, [])).toContain('not knowledge-base paths')
  })

  it('verifyResponse keeps this session’s artifacts citable under a scope', () => {
    const { deps } = makeScoped({ alsoCitable: () => ['/save/built-report.docx'] })
    const skill = createKnowledgeBaseSkill(deps)
    const created = 'built [report](filenav:///save/built-report.docx)'
    expect(skill.verifyResponse!(created, [])).toBeNull()
  })
})
