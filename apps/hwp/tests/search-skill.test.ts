import { describe, expect, it } from 'vitest'
import { createSearchSkill } from '../src/renderer/ai/search-skill'

describe('createSearchSkill', () => {
  it('asks for citations in the reply, not an edit of the document', () => {
    const skill = createSearchSkill(async () => ({ results: [], method: 'ok' }))
    expect(skill.systemPrompt).toMatch(/in your reply/i)
    expect(skill.systemPrompt).toMatch(/do not claim you wrote them into the document/i)
    expect(skill.systemPrompt).not.toMatch(/when writing search results into the document/i)
  })

  it('rejects an empty query', async () => {
    const skill = createSearchSkill(async () => {
      throw new Error('should not search')
    })
    const result = await skill.executeTool({ id: '1', name: 'web_search', input: { query: '  ' } })
    expect(result.isError).toBe(true)
    expect(result.output).toContain('query must not be empty')
  })

  it('surfaces a service error without treating it as empty results', async () => {
    const skill = createSearchSkill(async () => ({
      results: [],
      method: 'error',
      error: 'rate limited',
    }))
    const result = await skill.executeTool({
      id: '1',
      name: 'web_search',
      input: { query: '한글' },
    })
    expect(result.isError).toBe(true)
    expect(result.output).toContain('rate limited')
    expect(result.output).not.toContain('(no results)')
  })

  it('formats titles, links, and snippets', async () => {
    const skill = createSearchSkill(async (query, maxResults) => {
      expect(query).toBe('hwpx spec')
      expect(maxResults).toBe(3)
      return {
        answer: 'HWPX is an XML package.',
        results: [{ title: 'Spec', url: 'https://example.test', snippet: 'Open format' }],
        method: 'ok',
      }
    })
    const result = await skill.executeTool({
      id: '1',
      name: 'web_search',
      input: { query: 'hwpx spec', maxResults: 3 },
    })
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('HWPX is an XML package.')
    expect(result.output).toContain('https://example.test')
    expect(result.output).toContain('Open format')
  })
})
