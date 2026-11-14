import { describe, expect, it } from 'vitest'
import { setAuthorName } from '../src/renderer/author-name'
import { AI_AUTHOR, aiRevisionAuthor } from '../src/renderer/ai/revision-author'

describe('aiRevisionAuthor', () => {
  it('attributes AI edits to the AI author, not the Settings user name', () => {
    // the in-app assistant, its comment replies and the in-app MCP server's
    // content tools all route through this one author (#2654)
    setAuthorName('Pat')
    expect(aiRevisionAuthor()).toBe('AI')
    expect(AI_AUTHOR).toBe('AI')
  })
})
