import { describe, it, expect } from 'vitest'
import {
  parseFrontmatter,
  serializeFrontmatter,
  markdownToProseMirror,
  prosemirrorToMarkdown,
} from '../src/renderer/markdown-io'
import { markdownExtensions } from '../src/renderer/editor/extensions'

describe('parseFrontmatter', () => {
  it('extracts YAML frontmatter from markdown', () => {
    const input = '---\ntitle: Hello\ntags: [a, b]\n---\n\n# World\n\nBody text.'
    const result = parseFrontmatter(input)
    expect(result.frontmatter).toEqual({ title: 'Hello', tags: ['a', 'b'] })
    expect(result.body).toContain('# World')
    expect(result.body).toContain('Body text.')
  })

  it('returns null frontmatter when none present', () => {
    const input = '# Hello\n\nBody text.'
    const result = parseFrontmatter(input)
    expect(result.frontmatter).toBeNull()
    expect(result.body).toBe('# Hello\n\nBody text.')
  })

  it('handles empty frontmatter', () => {
    const input = '---\n---\n\nBody'
    const result = parseFrontmatter(input)
    expect(result.frontmatter).toBeNull()
    expect(result.body).toContain('Body')
  })
})

describe('serializeFrontmatter', () => {
  it('serializes object to YAML frontmatter block', () => {
    const data = { title: 'Hello', tags: ['a', 'b'] }
    const result = serializeFrontmatter(data)
    expect(result).toContain('---')
    expect(result).toContain('title: Hello')
    expect(result).toContain('tags:')
  })

  it('returns empty string for null', () => {
    expect(serializeFrontmatter(null)).toBe('')
  })

  it('returns empty string for empty object', () => {
    expect(serializeFrontmatter({})).toBe('')
  })
})

describe('prosemirrorToMarkdown', () => {
  it('serializes headings', () => {
    const json = {
      type: 'doc',
      content: [
        { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Title' }] },
      ],
    }
    const md = prosemirrorToMarkdown(json)
    expect(md).toContain('# Title')
  })

  it('serializes bold and italic', () => {
    const json = {
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          content: [
            { type: 'text', text: 'bold', marks: [{ type: 'bold' }] },
            { type: 'text', text: ' and ' },
            { type: 'text', text: 'italic', marks: [{ type: 'italic' }] },
          ],
        },
      ],
    }
    const md = prosemirrorToMarkdown(json)
    expect(md).toContain('**bold**')
    expect(md).toContain('*italic*')
  })

  it('serializes bullet lists', () => {
    const json = {
      type: 'doc',
      content: [
        {
          type: 'bulletList',
          content: [
            {
              type: 'listItem',
              content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Item 1' }] }],
            },
            {
              type: 'listItem',
              content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Item 2' }] }],
            },
          ],
        },
      ],
    }
    const md = prosemirrorToMarkdown(json)
    expect(md).toContain('- Item 1')
    expect(md).toContain('- Item 2')
  })

  it('serializes code blocks', () => {
    const json = {
      type: 'doc',
      content: [
        {
          type: 'codeBlock',
          attrs: { language: 'typescript' },
          content: [{ type: 'text', text: 'const x = 1' }],
        },
      ],
    }
    const md = prosemirrorToMarkdown(json)
    expect(md).toContain('```typescript')
    expect(md).toContain('const x = 1')
    expect(md).toContain('```')
  })

  it('serializes blockquotes', () => {
    const json = {
      type: 'doc',
      content: [
        {
          type: 'blockquote',
          content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Quote me' }] }],
        },
      ],
    }
    const md = prosemirrorToMarkdown(json)
    expect(md).toContain('>')
    expect(md).toContain('Quote me')
  })

  it('serializes horizontal rules', () => {
    const json = {
      type: 'doc',
      content: [{ type: 'horizontalRule' }],
    }
    const md = prosemirrorToMarkdown(json)
    expect(md).toContain('---')
  })

  it('serializes task lists', () => {
    const json = {
      type: 'doc',
      content: [
        {
          type: 'taskList',
          content: [
            {
              type: 'taskItem',
              attrs: { checked: true },
              content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Done' }] }],
            },
            {
              type: 'taskItem',
              attrs: { checked: false },
              content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Todo' }] }],
            },
          ],
        },
      ],
    }
    const md = prosemirrorToMarkdown(json)
    expect(md).toContain('- [x] Done')
    expect(md).toContain('- [ ] Todo')
  })
})

describe('markdownToProseMirror', () => {
  it('parses markdown to ProseMirror JSON with headings', () => {
    const input = '# Hello\n\nWorld'
    // @ts-expect-error — TipTap duplicate module type mismatch in workspace monorepo
    const json = markdownToProseMirror(input, markdownExtensions)
    expect(json.type).toBe('doc')
    expect(json.content).toBeDefined()
    // Should contain a heading node
    const heading = json.content?.find((n) => n.type === 'heading')
    expect(heading).toBeDefined()
    expect(heading?.attrs?.level).toBe(1)
  })

  it('parses markdown with frontmatter', () => {
    const input = '---\ntitle: Test\n---\n\n# Content'
    // @ts-expect-error — TipTap duplicate module type mismatch in workspace monorepo
    const json = markdownToProseMirror(input, markdownExtensions)
    const fmNode = json.content?.find((n) => n.type === 'frontmatter')
    expect(fmNode).toBeDefined()
    expect(fmNode?.attrs?.data).toEqual({ title: 'Test' })
  })
})
