/**
 * The hand-off a wiki link's `#fragment` makes: the tab that follows it and the
 * tab that shows the note are not always the same one, so the fragment waits
 * in main, keyed by path, until a renderer loads that path.
 */
import { describe, expect, it } from 'vitest'
import { rememberWikiFragment, takeWikiFragment } from '../src/main/wiki-fragment'

const PATH = '/notes/Index.md'

describe('wiki fragments', () => {
  it('hands the fragment to the first renderer that asks for that path', () => {
    rememberWikiFragment(PATH, 'Some Heading')
    expect(takeWikiFragment(PATH)).toBe('Some Heading')
    expect(takeWikiFragment(PATH)).toBe('')
  })

  it('keeps fragments apart per path', () => {
    rememberWikiFragment('/notes/A.md', 'One')
    rememberWikiFragment('/notes/B.md', 'Two')
    expect(takeWikiFragment('/notes/A.md')).toBe('One')
    expect(takeWikiFragment('/notes/B.md')).toBe('Two')
  })

  it('lets a later link to the same note replace the earlier target', () => {
    rememberWikiFragment(PATH, 'First')
    rememberWikiFragment(PATH, 'Second')
    expect(takeWikiFragment(PATH)).toBe('Second')
  })

  it('a link with no fragment cancels an earlier one', () => {
    rememberWikiFragment(PATH, 'First')
    rememberWikiFragment(PATH, '')
    expect(takeWikiFragment(PATH)).toBe('')
  })

  it('forgets the oldest fragment rather than growing without bound', () => {
    for (let i = 0; i < 40; i++) rememberWikiFragment(`/notes/${i}.md`, `H${i}`)
    // the oldest waited on is gone; the newest is still there
    expect(takeWikiFragment('/notes/0.md')).toBe('')
    expect(takeWikiFragment('/notes/39.md')).toBe('H39')
  })

  it('an untitled document has no path to wait on', () => {
    rememberWikiFragment('', 'Heading')
    expect(takeWikiFragment('')).toBe('')
  })
})
