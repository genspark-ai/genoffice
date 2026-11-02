/**
 * The fragment suggester's main-side half: what a note offers behind `[[Note#`
 * is read from the file — headings, fenced code, the frontmatter block and the
 * lone-marker layout the resolver and the id writer already
 * agree on.
 */
import { describe, expect, it } from 'vitest'
import { parseWikiHeads } from '../src/main/wiki-heads'

describe('parseWikiHeads', () => {
  it('a full note: frontmatter skipped, headings, code skipped, trailing and lone markers', () => {
    const raw = [
      '---',
      'tags: [x]',
      '# not a heading',
      '---',
      '',
      '# Setup',
      '',
      'Body paragraph. ^body-1',
      '',
      '```js',
      '// code ^not-an-id',
      'const x = 1',
      '```',
      '',
      '| A | B |',
      '| - | - |',
      '| 1 | 2 |',
      '',
      '^table-1',
      '',
      '## Second ^sec-1',
      '',
      'Tail.',
    ].join('\n')
    expect(parseWikiHeads(raw)).toEqual([
      { kind: 'heading', text: 'Setup', level: 1 },
      { kind: 'block', id: 'body-1', text: 'Body paragraph.' },
      { kind: 'block', id: 'table-1', text: '| 1 | 2 |' },
      { kind: 'heading', text: 'Second', level: 2 },
      { kind: 'block', id: 'sec-1', text: '## Second' },
    ])
  })

  it('a lone marker after a heading borrows the heading line as its context', () => {
    const raw = '# Rules\n\n^rules\n\nBody.'
    expect(parseWikiHeads(raw)).toEqual([
      { kind: 'heading', text: 'Rules', level: 1 },
      { kind: 'block', id: 'rules', text: '# Rules' },
    ])
  })

  it('an empty or absent note parses to nothing', () => {
    expect(parseWikiHeads('')).toEqual([])
    expect(parseWikiHeads('plain text only\n')).toEqual([])
  })

  it('closes a fence only on the same character, so ~~~ does not close ```', () => {
    const raw = '```\n~~~\n# inside\n```\n\n# Outside'
    expect(parseWikiHeads(raw)).toEqual([{ kind: 'heading', text: 'Outside', level: 1 }])
  })

  it('a shorter fence or one with an info string does not close a longer fence', () => {
    const raw = '````md\n```js\n# inside\n```\n# still inside\n````\n\n# Outside'
    expect(parseWikiHeads(raw)).toEqual([{ kind: 'heading', text: 'Outside', level: 1 }])
    expect(parseWikiHeads('```\n```js\n# inside\n```\n# Out')).toEqual([
      { kind: 'heading', text: 'Out', level: 1 },
    ])
  })

  it('a closing-bracket hash run is trimmed off the heading text', () => {
    expect(parseWikiHeads('### Title ###\n')).toEqual([
      { kind: 'heading', text: 'Title', level: 3 },
    ])
  })
})
