/**
 * The fragment list behind `[[Note#`: the target note's headings, then its
 * `^block` ids once the fragment starts with a caret; `[[#` reads this
 * document instead of going through main. The main-side parser is what a
 * plain file yields, so it is tested against raw text — the same shapes the
 * resolver and #2257/#2268 read.
 */
import { afterEach, beforeAll, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'

import { buildExtensions } from '../src/renderer/editor/extensions'
import {
  docFragmentItems,
  filterWikiHeads,
  invalidateWikiHeads,
  type WikiSuggestState,
} from '../src/renderer/editor/wikiSuggest'
import { parseWikiHeads } from '../src/main/wiki-heads'

const editors: Editor[] = []
afterEach(() => {
  editors.splice(0).forEach((e) => e.destroy())
})

function openDoc(
  source: string,
  path = '/vault/Note.md',
): { editor: Editor; states: WikiSuggestState[] } {
  const states: WikiSuggestState[] = []
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
      wikiSuggestController: {
        onOpen: (s) => states.push(s),
        onUpdate: (s) => states.push(s),
        onKeyDown: () => false,
        onClose: () => {},
      },
      wikiLink: { getDocumentPath: () => path },
    }),
    content: '',
  })
  editors.push(editor)
  editor.commands.setContent(source, { contentType: 'markdown' })
  return { editor, states }
}

const tick = () => new Promise((r) => setTimeout(r, 0))

describe('parseWikiHeads', () => {
  it('lists headings with their levels; frontmatter and fences are skipped', () => {
    const raw = [
      '---',
      'title: x',
      '# not a heading',
      '---',
      '',
      '# Top',
      '## Sub head',
      '```js',
      '# also not a heading',
      '```',
      '',
    ].join('\n')
    expect(parseWikiHeads(raw)).toEqual([
      { kind: 'heading', text: 'Top', level: 1 },
      { kind: 'heading', text: 'Sub head', level: 2 },
    ])
  })

  it('a trailing marker is a block with its line as context; a lone one borrows the previous line', () => {
    const raw = 'Para one. ^abc123\n\n| A |\n| - |\n| 1 |\n\n^tbl-1\n\nplain\n'
    expect(parseWikiHeads(raw)).toEqual([
      { kind: 'block', id: 'abc123', text: 'Para one.' },
      { kind: 'block', id: 'tbl-1', text: '| 1 |' },
    ])
  })

  it('a heading carrying its own marker is listed twice: by text and by id', () => {
    expect(parseWikiHeads('# Setup ^setup-1\n')).toEqual([
      { kind: 'heading', text: 'Setup', level: 1 },
      { kind: 'block', id: 'setup-1', text: '# Setup' },
    ])
  })
})

describe('this document as the fragment source', () => {
  const editorOf = () => {
    const { editor } = openDoc('# Alpha\n\nBody line. ^b1\n\n## Second')
    return editor
  }

  it('[[# offers the headings, longest match by what was typed', () => {
    const editor = editorOf()
    expect(docFragmentItems(editor.state.doc, '')).toEqual([
      { kind: 'heading', text: 'Alpha', level: 1 },
      { kind: 'heading', text: 'Second', level: 2 },
    ])
    expect(docFragmentItems(editor.state.doc, 'sec')).toEqual([
      { kind: 'heading', text: 'Second', level: 2 },
    ])
  })

  it('a caret fragment lists block ids and narrows by id or context', () => {
    const editor = editorOf()
    expect(docFragmentItems(editor.state.doc, '^')).toEqual([
      { kind: 'block', id: 'b1', text: 'Body line.' },
    ])
    expect(docFragmentItems(editor.state.doc, '^b1')).toHaveLength(1)
    expect(docFragmentItems(editor.state.doc, '^zzz')).toEqual([])
  })

  it('a heading carrying its own marker completes by text without the ^id', () => {
    const { editor } = openDoc('# Title ^hdr1\n\nplain')
    expect(docFragmentItems(editor.state.doc, '')).toEqual([
      { kind: 'heading', text: 'Title', level: 1 },
    ])
    expect(docFragmentItems(editor.state.doc, '^')).toEqual([
      { kind: 'block', id: 'hdr1', text: 'Title' },
    ])
  })
})

describe('the suggester in fragment mode', () => {
  beforeAll(() => {
    ;(window as unknown as { CSS?: unknown }).CSS ??= { escape: (v: string) => v }
    ;(window as unknown as { markdownApi?: Record<string, unknown> }).markdownApi ??= {}
  })
  beforeEach(() => {
    invalidateWikiHeads()
  })

  function mount(path = '/vault/Note.md'): { editor: Editor; states: WikiSuggestState[] } {
    const states: WikiSuggestState[] = []
    const editor = new Editor({
      extensions: buildExtensions({
        slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
        slashItems: () => [],
        wikiSuggestController: {
          onOpen: (s) => states.push(s),
          onUpdate: (s) => states.push(s),
          onKeyDown: () => false,
          onClose: () => {},
        },
        wikiLink: { getDocumentPath: () => path },
      }),
      content: '',
    })
    editors.push(editor)
    return { editor, states }
  }

  it('[[# lists this note’s headings and completing writes a real marked link', async () => {
    const { editor, states } = mount()
    editor.commands.setContent('# Alpha\n\nBody.', { contentType: 'markdown' })
    editor.commands.focus('end')
    editor.commands.insertContent(' See [[')
    await tick()
    editor.commands.insertContent('#')
    await tick()
    const state = states.at(-1)!
    expect(state.items).toEqual([{ kind: 'heading', text: 'Alpha', level: 1 }])
    // a fragment names something that exists; creation is the notes list's row
    expect(state.createName).toBe('')
    state.command(state.items[0]!)
    expect(editor.getMarkdown()).toContain('[[#Alpha]]')
    expect(editor.state.doc.textContent.endsWith('See [[#Alpha]]')).toBe(true)
    let marked = false
    editor.state.doc.descendants((n) => {
      if (n.isText && n.marks.some((m) => m.type.name === 'wikiLink')) marked = true
      return true
    })
    expect(marked).toBe(true)
  })

  it('[[Note# lists the target’s heads through main, and ^ narrows to its blocks', async () => {
    const asked: string[] = []
    ;(window as unknown as { markdownApi: Record<string, unknown> }).markdownApi.wikiHeads =
      async (request: { target: string; fromPath: string }) => {
        asked.push(`${request.fromPath}\u0000${request.target}`)
        return [
          { kind: 'heading', text: 'Alpha', level: 1 },
          { kind: 'block', id: 'abc123', text: 'Body line.' },
        ]
      }
    const { editor, states } = mount('/vault/Home.md')
    editor.commands.insertContent('See [[')
    await tick()
    editor.commands.insertContent('Alpha#')
    await tick()
    await tick()
    expect(asked).toEqual(['/vault/Home.md\u0000Alpha'])
    expect(states.at(-1)!.items).toEqual([{ kind: 'heading', text: 'Alpha', level: 1 }])
    editor.commands.insertContent('^')
    await tick()
    await tick()
    expect(states.at(-1)!.items).toEqual([{ kind: 'block', id: 'abc123', text: 'Body line.' }])
    states.at(-1)!.command(states.at(-1)!.items[0]!)
    expect(editor.state.doc.textContent).toBe('See [[Alpha#^abc123]]')
    // the cache answers the same question without asking main again
    editor.commands.insertContent('2')
    await tick()
    await tick()
    expect(asked).toHaveLength(1)
  })

  it('a typed alias narrows the fragment on the head and rides along on completion', async () => {
    ;(window as unknown as { markdownApi: Record<string, unknown> }).markdownApi.wikiHeads =
      async () => [
        { kind: 'heading', text: 'Alpha', level: 1 },
        { kind: 'heading', text: 'Beta', level: 2 },
      ]
    const { editor, states } = mount('/vault/Home.md')
    editor.commands.insertContent('See [[')
    await tick()
    editor.commands.insertContent('Note#Alpha|my alias')
    await tick()
    await tick()
    const state = states.at(-1)!
    // the filter saw `Alpha`, not `Alpha|my alias`
    expect(state.items).toEqual([{ kind: 'heading', text: 'Alpha', level: 1 }])
    state.command(state.items[0]!)
    expect(editor.state.doc.textContent).toBe('See [[Note#Alpha|my alias]]')
  })

  it('an embed completion keeps the bang and adds its own brackets only once', async () => {
    ;(window as unknown as { markdownApi: Record<string, unknown> }).markdownApi.wikiHeads =
      async () => [{ kind: 'heading', text: 'Alpha', level: 1 }]
    const { editor, states } = mount()
    editor.commands.insertContent('Look ![')
    await tick()
    editor.commands.insertContent('[')
    await tick()
    editor.commands.insertContent('Note#')
    await tick()
    await tick()
    const state = states.at(-1)!
    state.command(state.items[0]!)
    expect(editor.getMarkdown()).toBe('Look ![[Note#Alpha]]')
  })
})

describe('filterWikiHeads', () => {
  const entries = [
    { kind: 'heading', text: 'Alpha', level: 1 },
    { kind: 'block', id: 'b1', text: 'Body' },
  ] as const

  it('a caret fragment offers only blocks, plain fragments only headings', () => {
    expect(filterWikiHeads([...entries], '')).toEqual([
      { kind: 'heading', text: 'Alpha', level: 1 },
    ])
    expect(filterWikiHeads([...entries], '^')).toEqual([{ kind: 'block', id: 'b1', text: 'Body' }])
  })

  it('narrows by substring of the text or the id, case-insensitively', () => {
    expect(filterWikiHeads([...entries], 'alph')).toEqual([
      { kind: 'heading', text: 'Alpha', level: 1 },
    ])
    expect(filterWikiHeads([...entries], '^BOD')).toEqual([
      { kind: 'block', id: 'b1', text: 'Body' },
    ])
    expect(filterWikiHeads([...entries], '^zzz')).toEqual([])
  })
})
