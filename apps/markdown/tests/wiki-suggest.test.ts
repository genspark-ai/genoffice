/**
 * The `[[` suggester: the vault listing feeds it, the filter narrows it, and
 * completing a row inserts REAL wiki syntax — a marked run for `[[`, an embed
 * node for `![[` — because live typing does not re-run the tokenizers.
 */
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'

import { buildExtensions } from '../src/renderer/editor/extensions'
import {
  filterWikiNotes,
  refreshWikiNotes,
  wikiQueryOpen,
  type WikiSuggestState,
} from '../src/renderer/editor/wikiSuggest'
import type { WikiNoteEntry } from '../src/shared/ipc'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (v: string) => v }

const NOTES: WikiNoteEntry[] = [
  { path: '/vault/Alpha.md', file: 'Alpha.md', name: 'Alpha', relPath: 'Alpha.md' },
  { path: '/vault/deep/Beta.png', file: 'Beta.png', name: 'Beta', relPath: 'deep/Beta.png' },
  { path: '/vault/Alpine.md', file: 'Alpine.md', name: 'Alpine', relPath: 'Alpine.md' },
]

beforeAll(() => {
  ;(window as any).markdownApi = {
    ...(window as any).markdownApi,
    listWikiNotes: async () => NOTES,
  }
})

const editors: Editor[] = []
afterEach(() => {
  editors.splice(0).forEach((e) => e.destroy())
  vi.restoreAllMocks()
})

describe('filterWikiNotes', () => {
  it('matches the stem or the file name, case-insensitively', () => {
    const hit = filterWikiNotes(NOTES, 'alp').map((n) => n.name)
    expect(hit).toEqual(['Alpha', 'Alpine'])
    expect(filterWikiNotes(NOTES, 'beta.png')[0]?.name).toBe('Beta')
    expect(filterWikiNotes(NOTES, 'zzz')).toEqual([])
  })

  it('narrows on the name part only; an alias or fragment does not filter', () => {
    expect(filterWikiNotes(NOTES, 'alp|My Title').map((n) => n.name)).toEqual(['Alpha', 'Alpine'])
    expect(filterWikiNotes(NOTES, 'alp#Some Heading').map((n) => n.name)).toEqual([
      'Alpha',
      'Alpine',
    ])
  })

  it('an empty query offers everything', () => {
    expect(filterWikiNotes(NOTES, '  ')).toHaveLength(3)
  })
})

/** an editor whose `[[` suggester reports to `states`, shared by both describes */
function mount(opts: { onClose?: () => void } = {}): {
  editor: Editor
  states: WikiSuggestState[]
} {
  const states: WikiSuggestState[] = []
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
      wikiSuggestController: {
        onOpen: (s) => states.push(s),
        onUpdate: (s) => states.push(s),
        onKeyDown: () => false,
        onClose: () => opts.onClose?.(),
      },
    }),
    content: '',
  })
  editors.push(editor)
  return { editor, states }
}

describe('completion', () => {
  it('a typed [[ opens the suggester with the vault notes', async () => {
    refreshWikiNotes()
    await new Promise((r) => setTimeout(r, 0))
    const { editor, states } = mount()
    editor.commands.insertContent('See [[')
    await new Promise((r) => setTimeout(r, 0))
    const state = states.at(-1)
    expect(state).not.toBeNull()
    expect(state!.items.map((n) => (n.kind === 'note' ? n.entry.name : ''))).toEqual([
      'Alpha',
      'Beta',
      'Alpine',
    ])
  })

  it('a query the user closed by hand is no longer a query', () => {
    expect(wikiQueryOpen('Alp')).toBe(true)
    expect(wikiQueryOpen('Alp|My Title')).toBe(true)
    expect(wikiQueryOpen('Alp#Heading')).toBe(true)
    expect(wikiQueryOpen('Alp]]')).toBe(false)
    expect(wikiQueryOpen('Alp]] tail')).toBe(false)
  })

  it('does not keep offering after the user closed the brackets by hand', async () => {
    refreshWikiNotes()
    await new Promise((r) => setTimeout(r, 0))
    let closes = 0
    const { editor } = mount({ onClose: () => (closes += 1) })
    editor.commands.insertContent('See [[')
    await new Promise((r) => setTimeout(r, 0))
    // type the rest of a link and close it by hand, then keep typing after it
    editor.commands.insertContent('Alpha]] tail')
    await new Promise((r) => setTimeout(r, 0))
    // the plugin must have exited: an open suggester would keep re-emitting
    // onUpdate with 0 items for every later keystroke
    expect(closes).toBeGreaterThan(0)
  })

  it('completing a note writes a real marked link, not plain text', async () => {
    refreshWikiNotes()
    await new Promise((r) => setTimeout(r, 0))
    const { editor, states } = mount()
    editor.commands.insertContent('See [[')
    await new Promise((r) => setTimeout(r, 0))
    states.at(-1)!.command({ kind: 'note', entry: NOTES[0]! })
    expect(editor.state.doc.textContent).toBe('See [[Alpha]]')
    let marked = false
    editor.state.doc.descendants((n) => {
      if (n.isText && n.marks.some((m) => m.type.name === 'wikiLink')) marked = true
      return true
    })
    expect(marked).toBe(true)
  })

  it('a typed alias survives the completion; the brackets do not double up', async () => {
    refreshWikiNotes()
    await new Promise((r) => setTimeout(r, 0))
    const { editor, states } = mount()
    editor.commands.insertContent('[[Alp|My Title')
    await new Promise((r) => setTimeout(r, 0))
    states.at(-1)!.command({ kind: 'note', entry: NOTES[2]! })
    expect(editor.state.doc.textContent).toBe('[[Alpine|My Title]]')
  })

  it('after ![[ the same list completes into an embed node', async () => {
    refreshWikiNotes()
    await new Promise((r) => setTimeout(r, 0))
    const { editor, states } = mount()
    editor.commands.insertContent('Look ![[')
    await new Promise((r) => setTimeout(r, 0))
    states.at(-1)!.command({ kind: 'note', entry: NOTES[1]! })
    // the typed bang belongs to the embed, and completionFor adds its own — the
    // one in the text must go, or every completed embed saves as `!![[Beta]]`
    expect(editor.getMarkdown()).toBe('Look ![[Beta]]')
    expect(editor.state.doc.textContent).toBe('Look ')
    let embeds = 0
    editor.state.doc.descendants((n) => {
      if (n.type.name === 'wikiEmbed') {
        embeds += 1
        expect(n.attrs.raw).toBe('![[Beta]]')
      }
      return true
    })
    expect(embeds).toBe(1)
  })
})

describe('create new note', () => {
  function mountWithCreate(): { editor: Editor; states: WikiSuggestState[] } {
    ;(window as any).markdownApi.createWikiNote = async () => ({
      ok: true as const,
      path: '/vault/My_Zzz.md',
      target: 'My_Zzz',
      fragment: '',
      created: true,
    })
    return mount()
  }

  it('a query nothing matches offers its name for creation', async () => {
    refreshWikiNotes()
    await new Promise((r) => setTimeout(r, 0))
    const { editor, states } = mountWithCreate()
    editor.commands.insertContent('See [[')
    await new Promise((r) => setTimeout(r, 0))
    // the whole vault is on offer with nothing typed; creation needs a name
    expect(states.at(-1)!.createName).toBe('')
    editor.commands.insertContent('Zzz')
    await new Promise((r) => setTimeout(r, 0))
    expect(states.at(-1)!.items).toEqual([])
    expect(states.at(-1)!.createName).toBe('Zzz')
  })

  it('creating completes the link with the created file, not the typed text', async () => {
    refreshWikiNotes()
    await new Promise((r) => setTimeout(r, 0))
    const { editor, states } = mountWithCreate()
    editor.commands.insertContent('See [[Zzz')
    await new Promise((r) => setTimeout(r, 0))
    states.at(-1)!.create()
    await new Promise((r) => setTimeout(r, 0))
    // the stubbed note is `My_Zzz.md`: the sanitized stem is the spelling that
    // resolves, so the completion carries it rather than the typed `Zzz`
    expect(editor.state.doc.textContent).toBe('See [[My_Zzz]]')
    let marked = false
    editor.state.doc.descendants((n) => {
      if (n.isText && n.marks.some((m) => m.type.name === 'wikiLink')) marked = true
      return true
    })
    expect(marked).toBe(true)
  })

  it('a typed alias or fragment survives the creation', async () => {
    refreshWikiNotes()
    await new Promise((r) => setTimeout(r, 0))
    const { editor, states } = mountWithCreate()
    editor.commands.insertContent('[[Zzz|My Title')
    await new Promise((r) => setTimeout(r, 0))
    expect(states.at(-1)!.createName).toBe('Zzz')
    states.at(-1)!.create()
    await new Promise((r) => setTimeout(r, 0))
    expect(editor.state.doc.textContent).toBe('[[My_Zzz|My Title]]')
  })

  it('creating from an embed completes into an embed node', async () => {
    refreshWikiNotes()
    await new Promise((r) => setTimeout(r, 0))
    const { editor, states } = mountWithCreate()
    editor.commands.insertContent('Look ![[Zzz')
    await new Promise((r) => setTimeout(r, 0))
    states.at(-1)!.create()
    await new Promise((r) => setTimeout(r, 0))
    expect(editor.getMarkdown()).toBe('Look ![[My_Zzz]]')
  })

  it('a refused creation leaves the text alone', async () => {
    refreshWikiNotes()
    await new Promise((r) => setTimeout(r, 0))
    const { editor, states } = mount()
    ;(window as any).markdownApi.createWikiNote = async () => ({ ok: false as const })
    editor.commands.insertContent('See [[../evil')
    await new Promise((r) => setTimeout(r, 0))
    expect(states.at(-1)!.createName).toBe('../evil')
    states.at(-1)!.create()
    await new Promise((r) => setTimeout(r, 0))
    expect(editor.state.doc.textContent).toBe('See [[../evil')
  })
})
