/**
 * The alias display pass (`WikiAliasDisplay`): an aliased link reads as its
 * alias, with the `[[target|` prefix and the closing `]]` hidden by
 * decorations. Pure view state — the document text keeps the whole literal,
 * and nothing here waits on the resolver: what to hide follows from the
 * document alone.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { collectWikiTargets, wikiAliasDisplayKey } from '../src/renderer/editor/wikilink'
import { parseDocText } from '../src/renderer/markdown/docText'

const editors: Editor[] = []
afterEach(() => {
  editors.splice(0).forEach((e) => e.destroy())
})

function open(source: string) {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: '',
  })
  editors.push(editor)
  // a trailing paragraph parks the caret off the links: one the selection
  // touches shows its whole literal
  const envelope = parseDocText(source + '\n\n.')
  editor
    .chain()
    .setMeta('addToHistory', false)
    .setContent(envelope.body, { contentType: 'markdown' })
    .run()
  editor.commands.setTextSelection(editor.state.doc.content.size)
  return { editor }
}

/** the [from, to) ranges the pass hides, in document order */
function hidden(editor: Editor): [number, number][] {
  const state = wikiAliasDisplayKey.getState(editor.state)
  expect(state, 'the alias display plugin must be registered').not.toBeNull()
  return state!.decorations.find().map((d): [number, number] => [d.from, d.to])
}

/** the text the hidden spans carry, as the DOM spells them */
function hiddenText(editor: Editor): (string | null)[] {
  return [...editor.view.dom.querySelectorAll('.wiki-alias-hidden')].map((el) => el.textContent)
}

describe('alias display', () => {
  it('hides the [[target| prefix and the closing ]] of an aliased link', () => {
    const { editor } = open('[[Note|Alias]]')
    // the paragraph's content starts at 1: `[[Note|` is hidden, `Alias` shows
    expect(hidden(editor)).toEqual([
      [1, 8],
      [13, 15],
    ])
    expect(hiddenText(editor)).toEqual(['[[Note|', ']]'])
    // the document text is what it always was: this pass is view-only
    expect(editor.state.doc.textContent).toBe('[[Note|Alias]].')
  })

  it('shows the whole literal of a link the selection touches', () => {
    const { editor } = open('[[A|one]] and [[B|two]]')
    const away = editor.state.doc.content.size
    // caret inside the alias: that link reveals, the other stays hidden
    editor.commands.setTextSelection(6)
    expect(hidden(editor)).toEqual([
      [15, 19],
      [22, 24],
    ])
    // right after the closing brackets still counts as touching
    editor.commands.setTextSelection(10)
    expect(hidden(editor)).toEqual([
      [15, 19],
      [22, 24],
    ])
    // a range over both reveals both
    editor.commands.setTextSelection({ from: 3, to: 20 })
    expect(hidden(editor)).toEqual([])
    editor.commands.setTextSelection(away)
    expect(hidden(editor)).toEqual([
      [1, 5],
      [8, 10],
      [15, 19],
      [22, 24],
    ])
  })

  it.each([
    ['a plain link', '[[Note]]'],
    ['a heading ref', '[[Note#Heading]]'],
  ])('leaves %s uncovered', (_label, src) => {
    const { editor } = open(src)
    expect(hidden(editor)).toEqual([])
    expect(editor.view.dom.querySelectorAll('.wiki-alias-hidden').length).toBe(0)
  })

  it('an empty alias is not an alias and gets nothing hidden', () => {
    const { editor } = open('[[Note|]]')
    // the link regex wants 1+ alias chars, so this never became a link
    expect(collectWikiTargets(editor.state.doc)).toEqual([])
    expect(hidden(editor)).toEqual([])
  })

  it('hides the syntax around the label of a block ref with an alias', () => {
    const { editor } = open('[[Note#^blockid|label]]')
    // `[[Note#^blockid|` is hidden and the closing `]]` with it; `label` shows
    expect(hidden(editor)).toEqual([
      [1, 17],
      [22, 24],
    ])
    expect(hiddenText(editor)).toEqual(['[[Note#^blockid|', ']]'])
  })

  it('an embed belongs to the embed feature and gets nothing hidden', () => {
    const { editor } = open('![[Note|200]]')
    expect(hidden(editor)).toEqual([])
    expect(editor.view.dom.querySelectorAll('.wiki-alias-hidden').length).toBe(0)
  })

  it('two aliased links in one paragraph each hide their own pair', () => {
    const { editor } = open('[[A|one]] and [[B|two]]')
    expect(hidden(editor)).toEqual([
      [1, 5],
      [8, 10],
      [15, 19],
      [22, 24],
    ])
  })

  it('follows an edit that shifts the link', () => {
    const { editor } = open('xx [[Note|Alias]]')
    expect(hidden(editor)).toEqual([
      [4, 11],
      [16, 18],
    ])
    // the pass recomputes whole on every doc change, so the offsets track
    editor.view.dispatch(editor.state.tr.insertText('y', 1))
    expect(hidden(editor)).toEqual([
      [5, 12],
      [17, 19],
    ])
  })

  it('leaves an alias an inline mark split alone rather than hiding half of it', () => {
    const { editor } = open('[[Note|bold]]')
    // bold the last two letters of the alias: the literal now spans several
    // text nodes and no single span carries the whole of it
    const bold = editor.schema.marks.bold!.create()
    editor.view.dispatch(editor.state.tr.addMark(10, 12, bold))
    expect(editor.state.doc.textContent).toBe('[[Note|bold]].')
    // the full literal stays visible; the pass never clips a segment
    expect(hidden(editor)).toEqual([])
    expect(editor.view.dom.querySelectorAll('.wiki-alias-hidden').length).toBe(0)
  })
})
