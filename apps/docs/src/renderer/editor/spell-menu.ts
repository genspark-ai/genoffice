import { Extension, type Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { spellWords, wordsOfBlock, type SpellWord } from './spell-walk'

type IgnoreOnceMeta = { add: { from: number; to: number } } | { clear: true }

export const spellIgnoreOnceKey = new PluginKey<DecorationSet>('spellIgnoreOnce')

/**
 * Words set aside by Ignore Once: this occurrence only, for the session. The
 * decoration turns the browser's checking off on the word, and follows it
 * through edits.
 */
export const SpellIgnoreOnceExtension = Extension.create({
  name: 'spellIgnoreOnce',
  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key: spellIgnoreOnceKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            const meta = tr.getMeta(spellIgnoreOnceKey) as IgnoreOnceMeta | undefined
            if (meta && 'clear' in meta) return DecorationSet.empty
            const mapped = set.map(tr.mapping, tr.doc)
            if (!meta) return mapped
            const { from, to } = meta.add
            return mapped.add(tr.doc, [Decoration.inline(from, to, { spellcheck: 'false' })])
          },
        },
        props: {
          decorations: (state) => spellIgnoreOnceKey.getState(state),
        },
      }),
    ]
  },
})

export function ignoredOnce(state: EditorState, from: number, to: number): boolean {
  const set = spellIgnoreOnceKey.getState(state)
  return !!set && set.find(from, to).some((d) => d.from <= from && d.to >= to)
}

export function ignoreSpellingOnce(editor: Editor, range: { from: number; to: number }): void {
  editor.view.dispatch(editor.state.tr.setMeta(spellIgnoreOnceKey, { add: range }))
}

export function clearIgnoredOnce(editor: Editor): void {
  editor.view.dispatch(editor.state.tr.setMeta(spellIgnoreOnceKey, { clear: true }))
}

/**
 * Words Ignore All or Add to Dictionary set aside in this session. The word
 * lookup of the page does not see the session dictionary on every platform,
 * so the menu asks here first.
 */
const asideWords = new Map<string, 'ignored' | 'learned'>()

export function setWordAside(word: string, how: 'ignored' | 'learned'): void {
  asideWords.set(word, how)
}

export const wordSetAside = (word: string): boolean => asideWords.has(word)

/** Reset Ignored Words: the ignored ones are checked again, the learned ones stay */
export function forgetIgnoredWords(): void {
  for (const [word, how] of asideWords) if (how === 'ignored') asideWords.delete(word)
}

/** the checked word `pos` is in or touches; deleted and no-proofing text has none */
export function spellWordAt(doc: PmNode, pos: number): SpellWord | null {
  if (pos < 0 || pos > doc.content.size) return null
  const $pos = doc.resolve(pos)
  if (!$pos.parent.isTextblock) return null
  const words: SpellWord[] = []
  wordsOfBlock($pos.parent, $pos.before(), words)
  return words.find((w) => w.from <= pos && pos <= w.to) ?? null
}

/**
 * The misspelled word under a right-click or around the caret, from the
 * dictionary itself: the browser marks a word only after typing next to it,
 * so text that came with the file has no marker to report.
 */
export function misspellingAt(
  state: EditorState,
  pos: number,
  misspelled: (word: string) => boolean,
): SpellWord | null {
  const word = spellWordAt(state.doc, pos)
  if (!word || spellingDismissed(state, word)) return null
  return misspelled(word.word) ? word : null
}

/** Ignore Once on this occurrence, or Ignore All / Add to Dictionary on the word, took it out of the check */
export function spellingDismissed(state: EditorState, word: SpellWord): boolean {
  return wordSetAside(word.word) || ignoredOnce(state, word.from, word.to)
}

/**
 * Word's Next Misspelling: the first misspelled word from the caret on (the
 * one the caret is in counts, a selected one does not), going round the end
 * of the document; null when there is none.
 */
export function misspellingAfter(
  state: EditorState,
  misspelled: (word: string) => boolean,
): SpellWord | null {
  const { empty, to } = state.selection
  const known = new Map<string, boolean>()
  const wrong = (word: SpellWord): boolean => {
    if (spellingDismissed(state, word)) return false
    let bad = known.get(word.word)
    if (bad === undefined) known.set(word.word, (bad = misspelled(word.word)))
    return bad
  }
  const words = spellWords(state.doc)
  const ahead = (word: SpellWord) => (empty ? word.to >= to : word.from >= to)
  return (
    words.find((word) => ahead(word) && wrong(word)) ??
    words.find((word) => !ahead(word) && wrong(word)) ??
    null
  )
}
