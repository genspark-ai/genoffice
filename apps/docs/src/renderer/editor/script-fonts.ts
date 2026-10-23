import { Extension } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { touchedTopLevelBlocks } from './touched-blocks'

// Font fallback cannot select a script when the Latin face itself covers CJK.
// Decorations keep the text/runs intact, including undo, copying and DOCX export.
const eastAsianText =
  /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}\u3000-\u303f\uff01-\uff60\uffe0-\uffee]+/gu

// Halfwidth kana/hangul/symbols share the East Asian face but count as
// single-byte text for the balanced character grid (Word probe 2026-10-22).
const halfwidthForm = /[\uff61-\uffdc\uffe8-\uffee]/u
const halfwidthRun = /[\uff61-\uffdc\uffe8-\uffee]+|[^\uff61-\uffdc\uffe8-\uffee]+/gu

const EAST_ASIAN_CLASS = 'doc-east-asian-font'
const HALFWIDTH_CLASS = 'doc-east-asian-halfwidth'
const FONT_STYLE = 'font-family:var(--doc-east-asian-font, inherit)'

const classFor = (text: string): string =>
  halfwidthForm.test(text) ? `${EAST_ASIAN_CLASS} ${HALFWIDTH_CLASS}` : EAST_ASIAN_CLASS

/** East Asian runs in `text`, split so each piece is uniformly half- or full-width */
export function eastAsianSegments(text: string): { index: number; text: string }[] {
  const out: { index: number; text: string }[] = []
  for (const match of text.matchAll(eastAsianText)) {
    for (const part of match[0].matchAll(halfwidthRun)) {
      out.push({ index: match.index + part.index, text: part[0] })
    }
  }
  return out
}

function decorationsIn(doc: PmNode, from: number, to: number): Decoration[] {
  const ranges: Decoration[] = []
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) return
    for (const seg of eastAsianSegments(node.text!)) {
      ranges.push(
        Decoration.inline(pos + seg.index, pos + seg.index + seg.text.length, {
          class: classFor(seg.text),
          style: FONT_STYLE,
        }),
      )
    }
  })
  return ranges
}

const key = new PluginKey<DecorationSet>('scriptFonts')

export const ScriptFonts = Extension.create({
  name: 'scriptFonts',
  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: (_, state) =>
            DecorationSet.create(state.doc, decorationsIn(state.doc, 0, state.doc.content.size)),
          // Rebuilding for the whole document on every keystroke costs blocks×text
          // on long files; only the top-level blocks the transaction touched are
          // rescanned, the rest of the set is mapped through the change.
          apply(tr, old) {
            if (!tr.docChanged) return old
            const touched = touchedTopLevelBlocks(tr)
            if (!touched)
              return DecorationSet.create(tr.doc, decorationsIn(tr.doc, 0, tr.doc.content.size))
            let set = old.map(tr.mapping, tr.doc)
            for (const pos of touched) {
              const block = tr.doc.nodeAt(pos)
              if (!block) continue
              const to = pos + block.nodeSize
              set = set.remove(set.find(pos, to))
              set = set.add(tr.doc, decorationsIn(tr.doc, pos, to))
            }
            return set
          },
        },
        props: {
          decorations: (state) => key.getState(state),
        },
      }),
    ]
  },
})

/** The read-only split pane uses serialized HTML, which has no PM decorations. */
export function scriptFontHtml(html: string): string {
  const template = document.createElement('template')
  template.innerHTML = html
  const walker = document.createTreeWalker(template.content, NodeFilter.SHOW_TEXT)
  const texts: Text[] = []
  while (walker.nextNode()) texts.push(walker.currentNode as Text)
  for (const text of texts) {
    if (text.parentElement?.closest('script, style, svg, math')) continue
    const segments = eastAsianSegments(text.data)
    if (!segments.length) continue
    const fragment = document.createDocumentFragment()
    let end = 0
    for (const seg of segments) {
      fragment.append(text.data.slice(end, seg.index))
      const span = document.createElement('span')
      span.className = classFor(seg.text)
      span.style.fontFamily = 'var(--doc-east-asian-font, inherit)'
      span.textContent = seg.text
      fragment.append(span)
      end = seg.index + seg.text.length
    }
    fragment.append(text.data.slice(end))
    text.replaceWith(fragment)
  }
  return template.innerHTML
}
