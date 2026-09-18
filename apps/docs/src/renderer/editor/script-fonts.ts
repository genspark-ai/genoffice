import { Extension } from '@tiptap/core'
import { Plugin } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { cachedByDoc } from '../doc-cache'

// Font fallback cannot select a script when the Latin face itself covers CJK.
// Decorations keep the text/runs intact, including undo, copying and DOCX export.
const eastAsianText =
  /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}\u3000-\u303f\uff01-\uff60\uffe0-\uffee]+/gu

const decorations = cachedByDoc((doc) => {
  const ranges: Decoration[] = []
  doc.descendants((node, pos) => {
    if (!node.isText) return
    for (const match of node.text!.matchAll(eastAsianText)) {
      ranges.push(
        Decoration.inline(pos + match.index, pos + match.index + match[0].length, {
          class: 'doc-east-asian-font',
          style: 'font-family:var(--doc-east-asian-font, inherit)',
        }),
      )
    }
  })
  return DecorationSet.create(doc, ranges)
})
export const ScriptFonts = Extension.create({
  name: 'scriptFonts',
  addProseMirrorPlugins() {
    return [new Plugin({ props: { decorations: (state) => decorations(state.doc) } })]
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
    const matches = [...text.data.matchAll(eastAsianText)]
    if (!matches.length) continue
    const fragment = document.createDocumentFragment()
    let end = 0
    for (const match of matches) {
      fragment.append(text.data.slice(end, match.index))
      const span = document.createElement('span')
      span.className = 'doc-east-asian-font'
      span.style.fontFamily = 'var(--doc-east-asian-font, inherit)'
      span.textContent = match[0]
      fragment.append(span)
      end = match.index + match[0].length
    }
    fragment.append(text.data.slice(end))
    text.replaceWith(fragment)
  }
  return template.innerHTML
}
