import type { AnyExtension } from '@tiptap/core'
import { Extension } from '@tiptap/core'
import { Plugin } from '@tiptap/pm/state'
import { BlockMath, InlineMath } from '@tiptap/extension-mathematics'
import { openMathEditor } from './mathEdit'

/**
 * Stricter inline tokenizer than the upstream default (`$...$` with any
 * content): the content must not start or end with whitespace and the
 * closing `$` must not be followed by a digit, so running text with
 * currency amounts ("paid $5 and $10") never turns into formulas.
 */
const STRICT_INLINE_MATH_RE = /^\$(?!\s)([^$\n]*[^\s$])\$(?!\d)/

const StrictInlineMath = InlineMath.extend({
  markdownTokenizer: {
    name: 'inlineMath',
    level: 'inline',
    start: (src: string) => src.indexOf('$'),
    tokenize: (src: string) => {
      const match = STRICT_INLINE_MATH_RE.exec(src)
      if (!match) return undefined
      return { type: 'inlineMath', raw: match[0], latex: match[1].trim() }
    },
  },
})

/** Math nodes are atoms — clicking one opens the LaTeX edit popover. */
const MathClickEdit = Extension.create({
  name: 'mathClickEdit',

  addProseMirrorPlugins() {
    const editor = this.editor
    return [
      new Plugin({
        props: {
          handleClickOn: (view, _pos, node, nodePos, event) => {
            if (node.type.name !== 'blockMath' && node.type.name !== 'inlineMath') return false
            if (!view.editable) return false
            const target = event.target as HTMLElement | null
            const anchor = target?.closest?.('.tiptap-mathematics-render') ?? target
            if (!anchor) return false
            openMathEditor(editor, { pos: nodePos, anchor: anchor.getBoundingClientRect() })
            return true
          },
        },
      }),
    ]
  },
})

export function buildMathExtensions(): AnyExtension[] {
  return [BlockMath, StrictInlineMath, MathClickEdit]
}
