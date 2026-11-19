import { Extension } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { WORD_AUTO_SPACING_PT } from '../line-metrics'

const key = new PluginKey<DecorationSet>('floatTableSpacing')

export const AFTER_FLOAT_TABLE_CLASS = 'doc-after-tblp'

function spacingPt(node: ProseMirrorNode, side: 'Before' | 'After'): number | null {
  if (node.type.name !== 'docParagraph') return null
  if (node.attrs[`space${side}Auto`]) return WORD_AUTO_SPACING_PT
  const v = node.attrs[`space${side}`]
  return v == null ? null : Number(v) / 20
}

/**
 * Word does not collapse paragraph spacing across a floating table: the
 * paragraph after a w:tblpPr table sits its full space-before below the
 * previous paragraph's space-after (Word for Mac, page-anchored table between
 * two 9pt/9pt paragraphs: 18pt, the plain paragraph pair 9pt). CSS collapses
 * the two margins to their max, so the anchor paragraph carries the sum;
 * the stylesheet applies it only while the table is really floated.
 */
export function floatTableSpacingDecorations(doc: ProseMirrorNode): DecorationSet {
  const decos: Decoration[] = []
  let pos = 0
  let prev: ProseMirrorNode | null = null
  let prevTable: ProseMirrorNode | null = null
  doc.forEach((node) => {
    if (
      prevTable &&
      prev &&
      node.type.name === 'docParagraph' &&
      (node.attrs.spaceBefore != null || node.attrs.spaceBeforeAuto)
    ) {
      const after = spacingPt(prev, 'After')
      if (after) {
        const sum = (spacingPt(node, 'Before') ?? 0) + after
        decos.push(
          Decoration.node(pos, pos + node.nodeSize, {
            class: AFTER_FLOAT_TABLE_CLASS,
            style: `--tblp-sp-before:${sum.toFixed(1)}pt`,
          }),
        )
      }
    }
    const floated =
      node.type.name === 'docTable' &&
      !node.attrs.tblFloatSuppressed &&
      (node.attrs.tblFloat === 'left' || node.attrs.tblFloat === 'right')
    if (floated) prevTable = node
    else {
      prevTable = null
      prev = node
    }
    pos += node.nodeSize
  })
  return DecorationSet.create(doc, decos)
}

export const FloatTableSpacingExtension = Extension.create({
  name: 'floatTableSpacing',
  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: (_, state) => floatTableSpacingDecorations(state.doc),
          apply(tr, set) {
            return tr.docChanged ? floatTableSpacingDecorations(tr.doc) : set
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)
          },
        },
      }),
    ]
  },
})
