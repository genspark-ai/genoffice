import { combineTransactionSteps, Extension, getChangedRanges, type Editor } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { AttrStep } from '@tiptap/pm/transform'
import type { Node as PmNode } from '@tiptap/pm/model'
import { displayCharSpacingTwips } from '@genoffice/docx-engine'
import { inheritedEaFamily } from '../text-layout/run-faces'

declare module '@tiptap/core' {
  interface Storage {
    balancedSpacing: { enabled: boolean | null; owner: Editor | null }
  }
}

/** Recompute derived paint in the same edit, without changing authored spacing. */
export const BalancedSpacingExtension = Extension.create({
  name: 'balancedSpacing',
  addStorage() {
    return { enabled: null as boolean | null, owner: null as Editor | null }
  },
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('balancedSpacing'),
        appendTransaction: (transactions, oldState, state) => {
          const edits = transactions.filter((tr) => tr.docChanged)
          if (!edits.length) return null
          const type = state.schema.marks.docTextStyle
          let tr: typeof state.tr | null = null
          const transform = combineTransactionSteps(oldState.doc, [...edits])
          const ranges = getChangedRanges(transform).map(({ newRange }) => newRange)
          transform.steps.forEach((step, index) => {
            if (step instanceof AttrStep) {
              const pos = transform.mapping.slice(index + 1).map(step.pos)
              ranges.push({ from: pos, to: pos })
            }
          })
          const update = (node: PmNode, pos: number, parent: PmNode) => {
            if (!node.isText) return
            const mark = node.marks.find((m) => m.type === type)
            if (!mark) return
            const attrs = mark.attrs
            const ownerStorage = this.storage.owner?.storage ?? this.editor.storage
            const enabled = ownerStorage.balancedSpacing.enabled ?? attrs.spacingDisplay != null
            if (!enabled && attrs.spacingDisplay == null && attrs.charSpacingDisplayTwips == null)
              return
            const inheritedFont = inheritedEaFamily(
              parent?.attrs.styleId,
              ownerStorage.listNumbering,
              attrs.styleId,
            )
            const context =
              enabled && attrs.charSpacingTwips != null
                ? {
                    eastAsiaFont:
                      attrs.eastAsiaFont ?? inheritedFont ?? (attrs.spacingDisplay || undefined),
                  }
                : undefined
            const spacingDisplay = context ? (context.eastAsiaFont ?? '') : null
            const display = context
              ? (displayCharSpacingTwips({
                  text: node.text!,
                  charSpacingTwips: attrs.charSpacingTwips,
                  spacingDisplay: context,
                }) ?? null)
              : null
            if (
              spacingDisplay === attrs.spacingDisplay &&
              display === attrs.charSpacingDisplayTwips
            )
              return
            tr ??= state.tr
            tr.addMark(
              pos,
              pos + node.nodeSize,
              type.create({ ...attrs, spacingDisplay, charSpacingDisplayTwips: display }),
            )
          }
          for (const range of ranges) {
            // Attribute/markup steps may touch only a paragraph's opening token.
            state.doc.nodesBetween(
              Math.max(0, range.from - 1),
              Math.min(state.doc.content.size, range.to + 1),
              (node, pos) => {
                if (!node.isTextblock) return
                node.forEach((child, offset) => update(child, pos + 1 + offset, node))
                return false
              },
            )
          }
          return tr
        },
      }),
    ]
  },
})
