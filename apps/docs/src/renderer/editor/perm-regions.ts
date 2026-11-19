import { Extension, type Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type Transaction } from '@tiptap/pm/state'
import { AttrStep, type Step } from '@tiptap/pm/transform'
import { Decoration, DecorationSet } from '@tiptap/pm/view'

import { PHASED_APPEND } from './streaming-tail-guard'

/**
 * Word's Restrict Editing ▸ "No changes (Read only)" with exceptions open to
 * everyone: the document is locked except for the regions its editing
 * permission ranges open (w:permStart w:edGrp="everyone"). The parser marks
 * those regions on the open top-level paragraphs and table cells
 * (`permEditable`, docx-engine perm-ranges.ts).
 *
 * While enabled, this extension
 * - refuses every edit that reaches outside an open region (the rest of the
 *   document stays read-only, as in Word), and
 * - highlights the open regions the way Word's "Highlight the regions I can
 *   edit" does: a highlight behind their text.
 *
 * Open regions are a whole table cell or a run of top-level paragraphs; an
 * edit may cross paragraphs inside one cell or across open top-level
 * paragraphs, never into a locked block or another cell.
 */

export interface PermRegionsStorage {
  /** the document is read-only protected and has open regions */
  enabled: boolean
}

declare module '@tiptap/core' {
  interface Storage {
    permRegions: PermRegionsStorage
  }
}

/** transactions carrying this meta skip the guard (the extension's own attr fixes) */
export const PERM_REGIONS_BYPASS = 'permRegionsBypass'

const pluginKey = new PluginKey<DecorationSet>('permRegions')
const REFRESH = 'permRegionsRefresh'

const CELL_TYPES = new Set(['docTableCell', 'docTableHeader'])
const TOP_TEXTBLOCKS = new Set(['docParagraph', 'docHeading', 'docListItem'])

const isOpenTopBlock = (node: PmNode): boolean =>
  TOP_TEXTBLOCKS.has(node.type.name) && node.attrs.permEditable === true

/**
 * The open region holding `pos`: the open table cell around it, or the open
 * top-level paragraph it is in. null when `pos` is in locked content.
 */
export function openRegionAt(doc: PmNode, pos: number): { start: number; depth: number } | null {
  const $pos = doc.resolve(Math.max(0, Math.min(pos, doc.content.size)))
  for (let d = $pos.depth; d > 0; d--) {
    const node = $pos.node(d)
    if (CELL_TYPES.has(node.type.name)) {
      return node.attrs.permEditable === true ? { start: $pos.before(d), depth: d } : null
    }
    if (d === 1) return isOpenTopBlock(node) ? { start: $pos.before(1), depth: 1 } : null
  }
  return null
}

/** every position from `from` to `to` lies in one open region (or in open top-level paragraphs only) */
export function rangeIsOpen(doc: PmNode, from: number, to: number): boolean {
  const start = openRegionAt(doc, from)
  if (!start) return false
  if (to <= from) return true
  const end = openRegionAt(doc, to)
  if (!end) return false
  if (start.start === end.start) return true
  if (start.depth !== 1 || end.depth !== 1) return false
  let open = true
  doc.nodesBetween(from, to, (node, _pos, parent) => {
    if (parent === doc && !isOpenTopBlock(node)) open = false
    return false
  })
  return open
}

function stepIsOpen(step: Step, doc: PmNode): boolean {
  if (step instanceof AttrStep) {
    // paragraph formatting inside an open region; cell, row and table
    // attributes stay locked (they would regenerate the whole table)
    const node = doc.nodeAt(step.pos)
    return !!node?.isTextblock && openRegionAt(doc, step.pos + 1) !== null
  }
  const range = step as { from?: unknown; to?: unknown }
  if (typeof range.from === 'number' && typeof range.to === 'number') {
    return rangeIsOpen(doc, range.from, range.to)
  }
  // document attributes and anything else without a range stay locked
  return false
}

/** a whole-document replace: the initial load or a remount, never a user edit */
function replacesWholeDocument(tr: Transaction): boolean {
  return tr.steps.some((step, i) => {
    const s = step as { from?: unknown; to?: unknown }
    return s.from === 0 && typeof s.to === 'number' && s.to >= tr.docs[i]!.content.size
  })
}

export function permRegionsAllow(tr: Transaction): boolean {
  if (!tr.docChanged) return true
  if (tr.getMeta(PHASED_APPEND) || tr.getMeta(PERM_REGIONS_BYPASS)) return true
  // automatic follow-ups of an allowed edit (normalizations) are not user edits
  if (tr.getMeta('appendedTransaction')) return true
  if (replacesWholeDocument(tr)) return true
  return tr.steps.every((step, i) => stepIsOpen(step, tr.docs[i]!))
}

/**
 * Height of a paragraph's mark when the line is taller than the text: the
 * tallest in-line picture (a floating one does not size the line). 0 = text.
 */
export function paragraphMarkHeight(node: PmNode): number {
  let h = 0
  node.forEach((child) => {
    if (child.type.name !== 'docInlineImage' || child.attrs.wrap) return
    const px = Number(child.attrs.heightPx)
    if (px > h) h = px
  })
  return h
}

function paragraphMark(height: number): HTMLElement {
  const span = document.createElement('span')
  span.className = 'perm-region-mark'
  span.contentEditable = 'false'
  if (height > 0) {
    span.style.setProperty('--perm-mark-h', `${height}px`)
    span.style.setProperty('--perm-mark-drop', '0px')
  }
  return span
}

/**
 * Word's region bracket: drawn over the text at the region's edge, taking no
 * width, so a template that types its own "[" / "]" there shows one bracket
 * (Word's, on top), not two.
 */
function regionBracket(side: 'open' | 'close'): HTMLElement {
  const span = document.createElement('span')
  span.className = `perm-region-bracket perm-region-bracket-${side}`
  span.contentEditable = 'false'
  const glyph = document.createElement('span')
  glyph.textContent = side === 'open' ? '[' : ']'
  span.appendChild(glyph)
  return span
}

/**
 * Word's "Highlight the regions I can edit": a highlight behind the text of
 * every paragraph in an open region, tables nested in an open cell included,
 * and on each paragraph's mark: a patch at the paragraph's end, as tall as
 * its line (the full height beside a picture), taking no width. Each region
 * gets Word's olive brackets at its two ends (regionBracket).
 */
function regionDecorations(doc: PmNode): DecorationSet {
  const decorations: Decoration[] = []
  const mark = (node: PmNode, pos: number): boolean => {
    if (node.type.name === 'docNestedTable') {
      // an atom drawn from its attributes: the highlight goes on through CSS
      decorations.push(Decoration.node(pos, pos + node.nodeSize, { class: 'perm-region-nested' }))
      return false
    }
    if (!node.isTextblock) return true
    const end = pos + node.nodeSize - 1
    region.push({ start: pos + 1, end })
    if (node.content.size > 0) {
      decorations.push(Decoration.inline(pos + 1, end, { class: 'perm-region' }))
    }
    const height = paragraphMarkHeight(node)
    decorations.push(
      Decoration.widget(end, () => paragraphMark(height), {
        side: 2,
        ignoreSelection: true,
        key: `perm-mark-${height}`,
      }),
    )
    return false
  }
  let region: Array<{ start: number; end: number }> = []
  const closeRegion = () => {
    const first = region[0]
    const last = region[region.length - 1]
    if (first && last) {
      decorations.push(
        Decoration.widget(first.start, () => regionBracket('open'), {
          side: -1,
          ignoreSelection: true,
          key: 'perm-open',
        }),
        Decoration.widget(last.end, () => regionBracket('close'), {
          side: 1,
          ignoreSelection: true,
          key: 'perm-close',
        }),
      )
    }
    region = []
  }
  doc.forEach((node, offset) => {
    if (isOpenTopBlock(node)) {
      // consecutive open top-level paragraphs are one region
      mark(node, offset)
      return
    }
    closeRegion()
    if (node.type.name !== 'docTable') return
    node.descendants((cell, cellPos) => {
      if (!CELL_TYPES.has(cell.type.name)) return true
      if (cell.attrs.permEditable === true) {
        const base = offset + 1 + cellPos + 1
        cell.descendants((child, childPos) => mark(child, base + childPos))
        closeRegion()
      }
      return false
    })
  })
  closeRegion()
  return DecorationSet.create(doc, decorations)
}

export const PermRegionsExtension = Extension.create<object, PermRegionsStorage>({
  name: 'permRegions',
  addStorage() {
    return { enabled: false }
  },
  addProseMirrorPlugins() {
    const storage = this.storage
    return [
      new Plugin<DecorationSet>({
        key: pluginKey,
        state: {
          init: () => DecorationSet.empty,
          apply: (tr, old, _oldState, newState) => {
            if (!storage.enabled) return DecorationSet.empty
            if (tr.docChanged || tr.getMeta(REFRESH) || old === DecorationSet.empty) {
              return regionDecorations(newState.doc)
            }
            return old
          },
        },
        props: {
          decorations: (state) => pluginKey.getState(state),
        },
        filterTransaction: (tr) => !storage.enabled || permRegionsAllow(tr),
        // paragraphs an allowed edit creates at the top level (Enter, a paste of
        // several paragraphs) belong to the open region they were made in
        appendTransaction: (trs, _oldState, newState) => {
          if (!storage.enabled) return null
          const ranges: Array<[number, number]> = []
          trs.forEach((tr, t) => {
            if (!tr.docChanged || tr.getMeta(PERM_REGIONS_BYPASS)) return
            if (tr.getMeta(PHASED_APPEND) || replacesWholeDocument(tr)) return
            const after = trs.slice(t + 1)
            tr.mapping.maps.forEach((map, i) => {
              const rest = tr.mapping.slice(i + 1)
              map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
                let from = rest.map(newStart, -1)
                let to = rest.map(newEnd, 1)
                for (const later of after) {
                  from = later.mapping.map(from, -1)
                  to = later.mapping.map(to, 1)
                }
                ranges.push([from, to])
              })
            })
          })
          if (ranges.length === 0) return null
          const fix = newState.tr
          for (const [from, to] of ranges) {
            newState.doc.nodesBetween(from, to, (node, pos, parent) => {
              if (
                parent === newState.doc &&
                TOP_TEXTBLOCKS.has(node.type.name) &&
                !node.attrs.permEditable
              ) {
                fix.setNodeAttribute(pos, 'permEditable', true)
              }
              return false
            })
          }
          if (!fix.docChanged) return null
          return fix.setMeta(PERM_REGIONS_BYPASS, true).setMeta('addToHistory', false)
        },
      }),
    ]
  },
})

/** turn the guard and highlight on or off for this editor and redraw */
export function setPermRegionsEnabled(editor: Editor, enabled: boolean): void {
  const storage = editor.storage.permRegions
  if (!storage || storage.enabled === enabled) return
  storage.enabled = enabled
  if (!editor.isDestroyed) editor.view.dispatch(editor.state.tr.setMeta(REFRESH, true))
}
