/**
 * Endnotes at the end of each section (w:endnotePr w:pos="sectEnd"): which
 * section's end prints an endnote, and the rows drawn there. The rows are
 * widgets in the flow after a section's last block, one per note, so the
 * pagination pass measures and pages them like any block; the last section's
 * notes stay in the end-of-document area.
 */
import { Extension } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'
import type { SectionInfo } from '@genoffice/docx-engine'
import { SECT_NOTE_CLASS, type PageNoteItem } from '../pagination-types'
import { noteSeparatorPx, noteSeparatorRulePx } from '../line-metrics'
import { t } from '../i18n/locale'
import { docHyphenationSig } from './hyphenation'
import { appendNoteRuns, noteRowBox } from './note-dom'

type SectionLike = Pick<SectionInfo, 'lastBlockIndex' | 'sectPrXml'> & { pendingBreak?: boolean }

export interface EndnoteHomes {
  /** endnote id → the section holding its reference */
  sectionOf: Map<string, number>
  /** per section: its closing break paragraph is in the document */
  closed: boolean[]
  sig: string
}

/** the section of each top-level block (ai/pending-sections' rule, in one walk) and which sections are closed */
export function topBlockSections(
  doc: PmNode,
  sections: readonly SectionLike[],
): { sectionOf: number[]; closed: boolean[] } {
  const last = sections.length - 1
  const closed = sections.map(() => false)
  const sectionOf: number[] = []
  let cur = 0
  doc.forEach((node) => {
    const di = node.attrs?.docxIndex as number | null | undefined
    if (di !== null && di !== undefined) {
      const s = sections.findIndex((sec) => di <= sec.lastBlockIndex)
      cur = s >= 0 ? s : Math.max(last, 0)
      if (sections[cur]?.lastBlockIndex === di && !sections[cur]?.pendingBreak) closed[cur] = true
      sectionOf.push(cur)
      return
    }
    sectionOf.push(cur)
    const genXml = node.attrs?.genXml
    const sec = sections[cur]
    if (
      sec?.pendingBreak &&
      typeof genXml === 'string' &&
      genXml.includes('<w:sectPr') &&
      genXml.includes(sec.sectPrXml)
    ) {
      closed[cur] = true
      cur = Math.min(cur + 1, Math.max(last, 0))
    }
  })
  if (last >= 0) closed[last] = true
  return { sectionOf, closed }
}

/** where each endnote reference sits */
export function endnoteHomesOf(doc: PmNode, sections: readonly SectionLike[]): EndnoteHomes {
  const { sectionOf: blockSection, closed } = topBlockSections(doc, sections)
  const sectionOf = new Map<string, number>()
  doc.forEach((block, _offset, index) => {
    block.descendants((node) => {
      if (node.type.name !== 'docNoteRef') return true
      if (node.attrs.kind === 'endnote' && !sectionOf.has(String(node.attrs.id)))
        sectionOf.set(String(node.attrs.id), blockSection[index] ?? 0)
      return false
    })
  })
  const sig = `${closed.map((c) => (c ? 1 : 0)).join('')}|${[...sectionOf].map(([id, s]) => `${id}:${s}`).join(',')}`
  return { sectionOf, closed, sig }
}

/**
 * The section whose end prints each section's endnotes: its own, unless it
 * suppresses them (w:noEndnote) or its break is gone — then the next one that
 * prints. The last section always prints.
 */
export function endnotePrintSections(
  suppressed: readonly boolean[],
  closed: readonly boolean[],
): number[] {
  const last = suppressed.length - 1
  const out: number[] = new Array(suppressed.length)
  let next = last
  for (let i = last; i >= 0; i--) {
    if (i === last || (!suppressed[i] && closed[i])) next = i
    out[i] = next
  }
  return out
}

export interface SectionEndnoteGroup {
  /** the section these notes print at the end of (never the last one) */
  section: number
  items: PageNoteItem[]
}

/** split the endnote list into the groups printed at section ends and the rest, printed at the document end */
export function groupSectionEndnotes(
  items: readonly PageNoteItem[],
  homes: EndnoteHomes,
  suppressed: readonly boolean[],
): { groups: SectionEndnoteGroup[]; tail: PageNoteItem[] } {
  const printAt = endnotePrintSections(suppressed, homes.closed)
  const last = suppressed.length - 1
  const bySection = new Map<number, PageNoteItem[]>()
  const tail: PageNoteItem[] = []
  // in the order of the references, as they are numbered; the list is in the order they were added
  const order = new Map([...homes.sectionOf.keys()].map((id, i) => [id, i]))
  const rank = (item: PageNoteItem) => order.get(item.id) ?? Infinity
  for (const item of [...items].sort((a, b) => rank(a) - rank(b))) {
    const home = homes.sectionOf.get(item.id)
    const at = home === undefined ? last : (printAt[home] ?? last)
    if (at >= last) tail.push(item)
    else bySection.set(at, [...(bySection.get(at) ?? []), item])
  }
  const groups = [...bySection]
    .sort((a, b) => a[0] - b[0])
    .map(([section, list]) => ({ section, items: list }))
  return { groups, tail }
}

interface SectionEndnotesState {
  groups: SectionEndnoteGroup[]
  sections: readonly SectionLike[]
  onEdit: (id: string) => void
  decos: DecorationSet
}

const key = new PluginKey<SectionEndnotesState>('sectionEndnotes')

/** rows sort ahead of the page gap (side -1) that opens the next section's page */
const FIRST_SIDE = -100000

function rowEl(item: PageNoteItem, first: boolean): HTMLElement {
  const row = document.createElement('div')
  const box = noteRowBox(item)
  row.className = ['page-gap-note', SECT_NOTE_CLASS, ...box.classes].join(' ')
  if (first) {
    row.classList.add('page-sect-note-first')
    row.style.setProperty('--note-sep-h', `${noteSeparatorPx('endnote')}px`)
    row.style.setProperty('--note-sep-rule', `${noteSeparatorRulePx('endnote')}px`)
  }
  row.contentEditable = 'false'
  row.title = t('appEditEndnote')
  row.dataset.noteId = item.id
  for (const [k, v] of Object.entries(box.style)) {
    if (k.startsWith('--')) row.style.setProperty(k, v)
    else row.style[k as 'lineHeight'] = v
  }
  appendNoteRuns(row, item, 'endnote')
  return row
}

const itemSig = (item: PageNoteItem) =>
  `h${docHyphenationSig()}:` +
  `${item.no}:${item.fmt ?? ''}:${item.noRefMark ? 0 : 1}:${item.lineHeightPx ?? ''}:${item.fontSizePt ?? ''}:${item.fontFamily ?? ''}:${item.richParas ? JSON.stringify(item.richParas) : item.text}`

function buildDecos(doc: PmNode, state: Omit<SectionEndnotesState, 'decos'>): DecorationSet {
  if (state.groups.length === 0) return DecorationSet.empty
  const { sectionOf } = topBlockSections(doc, state.sections)
  const endOf = new Map<number, number>()
  doc.forEach((node, offset, index) => endOf.set(sectionOf[index]!, offset + node.nodeSize))
  const decos: Decoration[] = []
  for (const group of state.groups) {
    const pos = endOf.get(group.section)
    if (pos === undefined) continue
    group.items.forEach((item, i) => {
      const side = FIRST_SIDE + i * 2
      decos.push(
        Decoration.widget(
          pos,
          () => {
            const el = rowEl(item, i === 0)
            el.dataset.side = String(side)
            return el
          },
          { side, key: `sect-endnote-${group.section}-${i}-${itemSig(item)}`, sectNote: true },
        ),
      )
    })
  }
  return DecorationSet.create(doc, decos)
}

export const SectionEndnotesExtension = Extension.create({
  name: 'sectionEndnotes',

  addProseMirrorPlugins() {
    return [
      new Plugin<SectionEndnotesState>({
        key,
        state: {
          init: () => ({
            groups: [],
            sections: [],
            onEdit: () => undefined,
            decos: DecorationSet.empty,
          }),
          apply(tr, prev) {
            const next = tr.getMeta(key) as Omit<SectionEndnotesState, 'decos'> | undefined
            if (!next && !tr.docChanged) return prev
            const base = next ?? prev
            if (base.groups.length === 0 && prev.groups.length === 0)
              return { ...base, decos: prev.decos }
            return { ...base, decos: buildDecos(tr.doc, base) }
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)?.decos
          },
          handleDOMEvents: {
            // read at event time: a reused row outlives the handler it was built under
            dblclick(view, event) {
              const row = (event.target as Element | null)?.closest?.(`.${SECT_NOTE_CLASS}`)
              const id = row instanceof HTMLElement ? row.dataset.noteId : undefined
              if (!id) return false
              key.getState(view.state)?.onEdit(id)
              return true
            },
          },
        },
      }),
    ]
  },
})

/** set the endnote groups printed at section ends (an empty list clears them) */
export function setSectionEndnotes(
  view: EditorView,
  groups: SectionEndnoteGroup[],
  sections: readonly SectionLike[],
  onEdit: (id: string) => void,
): void {
  const cur = key.getState(view.state)
  if (!cur || (groups.length === 0 && cur.groups.length === 0)) return
  view.dispatch(
    view.state.tr.setMeta('addToHistory', false).setMeta(key, { groups, sections, onEdit }),
  )
}

/** a section-end note row's position and widget side: where a page gap that opens at it goes */
export function sectionNoteSlot(
  view: EditorView,
  el: HTMLElement,
): { pos: number; side: number } | null {
  if (!el.classList.contains(SECT_NOTE_CLASS)) return null
  const side = Number(el.dataset.side)
  if (!Number.isFinite(side)) return null
  try {
    return { pos: view.posAtDOM(el, 0), side: side - 1 }
  } catch {
    return null
  }
}
