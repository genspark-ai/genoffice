import { Extension } from '@tiptap/core'
import type { SectionInfo } from '@genoffice/docx-engine'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { t } from '../i18n/locale'
import { containsNode, localEditBlocks, touchedNeedsRecompute } from './local-edit'

/**
 * Word's formatting marks for breaks (Home > ¶). The pilcrow itself is CSS;
 * these are DOM widgets: a chip at a w:br page/column break, an arrow at a
 * line / text-wrapping break, and a section-break chip on the last line of
 * every section-ending paragraph (labelled with how the NEXT section starts,
 * as Word does). Screen only: the html export lifts .show-marks and print
 * hides the widgets.
 */
export interface FormattingMarksMeta {
  on: boolean
  /** docxIndex of each section-break paragraph → start type of the section it opens */
  breaks: ReadonlyMap<number, SectionInfo['startType']>
}

export const formattingMarksPluginKey = new PluginKey<FormattingMarksState>('formattingMarks')

interface FormattingMarksState {
  meta: FormattingMarksMeta
  decos: DecorationSet
}

export const sectionBreakMap = (
  sections: readonly SectionInfo[],
): ReadonlyMap<number, SectionInfo['startType']> => {
  const map = new Map<number, SectionInfo['startType']>()
  for (let i = 0; i + 1 < sections.length; i++) {
    map.set(sections[i].lastBlockIndex, sections[i + 1].startType)
  }
  return map
}

const startTypeLabel = (type: SectionInfo['startType']): string =>
  type === 'continuous'
    ? t('appBreakContinuous')
    : type === 'evenPage'
      ? t('appBreakEvenPage')
      : type === 'oddPage'
        ? t('appBreakOddPage')
        : t('appBreakNextPage')

const sectionBreakLabel = (type: SectionInfo['startType']): string =>
  `${t('editorSectionBreak')} (${startTypeLabel(type)})`

const chip = (cls: string, label: string) => () => {
  const span = document.createElement('span')
  span.className = `doc-fmt-mark doc-fmt-chip ${cls}`
  span.contentEditable = 'false'
  span.setAttribute('aria-hidden', 'true')
  const text = document.createElement('span')
  text.className = 'doc-fmt-label'
  text.textContent = label
  span.append(rule(), text, rule())
  return span
}

const rule = () => {
  const span = document.createElement('span')
  span.className = 'doc-fmt-rule'
  return span
}

const BLOCK_SELECTOR = 'p, h1, h2, h3, h4, h5, h6, li, .doc-li, .doc-textbox-para, td, th'

/** Word runs the rule from the break to the right margin: size each chip to
 *  the rest of its line (a chip that would not fit wraps, then refits) */
function fitChip(chip: HTMLElement, pass = 0): void {
  const block = chip.parentElement?.closest<HTMLElement>(BLOCK_SELECTOR)
  if (!block) return
  const br = block.getBoundingClientRect()
  if (br.width === 0 || block.offsetWidth === 0) return
  const scale = br.width / block.offsetWidth
  const cs = getComputedStyle(block)
  const right = br.right - (parseFloat(cs.paddingRight) + parseFloat(cs.borderRightWidth)) * scale
  const left = chip.getBoundingClientRect().left
  const label = chip.querySelector<HTMLElement>('.doc-fmt-label')
  const min = label ? Math.ceil(label.getBoundingClientRect().width / scale) + 8 : 0
  const width = Math.max(min, Math.floor((right - left) / scale))
  if (chip.style.width === `${width}px`) return
  chip.style.width = `${width}px`
  if (pass === 0) fitChip(chip, 1)
}

function fitAll(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>('.doc-fmt-chip').forEach((chip) => fitChip(chip))
}

const glyph = (cls: string) => () => {
  const span = document.createElement('span')
  span.className = `doc-fmt-mark doc-fmt-glyph ${cls}`
  span.contentEditable = 'false'
  span.setAttribute('aria-hidden', 'true')
  return span
}

function build(doc: PmNode, meta: FormattingMarksMeta): DecorationSet {
  if (!meta.on) return DecorationSet.empty
  const decos: Decoration[] = []
  doc.descendants((node, pos) => {
    if (node.type.name === 'hardBreak') {
      if (node.attrs.mathBreak) return false
      const widget = node.attrs.pageBreak
        ? chip('doc-fmt-pagebr', t('editorPageBreak'))
        : node.attrs.colBreak
          ? chip('doc-fmt-colbr', t('appScColumnBreak'))
          : glyph(node.attrs.wrapBreak ? 'doc-fmt-wrapbr' : 'doc-fmt-br')
      decos.push(Decoration.widget(pos, widget, { side: -1, key: `fmt-br-${pos}` }))
      return false
    }
    if (node.isInline) return false
    const idx = node.attrs.docxIndex
    if (typeof idx !== 'number') return true
    const type = meta.breaks.get(idx)
    if (type === undefined) return true
    const label = sectionBreakLabel(type)
    decos.push(
      Decoration.node(pos, pos + node.nodeSize, {
        class: 'doc-sect-break-mark',
        'data-sect-break-label': label,
      }),
    )
    if (node.isTextblock) {
      decos.push(
        Decoration.widget(pos + node.nodeSize - 1, chip('doc-fmt-sectbr', label), {
          side: 1,
          key: `fmt-sect-${pos}`,
        }),
      )
    }
    return true
  })
  return DecorationSet.create(doc, decos)
}

const OFF: FormattingMarksMeta = { on: false, breaks: new Map() }

export const FormattingMarksExtension = Extension.create({
  name: 'formattingMarks',
  addProseMirrorPlugins() {
    return [
      new Plugin<FormattingMarksState>({
        key: formattingMarksPluginKey,
        state: {
          init: () => ({ meta: OFF, decos: DecorationSet.empty }),
          apply(tr, old) {
            const meta = tr.getMeta(formattingMarksPluginKey) as FormattingMarksMeta | undefined
            if (meta) return { meta, decos: build(tr.doc, meta) }
            if (!old.meta.on || !tr.docChanged) return old
            const touched = localEditBlocks(tr)
            if (!touched) return { meta: old.meta, decos: build(tr.doc, old.meta) }
            const mapped = old.decos.map(tr.mapping, tr.doc)
            const marked = (n: PmNode) =>
              n.type.name === 'hardBreak' ||
              (typeof n.attrs.docxIndex === 'number' && old.meta.breaks.has(n.attrs.docxIndex))
            return touchedNeedsRecompute(tr, touched, mapped, (node) => containsNode(node, marked))
              ? { meta: old.meta, decos: build(tr.doc, old.meta) }
              : { meta: old.meta, decos: mapped }
          },
        },
        props: {
          decorations(state) {
            return this.getState(state)?.decos ?? DecorationSet.empty
          },
        },
        view(view) {
          let frame = 0
          const schedule = () => {
            if (frame) return
            frame = requestAnimationFrame(() => {
              frame = 0
              fitAll(view.dom)
            })
          }
          const ro = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(schedule)
          ro?.observe(view.dom)
          schedule()
          return {
            update(v, prev) {
              if (
                v.state.doc !== prev.doc ||
                formattingMarksPluginKey.getState(v.state) !==
                  formattingMarksPluginKey.getState(prev)
              )
                schedule()
            },
            destroy() {
              ro?.disconnect()
              if (frame) cancelAnimationFrame(frame)
            },
          }
        },
      }),
    ]
  },
})
