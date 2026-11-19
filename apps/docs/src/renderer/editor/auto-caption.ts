/**
 * AutoCaption: the caption Word adds by itself when a table or an equation is
 * inserted. The choices are kept on this computer, as Word keeps them outside
 * the document.
 */
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import {
  formatSeqNumber,
  generateCaptionXml,
  SEQ_NUMBER_FORMATS,
  type Block,
  type SeqNumberFormat,
} from '@genoffice/docx-engine'
import { collectCaptions } from '../commands/reference-commands'
import { renumberAfterInsert, unreliableSeqLabelsOf } from './field-results'
import { t } from '../i18n/locale'

export const AUTO_CAPTION_KINDS = ['table', 'equation'] as const
export type AutoCaptionKind = (typeof AUTO_CAPTION_KINDS)[number]

export interface AutoCaptionRule {
  on: boolean
  label: string
  position: 'above' | 'below'
  format: SeqNumberFormat
}

export type AutoCaptionRules = Record<AutoCaptionKind, AutoCaptionRule>

const KEY = 'docs.autoCaption'

/** Word's own defaults: a table is captioned above it, everything else below */
export const defaultAutoCaptionRules = (): AutoCaptionRules => ({
  table: { on: false, label: t('ribbonCaptionTable'), position: 'above', format: 'ARABIC' },
  equation: { on: false, label: t('ribbonCaptionEquation'), position: 'below', format: 'ARABIC' },
})

export function autoCaptionRules(): AutoCaptionRules {
  const rules = defaultAutoCaptionRules()
  try {
    const stored = JSON.parse(localStorage.getItem(KEY) ?? '{}') as Partial<
      Record<AutoCaptionKind, Partial<AutoCaptionRule>>
    >
    for (const kind of AUTO_CAPTION_KINDS) {
      const s = stored?.[kind]
      if (!s || typeof s !== 'object') continue
      const label = typeof s.label === 'string' ? s.label.trim() : ''
      rules[kind] = {
        on: s.on === true,
        label: label || rules[kind].label,
        position:
          s.position === 'above' || s.position === 'below' ? s.position : rules[kind].position,
        format: SEQ_NUMBER_FORMATS.find((f) => f === s.format) ?? 'ARABIC',
      }
    }
  } catch {
    // unreadable choices count as none
  }
  return rules
}

export function saveAutoCaptionRules(rules: AutoCaptionRules): void {
  localStorage.setItem(KEY, JSON.stringify(rules))
}

let blocksOf: () => Block[] = () => []

/** the open document's parsed blocks: a caption read from the file has its SEQ field there */
export function setAutoCaptionBlocks(source: () => Block[]): void {
  blocksOf = source
}

const isKind = (node: PmNode, kind: AutoCaptionKind): boolean =>
  kind === 'table'
    ? node.type.name === 'docTable'
    : node.type.name === 'docProtected' && !!node.attrs.formulaDisplay

/**
 * Captions the block an insert just added, when AutoCaption is on for its
 * kind. `before` is the document as it was before the insert.
 */
export function autoCaptionInserted(editor: Editor, kind: AutoCaptionKind, before: PmNode): void {
  const rule = autoCaptionRules()[kind]
  if (!rule.on) return
  const { doc } = editor.state
  const start = before.content.findDiffStart(doc.content)
  if (start === null || start === undefined) return
  // the new block is the first of its kind at or after the change that the old document
  // did not hold: a paragraph split around it, or a cell that took a table, starts earlier
  const old = new Set<PmNode>()
  before.forEach((child) => old.add(child))
  let pos = -1
  let size = 0
  doc.forEach((child, offset) => {
    if (pos < 0 && offset >= start && isKind(child, kind) && !old.has(child)) {
      pos = offset
      size = child.nodeSize
    }
  })
  if (pos < 0) return
  const number =
    collectCaptions(editor, blocksOf()).filter((c) => c.label === rule.label).length + 1
  const caption = editor.schema.nodes.docProtected.create({
    docxIndex: null,
    blockType: 'passthrough',
    label: t('ribbonCaption'),
    genXml: generateCaptionXml(rule.label, number, '', {
      format: rule.format,
      dirty: unreliableSeqLabelsOf(blocksOf(), editor.schema).has(rule.label),
    }),
    fieldDisplay: { kind: 'text', left: `${rule.label} ${formatSeqNumber(number, rule.format)}` },
  })
  const tr = editor.state.tr.insert(rule.position === 'above' ? pos : pos + size, caption)
  // one undo takes the caption away with the block it came with
  editor.view.dispatch(tr.setMeta('appendedTransaction', tr))
  renumberAfterInsert(editor, blocksOf())
}
