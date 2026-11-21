import type { Editor } from '@tiptap/core'
import type { Mark } from '@tiptap/pm/model'
import { NodeSelection, type EditorState } from '@tiptap/pm/state'
import { cssFontFamily } from '../text-layout/font-rules/css-chains'
import {
  generateRubyXml,
  rubyAlignOf,
  rubyReadingStyleOf,
  setRubyAlignXml,
  setRubyBaseXml,
  setRubyReadingStyleXml,
  setRubyReadingXml,
  type RubyAlign,
} from '@genoffice/docx-engine'

/** what a phonetic guide would be put on: a run of selected text, or a guide already there */
export interface RubyTarget {
  from: number
  to: number
  base: string
  /** the reading the selected guide has; empty for plain text */
  rt: string
  /** the alignment the selected guide has; null when it names none the dialog offers */
  align: RubyAlign | null
  /** how the reading is set: what the guide says, or what a new one on this text would get */
  reading: RubyReading
}

/** a reading's font ('' = the base text's), its size and its distance above the base, in points */
export interface RubyReading {
  font: string
  sizePt: number
  offsetPt: number
}

const DEFAULT_BASE_HALF_POINTS = 21

/** Word raises a reading by the base size less a point; anything more is the dialog's distance */
const restRaise = (baseHalfPoints: number): number => Math.max(0, baseHalfPoints - 2)

/** the reading of a guide as the dialog and the page show it */
export function rubyReadingOf(xml: string, baseHalfPoints = DEFAULT_BASE_HALF_POINTS): RubyReading {
  const style = rubyReadingStyleOf(xml)
  const base = style.baseSizeHalfPoints ?? baseHalfPoints
  const raise = style.raiseHalfPoints ?? restRaise(base)
  return {
    font: style.font ?? '',
    sizePt: (style.sizeHalfPoints ?? Math.max(2, Math.round(base / 2))) / 2,
    offsetPt: Math.max(0, raise - restRaise(base)) / 2,
  }
}

/** the inline style of a guide's reading on the page; empty for one that says nothing */
export function rubyReadingCss(xml: string): string {
  const style = rubyReadingStyleOf(xml)
  const { font, sizePt, offsetPt } = rubyReadingOf(xml)
  return [
    style.sizeHalfPoints ? `font-size:${sizePt}pt` : '',
    font ? `font-family:${cssFontFamily(font)}` : '',
    // moved, not padded: the line keeps its height
    offsetPt > 0 ? `position:relative;bottom:${offsetPt}pt` : '',
  ]
    .filter(Boolean)
    .join(';')
}

/** Word's limit on the base text of one guide */
export const RUBY_BASE_MAX = 30

/** the selected guide, or the selected text when it is one stretch inside one paragraph */
export function rubyTargetOf(state: EditorState): RubyTarget | null {
  const { selection } = state
  if (selection instanceof NodeSelection) {
    const { node } = selection
    if (node.type.name !== 'docRuby') return null
    return {
      from: selection.from,
      to: selection.to,
      base: String(node.attrs.base ?? ''),
      rt: String(node.attrs.rt ?? ''),
      align: rubyAlignOf(String(node.attrs.xml ?? '')),
      reading: rubyReadingOf(String(node.attrs.xml ?? ''), sizeOf(node.marks)),
    }
  }
  const { from, to, $from, $to } = selection
  if (from === to || !$from.sameParent($to) || !$from.parent.isTextblock) return null
  let plain = true
  state.doc.nodesBetween(from, to, (node) => {
    if (node.isInline && !node.isText) plain = false
  })
  const base = state.doc.textBetween(from, to)
  if (!plain || !base.trim() || base.length > RUBY_BASE_MAX) return null
  const size = sizeOf(state.doc.nodeAt(from)?.marks ?? [])
  return { from, to, base, rt: '', align: null, reading: rubyReadingOf('', size) }
}

function sizeOf(marks: readonly Mark[]): number | undefined {
  const size = Number(marks.find((m) => m.type.name === 'docTextStyle')?.attrs.sizeHalfPoints)
  return size > 0 ? size : undefined
}

/** the guide as the dialog leaves it; a base or alignment left alone is not passed */
export interface RubyEdit {
  rt: string
  base?: string
  align?: RubyAlign
  reading?: Partial<RubyReading>
}

/** the kept fragment with the edit applied; empty when a part it must change is not there */
function editedRubyXml(kept: string, edit: RubyEdit): string {
  let xml: string | null = setRubyReadingXml(kept, edit.rt.trim())
  if (xml && edit.base !== undefined) xml = setRubyBaseXml(xml, edit.base)
  if (xml && edit.align) xml = setRubyAlignXml(xml, edit.align)
  if (xml && edit.reading) xml = withReading(xml, edit.reading)
  return xml ?? ''
}

function withReading(xml: string, reading: Partial<RubyReading>): string | null {
  const base = rubyReadingStyleOf(xml).baseSizeHalfPoints ?? DEFAULT_BASE_HALF_POINTS
  return setRubyReadingStyleXml(xml, {
    ...(reading.font === undefined ? {} : { font: reading.font }),
    ...(reading.sizePt === undefined ? {} : { sizeHalfPoints: Math.round(reading.sizePt * 2) }),
    ...(reading.offsetPt === undefined
      ? {}
      : { raiseHalfPoints: restRaise(base) + Math.round(reading.offsetPt * 2) }),
  })
}

/**
 * Puts the reading on the target, keeping the formatting of the base text. An
 * empty reading takes the guide off and leaves the base as plain text.
 */
export function applyRuby(
  editor: Editor,
  target: RubyTarget,
  edit: RubyEdit,
  lang: string,
): boolean {
  const { state } = editor
  const { schema } = state
  const at = state.doc.nodeAt(target.from)
  const marks = at?.marks ?? []
  const reading = edit.rt.trim()
  const base = edit.base?.trim() ? edit.base : target.base
  if (!reading) {
    if (!target.rt) return false
    return editor
      .chain()
      .focus()
      .command(({ tr }) => {
        tr.replaceWith(target.from, target.to, schema.text(base, marks))
        return true
      })
      .run()
  }
  // a guide that is already there keeps everything the dialog did not change
  const kept = at?.type.name === 'docRuby' ? String(at.attrs.xml ?? '') : ''
  const fresh = () => {
    const made = generateRubyXml(
      base,
      reading,
      sizeOf(marks),
      lang,
      edit.align ?? target.align ?? 'center',
    )
    return edit.reading ? (withReading(made, edit.reading) ?? made) : made
  }
  const xml =
    (kept && editedRubyXml(kept, { ...edit, base: base === target.base ? undefined : base })) ||
    fresh()
  const ruby = schema.nodes.docRuby.create({ base, rt: reading, xml }, null, marks)
  return editor
    .chain()
    .focus()
    .command(({ tr }) => {
      tr.replaceWith(target.from, target.to, ruby)
      // the guide stays selected, so the dialog can be opened on it again
      tr.setSelection(NodeSelection.create(tr.doc, target.from))
      return true
    })
    .run()
}
