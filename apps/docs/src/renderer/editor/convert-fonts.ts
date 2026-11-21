import type { Editor } from '@tiptap/core'
import { replaceFontsXml } from '@genoffice/docx-engine'
import { FONT_CONVERSION_RECORDING, FontConversionStep } from './font-conversion-history'

const FONT_ATTRS = ['font', 'fontAscii', 'eastAsiaFont', 'csFont'] as const

/**
 * Convert Permanently for the text the editor holds: every run that names one
 * of the fonts (lower-case comparison) names its stand-in instead, as one undo
 * step. The parts the editor does not hold are rewritten at the save; the step
 * carries the fonts so an undo or redo reaches that share too. Returns whether
 * a run changed.
 */
export function convertFontsInEditor(
  editor: Editor,
  fonts: Readonly<Record<string, string>>,
): boolean {
  const byName = new Map(Object.entries(fonts).map(([from, to]) => [from.toLowerCase(), to]))
  const type = editor.schema.marks.docTextStyle
  if (!type || byName.size === 0) return false
  const { tr, doc } = editor.state
  doc.descendants((node, pos) => {
    if (!node.isText) return
    const mark = type.isInSet(node.marks)
    if (!mark) return
    const next: Record<string, string> = {}
    for (const attr of FONT_ATTRS) {
      const to = byName.get(String(mark.attrs[attr] ?? '').toLowerCase())
      if (to) next[attr] = to
    }
    if (Object.keys(next).length === 0) return
    // the run's own rPr is converted with it: saved against the old one, a
    // changed font would also be written into an eastAsia slot the run never had
    const raw = mark.attrs.rawRPr
    if (typeof raw === 'string') next.rawRPr = replaceFontsXml(raw, byName)
    tr.addMark(pos, pos + node.nodeSize, type.create({ ...mark.attrs, ...next }))
  })
  const runsChanged = tr.docChanged
  editor.view.dispatch(
    tr.step(new FontConversionStep(fonts, true)).setMeta(FONT_CONVERSION_RECORDING, true),
  )
  return runsChanged
}
