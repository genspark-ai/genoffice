import type { Editor } from '@tiptap/core'
import { resolveThemeColor, type ThemeColors } from '@genoffice/docx-engine'
import { TRACK_IGNORE } from './revisions'

const attrOf = (tag: string, name: string): string | undefined =>
  new RegExp(`\\bw:${name}="([^"]*)"`).exec(tag)?.[1]

/** what a run's own w:color theme reference resolves to on `colors`; null when it has none */
export function themeRunColor(rawRPr: string, colors: ThemeColors): string | null {
  const tag = /<w:color\b[^>]*>/.exec(rawRPr)?.[0]
  const ref = tag && attrOf(tag, 'themeColor')
  if (!tag || !ref) return null
  return resolveThemeColor(ref, colors, attrOf(tag, 'themeTint'), attrOf(tag, 'themeShade'))
}

/**
 * Body runs carry their theme colour resolved at parse; a theme pick repaints the ones
 * still on their reference. A run the user recoloured (color != themeColor) is left alone.
 */
export function retintThemeRuns(editor: Editor, colors: ThemeColors): void {
  const { state } = editor
  const type = state.schema.marks.docTextStyle
  if (!type) return
  const tr = state.tr
  state.doc.descendants((node, pos) => {
    if (!node.isText) return
    const mark = type.isInSet(node.marks)
    const { color, themeColor, rawRPr } = mark?.attrs ?? {}
    if (!mark || !themeColor || !rawRPr || color !== themeColor) return
    const next = themeRunColor(String(rawRPr), colors)
    if (!next || next === themeColor) return
    tr.addMark(
      pos,
      pos + node.nodeSize,
      type.create({ ...mark.attrs, color: next, themeColor: next }),
    )
  })
  if (!tr.docChanged) return
  editor.view.dispatch(tr.setMeta(TRACK_IGNORE, true).setMeta('addToHistory', false))
}
