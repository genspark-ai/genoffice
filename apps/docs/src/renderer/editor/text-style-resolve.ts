import type { Editor } from '@tiptap/core'
import type { DocDefaults, StyleInfo } from '@genoffice/docx-engine'

const TEXT_BLOCK_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

function styleSize(styleId: unknown, styles: Map<string, StyleInfo> | undefined): number | null {
  if (typeof styleId !== 'string' || !styleId) return null
  const size = styles?.get(styleId)?.display?.sizeHalfPoints
  return typeof size === 'number' && size > 0 ? size : null
}

/** Resolve the font size Word displays at the caret without adding direct formatting. */
export function effectiveSizeHalfPoints(
  editor: Editor,
  styles?: Map<string, StyleInfo>,
  docDefaults?: DocDefaults,
): number | null {
  const textAttrs = editor.getAttributes('docTextStyle')
  const directSize = Number(textAttrs.sizeHalfPoints)
  if (Number.isFinite(directSize) && directSize > 0) return directSize

  const characterStyleSize = styleSize(textAttrs.styleId, styles)
  if (characterStyleSize !== null) return characterStyleSize

  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    const node = $from.node(depth)
    if (!TEXT_BLOCK_TYPES.has(node.type.name)) continue
    const paragraphStyleSize = styleSize(node.attrs.styleId, styles)
    if (paragraphStyleSize !== null) return paragraphStyleSize
    break
  }

  const defaultSize = docDefaults?.sizeHalfPoints
  return typeof defaultSize === 'number' && defaultSize > 0 ? defaultSize : null
}

function runSizeHalfPoints(
  attrs: Record<string, unknown>,
  paraStyleId: unknown,
  styles: Map<string, StyleInfo> | undefined,
  defaults: DocDefaults | undefined,
): number | null {
  const direct = Number(attrs.sizeHalfPoints)
  if (Number.isFinite(direct) && direct > 0) return direct
  const styled = styleSize(attrs.styleId, styles) ?? styleSize(paraStyleId, styles)
  if (styled !== null) return styled
  const d = defaults?.sizeHalfPoints
  return typeof d === 'number' && d > 0 ? d : null
}

/** Word blanks the size box when the selection's runs resolve to different sizes */
export function selectedSizeMixed(
  editor: Editor,
  styles?: Map<string, StyleInfo>,
  defaults?: DocDefaults,
): boolean {
  const { from, to, empty } = editor.state.selection
  if (empty) return false
  let seen: number | null | undefined
  let mixed = false
  editor.state.doc.nodesBetween(from, to, (node, _pos, parent) => {
    if (mixed) return false
    let size: number | null
    if (node.isText)
      size = runSizeHalfPoints(
        node.marks.find((m) => m.type.name === 'docTextStyle')?.attrs ?? {},
        parent?.attrs.styleId,
        styles,
        defaults,
      )
    else if (node.isTextblock && node.childCount === 0)
      size = runSizeHalfPoints({}, node.attrs.styleId, styles, defaults)
    else return true
    if (seen === undefined) seen = size
    else if (seen !== size) mixed = true
    return true
  })
  return mixed
}

function defaultParagraphStyle(styles: Map<string, StyleInfo> | undefined) {
  return (
    [...(styles?.values() ?? [])].find((style) => style.type === 'paragraph' && style.isDefault) ??
    styles?.get('Normal')
  )
}

function styleEastAsiaFont(display: StyleInfo['display']): string | undefined {
  return (
    display?.eastAsiaFont ??
    (display?.eaSlotEmpty || display?.font === display?.fontAscii ? undefined : display?.font)
  )
}

export function selectedFonts(
  editor: Editor,
  styles?: Map<string, StyleInfo>,
  defaults?: DocDefaults,
) {
  const defaultPara = defaultParagraphStyle(styles)
  let result: { fontEastAsia: string | null; fontLatin: string | null } | undefined
  const add = (attrs: Record<string, unknown>, styleId: unknown) => {
    const char = styles?.get(String(attrs.styleId ?? ''))?.display
    const para = (styles?.get(String(styleId ?? '')) ?? defaultPara)?.display
    const ea =
      typeof attrs.eastAsiaFont === 'string'
        ? attrs.eastAsiaFont
        : typeof attrs.font === 'string' &&
            !attrs.eaSlotEmpty &&
            !attrs.rawRPr &&
            attrs.font !== attrs.fontAscii
          ? attrs.font
          : undefined
    const next = {
      fontEastAsia:
        ea ?? styleEastAsiaFont(char) ?? styleEastAsiaFont(para) ?? defaults?.eastAsiaFont ?? '',
      fontLatin:
        typeof attrs.fontAscii === 'string'
          ? attrs.fontAscii
          : (char?.fontAscii ?? para?.fontAscii ?? defaults?.asciiFont ?? ''),
    }
    if (!result) result = next
    else {
      if (result.fontEastAsia !== next.fontEastAsia) result.fontEastAsia = null
      if (result.fontLatin !== next.fontLatin) result.fontLatin = null
    }
  }
  const { from, to, $from, empty } = editor.state.selection
  if (empty) add(editor.getAttributes('docTextStyle'), $from.parent.attrs.styleId)
  else
    editor.state.doc.nodesBetween(from, to, (node, _pos, parent) => {
      if (node.isText)
        add(
          node.marks.find((m) => m.type.name === 'docTextStyle')?.attrs ?? {},
          parent?.attrs.styleId,
        )
      else if (node.isTextblock && node.childCount === 0) add({}, node.attrs.styleId)
    })
  return result ?? { fontEastAsia: '', fontLatin: '' }
}
