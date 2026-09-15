/**
 * Content-control checkboxes (w14:checkbox). Word draws the box from the
 * control's own state: w14:checked picks between the checkedState and
 * uncheckedState glyphs, whose values are hex code points. The run inside
 * sdtContent holds the same glyph, so in the editor the glyph is the state and
 * write-back derives w14:checked from it, the way FORMCHECKBOX does.
 */

import { attrsOf, findChild, serializeXNode, type XNode } from './xml-utils'

export interface CheckboxGlyphs {
  readonly checked: string
  readonly unchecked: string
}

const DEFAULT_GLYPHS: CheckboxGlyphs = { checked: '☒', unchecked: '☐' }

const glyphOf = (hex: string | undefined, fallback: string): string => {
  const code = parseInt(hex ?? '', 16)
  return Number.isFinite(code) && code > 0 && code <= 0x10ffff
    ? String.fromCodePoint(code)
    : fallback
}

/** Checkbox state values are case-insensitive in the wild (TRUE/ON/True):
    an absent value means checked, mirroring the on/off falsy-list parity. */
export const isOn = (val: string | undefined): boolean => {
  if (val === undefined) return true
  const lower = val.toLowerCase()
  return lower === '1' || lower === 'true' || lower === 'on'
}

/** The checkbox glyph pair a w:sdtPr declares (defaults are Word's own). */
export function sdtCheckboxGlyphs(sdtPrXml: string): CheckboxGlyphs {
  const state = (name: string): string | undefined => {
    const m = new RegExp(`<w14:${name}\\b[^>]*\\bw14:val=(?:"([^"]*)"|'([^']*)')`).exec(sdtPrXml)
    return m?.[1] ?? m?.[2]
  }
  return {
    checked: glyphOf(state('checkedState'), DEFAULT_GLYPHS.checked),
    unchecked: glyphOf(state('uncheckedState'), DEFAULT_GLYPHS.unchecked),
  }
}

/** A w:sdt element that is a checkbox control, read for the run model. */
export function sdtCheckboxControl(
  sdt: XNode,
): { readonly sdtPrXml: string; readonly glyph: string } | null {
  const sdtPr = findChild(sdt, 'w:sdtPr')
  const box = sdtPr ? findChild(sdtPr, 'w14:checkbox') : undefined
  if (!sdtPr || !box) return null
  const checkedEl = findChild(box, 'w14:checked')
  const checked = checkedEl ? isOn(attrsOf(checkedEl)['w14:val']) : false
  const sdtPrXml = serializeXNode(sdtPr)
  const glyphs = sdtCheckboxGlyphs(sdtPrXml)
  return { sdtPrXml, glyph: checked ? glyphs.checked : glyphs.unchecked }
}

/** Whether `glyph` reads as the checked state of a control declaring `sdtPrXml`. */
export function sdtCheckboxIsChecked(sdtPrXml: string, glyph: string): boolean {
  const glyphs = sdtCheckboxGlyphs(sdtPrXml)
  if (glyph === glyphs.checked) return true
  if (glyph === glyphs.unchecked) return false
  return /[☑☒✓✔✅]/u.test(glyph)
}

/** The w:sdtPr with w14:checked set to `checked`. */
export function syncSdtCheckbox(sdtPrXml: string, checked: boolean): string {
  const val = `<w14:checked w14:val="${checked ? '1' : '0'}"/>`
  const stripped = sdtPrXml.replace(/<w14:checked(?:\s[^>]*)?\/>/g, '')
  if (/<w14:checkbox(?:\s[^>]*)?\/>/.test(stripped))
    return stripped.replace(
      /<w14:checkbox((?:\s[^>]*)?)\/>/,
      `<w14:checkbox$1>${val}</w14:checkbox>`,
    )
  return stripped.replace(/<w14:checkbox(?:\s[^>]*)?>/, (open) => `${open}${val}`)
}
