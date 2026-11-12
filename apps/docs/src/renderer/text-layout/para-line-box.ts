/**
 * The line box of a text paragraph: the strut factor, the declared strut face,
 * the smallest run size and the mixed-face / mixed-script classes. One
 * function for the live decoration, the static DOM of protected tables and
 * textboxes; the flags name what each emitter has always added on top (the
 * differences are recorded, not decided, here).
 */
import {
  declaredCjkStrutClass,
  eaStrutClass,
  paraCsMixedClass,
  paraCutAdjustedClass,
  paraDeclaredFontFamily,
  paraLineFactor,
  paraMixedFaceClass,
  paraStrutEaFace,
  STRUT_MIXED_CLASS,
  strutHalfPoints,
  type LayoutPara,
} from '../editor/para-layout-spec'
import { fontBoxCss, strutFontCss, textHasCjk } from '../line-metrics'
import { declaredStrutFace } from './strut-face'

export interface TextLineBoxOptions {
  /** --doc-line-factor value; defaults to paraLineFactor (the live path passes the per-line strut) */
  factor?: string
  /** apply the smallest-run strut size (notes keep their reservation size) */
  strut?: boolean
}

export interface TextLineBox {
  styles: string[]
  classes: string[]
}

export function textParaLineBox(para: LayoutPara, opts: TextLineBoxOptions = {}): TextLineBox {
  const { strut: applyStrut = true } = opts
  const factor = opts.factor ?? paraLineFactor(para)
  const styles = [`--doc-line-factor:${factor}`]
  const classes: string[] = []
  if (textHasCjk(para.text)) styles.push('--doc-lead-gap:0px')
  const fam = paraDeclaredFontFamily(para)
  if (fam) {
    // the declared face's box places the glyphs (Word probe 2026-11-05: a Times /
    // Georgia / Arial / Cambria paragraph under a Calibri Normal rests its own
    // hhea descent on the line bottom, the style's face and the mark do not count)
    styles.push(`font-family:${fam}`, `--doc-grid-strut-tail:${fam}`, ...fontBoxCss(fam))
  }
  const strutCls = fam
    ? (paraMixedFaceClass(para, factor) ?? declaredCjkStrutClass(para))
    : eaStrutClass(para)
  if (strutCls) classes.push(strutCls)
  if (strutCls === 'doc-grid-strut') {
    // the runs' own East Asian face leads the strut and the spans (strut-face.ts)
    const ea = paraStrutEaFace(para)
    const face = ea ? declaredStrutFace(ea) : null
    if (face) styles.push(`--doc-strut-face:'${face}'`)
  }
  const csMixed = paraCsMixedClass(para)
  if (csMixed) classes.push(csMixed)
  const cutAdjusted = paraCutAdjustedClass(para)
  if (cutAdjusted) classes.push(cutAdjusted)
  const strut = applyStrut ? strutHalfPoints(para) : null
  if (strut) {
    styles.push(...strutFontCss(strut.halfPoints, strut.mixed))
    if (strut.mixed) classes.push(STRUT_MIXED_CLASS)
  }
  return { styles, classes }
}
