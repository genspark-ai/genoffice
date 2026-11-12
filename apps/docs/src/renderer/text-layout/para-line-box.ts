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
  STRUT_MIXED_CLASS,
  strutHalfPoints,
  type LayoutPara,
} from '../editor/para-layout-spec'
import { fontBoxCss, strutFontCss, textHasCjk } from '../line-metrics'

export interface TextLineBoxOptions {
  /** --doc-line-factor value; defaults to paraLineFactor (the live path passes the per-line strut) */
  factor?: string
  /** emit the declared face's --doc-font-box / --doc-font-skew (static emitters do, the live decoration does not yet) */
  fontBox?: boolean
  /** --doc-lead-gap:0px for CJK text (static emitters; the live paragraph gets it from its blockAttrs) */
  leadGapZero?: boolean
  /** apply the smallest-run strut size (notes keep their reservation size) */
  strut?: boolean
  /** .doc-ea-strut / .doc-mixed-face (the textbox path never added them) */
  struts?: boolean
}

export interface TextLineBox {
  styles: string[]
  classes: string[]
}

export function textParaLineBox(para: LayoutPara, opts: TextLineBoxOptions = {}): TextLineBox {
  const { fontBox = false, leadGapZero = false, strut: applyStrut = true, struts = true } = opts
  const factor = opts.factor ?? paraLineFactor(para)
  const styles = [`--doc-line-factor:${factor}`]
  const classes: string[] = []
  if (leadGapZero && textHasCjk(para.text)) styles.push('--doc-lead-gap:0px')
  const fam = paraDeclaredFontFamily(para)
  if (fam) {
    styles.push(`font-family:${fam}`, `--doc-grid-strut-tail:${fam}`)
    if (fontBox) styles.push(...fontBoxCss(fam))
  }
  const strutCls = fam
    ? struts
      ? (paraMixedFaceClass(para, factor) ?? declaredCjkStrutClass(para))
      : declaredCjkStrutClass(para)
    : struts
      ? eaStrutClass(para)
      : undefined
  if (strutCls) classes.push(strutCls)
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
