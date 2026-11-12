/**
 * Inline CSS for the non-editable surfaces (field / formula / picture
 * wrappers, textbox paragraphs) from the engine's resolved paragraph and run
 * values. These surfaces have no style cascade worth keeping, so one helper
 * turns the resolved twips into declarations instead of five hand-rolled
 * margin/indent blocks.
 */
import {
  additiveParagraphSpacing,
  autoSpacingPt,
  cssFontFamily,
  cssGridSpacingPt,
  spaceBeforeProp,
} from '../line-metrics'

export interface EffectiveParaCss {
  spaceBeforeTwips?: number | null
  spaceAfterTwips?: number | null
  spaceBeforeAuto?: boolean | null
  spaceAfterAuto?: boolean | null
  indentLeftTwips?: number | null
  indentRightTwips?: number | null
  indentFirstLineTwips?: number | null
  /** settings.xml doNotUseHTMLParagraphAutoSpacing; absent = the document's loaded setting */
  additiveSpacing?: boolean | null
}

export interface EffectiveParaCssOptions {
  /** snap spacing to 0.1pt (cssGridSpacingPt); off keeps the raw twips/20 */
  snap?: boolean
  /** emit margin-inline-start/end (text-indent is always emitted when given) */
  margins?: boolean
  /** length formatter for every twips value (default: points, spacing snapped per `snap`) */
  length?: (twips: number) => string
  /** the left indent lands on padding (header/footer cells: tab stops measure from the cell edge) */
  indentLeftAsPadding?: boolean
  /** which lines of a multi-line box this is: first-line props (space before,
   *  first-line indent) and last-line props (space after) land where they apply */
  line?: { first: boolean; last: boolean }
}

/** camelCase style record form of effectivePPrCss */
export function effectivePPrStyle(
  eff: EffectiveParaCss,
  opts: EffectiveParaCssOptions = {},
): Record<string, string> {
  const additive = eff.additiveSpacing ?? additiveParagraphSpacing()
  const pt = (v: number) => (opts.snap === false ? `${v}pt` : cssGridSpacingPt(v))
  const spacing = opts.length ?? ((twips: number) => pt(twips / 20))
  const indent = opts.length ?? ((twips: number) => `${twips / 20}pt`)
  const before = spaceBeforeProp(additive) === 'padding-top' ? 'paddingTop' : 'marginTop'
  const first = opts.line?.first ?? true
  const last = opts.line?.last ?? true
  const out: Record<string, string> = {}
  if (first) {
    if (eff.spaceBeforeAuto) out[before] = spacing(autoSpacingPt(additive) * 20)
    else if (eff.spaceBeforeTwips != null) out[before] = spacing(Number(eff.spaceBeforeTwips))
  }
  if (last) {
    if (eff.spaceAfterAuto) out.marginBottom = spacing(autoSpacingPt(additive) * 20)
    else if (eff.spaceAfterTwips != null) out.marginBottom = spacing(Number(eff.spaceAfterTwips))
  }
  if (opts.margins !== false) {
    // logical sides: w:ind left/right swap in bidi paragraphs (Word)
    if (eff.indentLeftTwips != null) {
      out[opts.indentLeftAsPadding ? 'paddingLeft' : 'marginInlineStart'] = indent(
        Number(eff.indentLeftTwips),
      )
    }
    if (eff.indentRightTwips != null) out.marginInlineEnd = indent(Number(eff.indentRightTwips))
  }
  if (first && eff.indentFirstLineTwips != null)
    out.textIndent = indent(Number(eff.indentFirstLineTwips))
  return out
}

/** spacing (before as padding in additive mode, auto = Word's 14pt) and indents */
export function effectivePPrCss(
  eff: EffectiveParaCss,
  opts: EffectiveParaCssOptions = {},
): string[] {
  return Object.entries(effectivePPrStyle(eff, opts)).map(([k, v]) => `${kebab(k)}:${v}`)
}

function kebab(prop: string): string {
  return prop.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`)
}

export interface EffectiveRunCss {
  sizeHalfPoints?: number | null
  fontFamily?: string | null
  bold?: boolean | null
}

export interface EffectiveRunCssOptions {
  /** font-weight for bold (header/footer strips paint 600) */
  boldWeight?: string
  /** font-family already mapped to CSS (two-slot run faces); default cssFontFamily */
  fontFamilyCss?: (font: string) => string
}

/** camelCase style record form of effectiveRPrCss */
export function effectiveRPrStyle(
  eff: EffectiveRunCss,
  opts: EffectiveRunCssOptions = {},
): Record<string, string> {
  const out: Record<string, string> = {}
  if (eff.sizeHalfPoints) out.fontSize = `${Number(eff.sizeHalfPoints) / 2}pt`
  if (eff.fontFamily) out.fontFamily = (opts.fontFamilyCss ?? cssFontFamily)(eff.fontFamily)
  if (eff.bold) out.fontWeight = opts.boldWeight ?? 'bold'
  return out
}

/** size / face / weight of a resolved run */
export function effectiveRPrCss(eff: EffectiveRunCss, opts?: EffectiveRunCssOptions): string[] {
  return Object.entries(effectiveRPrStyle(eff, opts)).map(([k, v]) => `${kebab(k)}:${v}`)
}
