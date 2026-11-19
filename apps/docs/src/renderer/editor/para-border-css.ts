/** w:pBdr paragraph border CSS shared by the body paragraphs (para-layout-spec), the
 *  document stylesheet and the header/footer strips. */
import { borderCssStyle, borderDrawnPx } from './border-metrics'
import { darkPageColor } from './dark-page'

/** w:pBdr line CSS: declared color/width when present, legacy 1px #444 otherwise (document content color) */
export function paraBorderCss(line?: {
  color?: string
  szPt?: number
  val?: string
  frame?: true
}): string {
  const style = borderCssStyle(line?.val ?? 'single')
  const szEighths = line?.szPt ? line.szPt * 8 : undefined
  const drawn = borderDrawnPx({ style: line?.val ?? 'single', szEighths })
  // Word's 3-D setting bevels a plain line like a raised frame, as the page border does
  if (line?.frame && style === 'solid')
    return `${Math.max(3, drawn)}px ridge #${line.color ?? '444'}`
  // both lines of a double need a pixel each and one between them
  const widthPx = style === 'double' ? Math.max(3, drawn) : line?.szPt ? drawn : 1
  return `${widthPx}px ${style} #${line?.color ?? '444'}`
}

export type ParaBorderPadding = Partial<
  Record<'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft', string>
>

export type ParaBorderSideLines = Partial<
  Record<
    't' | 'b' | 'l' | 'r',
    {
      color?: string
      szPt?: number
      spacePt?: number
      none?: true
      val?: string
      shadow?: true
      frame?: true
    } | null
  >
>

/** the model's drawn lines plus the reset sides whose w:space still pads (borderPad) */
export function paraBorderLinesWithPad(
  lines?: ParaBorderSideLines | null,
  pad?: Partial<Record<'t' | 'b' | 'l' | 'r', number>> | null,
): ParaBorderSideLines {
  const out: ParaBorderSideLines = { ...lines }
  for (const ch of ['t', 'b', 'l', 'r'] as const) {
    if (pad?.[ch]) out[ch] = { none: true, spacePt: pad[ch] }
  }
  return out
}

/**
 * Padding for the drawn w:pBdr sides and the none sides that keep a w:space
 * (Word pads those too), as longhands so a direct side and a style-level side
 * on different edges layer per side (Word merges pBdr per side). Top/bottom:
 * the line sits w:space pt from the text and the gap is part of the paragraph
 * height; left/right sides never move the text (Word probe: x stays at the
 * margin, w:space only widens the border/shading box; paraBorderOutsetDecls).
 */
export function paraBorderPadding(drawn: string, lines?: ParaBorderSideLines): ParaBorderPadding {
  const out: ParaBorderPadding = {}
  const v = (ch: 't' | 'b') => {
    const space = lines?.[ch]?.spacePt
    return space ? `${space}pt` : '0'
  }
  const padded = (ch: 't' | 'b') => drawn.includes(ch) || lines?.[ch]?.none === true
  if (padded('t')) out.paddingTop = v('t')
  if (padded('b')) out.paddingBottom = v('b')
  return out
}

export type ParaBorderOutset = { l: number; r: number }

/** w:space (pt) of the drawn left/right sides: how far the line stands out from the text */
export function paraBorderOutset(
  drawn: string,
  lines?: ParaBorderSideLines,
): ParaBorderOutset | null {
  const of = (ch: 'l' | 'r') => {
    const line = lines?.[ch]
    return drawn.includes(ch) && line && !line.none && line.spacePt ? line.spacePt : 0
  }
  const out = { l: of('l'), r: of('r') }
  return out.l || out.r ? out : null
}

const pt = (n: number) => `${Math.round(n * 100) / 100}pt`

/**
 * The box grows outward by the outset (margin) and pads it back, so the text
 * column keeps its width. `styles` are the declarations already emitted: an
 * own indent there is the margin to move; otherwise the stylesheet's indent
 * arrives as `--pind-s/e`. List items pad their start through `--pbdr-os`.
 */
export function paraBorderOutsetDecls(
  styles: readonly string[],
  out: ParaBorderOutset,
  opts: { rtl?: boolean; list?: boolean; anchorHost?: boolean } = {},
): string[] {
  const decls: string[] = []
  const edges = [
    ['start', opts.rtl ? out.r : out.l],
    ['end', opts.rtl ? out.l : out.r],
  ] as const
  for (const [edge, space] of edges) {
    if (!space) continue
    const prop = `margin-inline-${edge}:`
    const own = styles.filter((d) => d.startsWith(prop)).pop()
    const ownPt = own ? parseFloat(own.slice(prop.length)) : NaN
    const k = edge[0]
    decls.push(
      Number.isFinite(ownPt)
        ? `${prop}${pt(ownPt - space)}`
        : `${prop}calc(var(--pind-${k}, 0pt) - ${pt(space)})`,
      `--pbdr-o${k}:${pt(space)}`,
    )
    if (!(opts.list && edge === 'start')) decls.push(`padding-inline-${edge}:${pt(space)}`)
  }
  // an anchored picture is placed from the box's physical left edge, in either direction
  if (opts.anchorHost && out.l) {
    const own = styles.filter((d) => d.startsWith('--anchor-ind:')).pop()
    decls.push(`--anchor-ind:${pt((own ? parseFloat(own.slice(13)) : 0) - out.l)}`)
  }
  return decls
}

/** the shading box extended by the none-side left/right w:space; the text never moves */
export const PARA_BORDER_SHADOW =
  'calc(-1 * var(--pbdr-l, 0pt)) 0 0 var(--pbdr-fill), var(--pbdr-r, 0pt) 0 0 var(--pbdr-fill)'

/** w:shadow: a hard drop shadow at the bottom-right; `--pbdr-drop-y` keeps it off a border group's seams */
export function paraBorderDropShadow(lines?: ParaBorderSideLines): string | null {
  const line = shadowLine(lines)
  if (!line) return null
  const px = Math.max(3, Math.round(((line.szPt ?? 0.5) * 96) / 72) * 2)
  return `${px}px var(--pbdr-drop-y, ${px}px) 0 var(--pbdr-drop-c, #${line.color ?? '444'})`
}

const shadowLine = (lines?: ParaBorderSideLines) =>
  (['b', 'r', 't', 'l'] as const).map((ch) => lines?.[ch]).find((l) => l && !l.none && l.shadow)

export function paraBorderDropColor(lines?: ParaBorderSideLines): string | null {
  const line = shadowLine(lines)
  return line ? `#${line.color ?? '444'}` : null
}

/** `--pbdr-l/r` for none left/right sides with a w:space (plus the shadow that paints them) */
export function paraBorderShadowDecls(lines?: ParaBorderSideLines): string[] {
  const out: string[] = []
  for (const ch of ['l', 'r'] as const) {
    const line = lines?.[ch]
    if (line?.none && line.spacePt) out.push(`--pbdr-${ch}:${line.spacePt}pt`)
  }
  const drop = paraBorderDropShadow(lines)
  // the inline twin the dark page reads; a stylesheet rule sets --pbdr-drop-c itself
  if (drop)
    out.push(`box-shadow:${drop}`, `--dk-drop:${darkPageColor(paraBorderDropColor(lines)!)}`)
  else if (out.length > 0) out.push(`box-shadow:${PARA_BORDER_SHADOW}`)
  return out
}

const PADDING_PROP = {
  paddingTop: 'padding-top',
  paddingRight: 'padding-right',
  paddingBottom: 'padding-bottom',
  paddingLeft: 'padding-left',
} as const

export function paraBorderPaddingDecls(padding: ParaBorderPadding): string[] {
  return (Object.keys(PADDING_PROP) as Array<keyof typeof PADDING_PROP>)
    .filter((k) => padding[k] !== undefined)
    .map((k) => `${PADDING_PROP[k]}:${padding[k]}`)
}
