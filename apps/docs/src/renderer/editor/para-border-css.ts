/** w:pBdr paragraph border CSS shared by the body paragraphs (para-layout-spec), the
 *  document stylesheet and the header/footer strips. */
/** w:pBdr line CSS: declared color/width when present, legacy 1px #444 otherwise (document content color) */
export function paraBorderCss(line?: { color?: string; szPt?: number }): string {
  const widthPx = line?.szPt ? Math.max(1, Math.round((line.szPt * 96) / 72)) : 1
  return `${widthPx}px solid #${line?.color ?? '444'}`
}

export type ParaBorderPadding = Partial<
  Record<'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft', string>
>

export type ParaBorderSideLines = Partial<
  Record<
    't' | 'b' | 'l' | 'r',
    { color?: string; szPt?: number; spacePt?: number; none?: true } | null
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
 * margin, w:space only widens the border/shading box).
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

/** the shading box extended by the none-side left/right w:space; the text never moves */
export const PARA_BORDER_SHADOW =
  'calc(-1 * var(--pbdr-l, 0pt)) 0 0 var(--pbdr-fill), var(--pbdr-r, 0pt) 0 0 var(--pbdr-fill)'

/** `--pbdr-l/r` for none left/right sides with a w:space (plus the shadow that paints them) */
export function paraBorderShadowDecls(lines?: ParaBorderSideLines): string[] {
  const out: string[] = []
  for (const ch of ['l', 'r'] as const) {
    const line = lines?.[ch]
    if (line?.none && line.spacePt) out.push(`--pbdr-${ch}:${line.spacePt}pt`)
  }
  if (out.length > 0) out.push(`box-shadow:${PARA_BORDER_SHADOW}`)
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
