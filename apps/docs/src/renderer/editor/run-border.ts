import type { Run } from '@genoffice/docx-engine'
import { dkBorder } from './dark-page'

type RunBorder = NonNullable<Run['bdr']>

const BORDER_STYLES: Record<string, string> = {
  dotted: 'dotted',
  dashed: 'dashed',
  dashSmallGap: 'dashed',
  dotDash: 'dashed',
  dotDotDash: 'dashed',
  double: 'double',
  triple: 'double',
  thinThickSmallGap: 'double',
  thickThinSmallGap: 'double',
  thinThickMediumGap: 'double',
  thickThinMediumGap: 'double',
  inset: 'inset',
  outset: 'outset',
  threeDEmboss: 'ridge',
  threeDEngrave: 'groove',
}

export function parseRunBorderAttr(json: string): RunBorder | undefined {
  try {
    const v = JSON.parse(json) as RunBorder
    return v && typeof v.val === 'string' && typeof v.sz === 'number' ? v : undefined
  } catch {
    return undefined
  }
}

/** `0.75pt solid #000000` for a character border (w:sz is eighths of a point; Word draws at least 1/4pt) */
export function runBorderCss(bdr: RunBorder): string {
  const style = BORDER_STYLES[bdr.val] ?? 'solid'
  const color = bdr.color ? `#${bdr.color}` : 'currentColor'
  return `${Math.max(bdr.sz, 2) / 8}pt ${style} ${color}`
}

/** inline declarations for the docTextStyle span: border, the w:space text gap, dark-page twins */
export function runBorderDecls(json: string): string[] {
  const bdr = parseRunBorderAttr(json)
  if (!bdr) return []
  const border = runBorderCss(bdr)
  const decls = [`border:${border}`]
  if (bdr.space) decls.push(`padding:${bdr.space}pt`)
  if (bdr.color) decls.push(...(['t', 'r', 'b', 'l'] as const).map((s) => dkBorder(s, border)))
  return decls
}
