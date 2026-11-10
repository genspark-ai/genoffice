import type { TextOutline } from '@genoffice/docx-engine'

/** CSS value for a w14:textOutline stroke (`-webkit-text-stroke`) */
export function textOutlineCssValue(o: TextOutline): string {
  const n = parseInt(o.color, 16)
  const color =
    o.alpha !== undefined && o.alpha < 1
      ? `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${o.alpha})`
      : `#${o.color}`
  return `${o.widthPt}pt ${color}`
}

export function textOutlineDecl(o: TextOutline): string {
  return `-webkit-text-stroke:${textOutlineCssValue(o)}`
}

export function parseTextOutlineAttr(raw: string): TextOutline | undefined {
  try {
    const o = JSON.parse(raw) as Partial<TextOutline>
    if (typeof o.color !== 'string' || typeof o.widthPt !== 'number') return undefined
    return {
      color: o.color,
      widthPt: o.widthPt,
      ...(o.alpha !== undefined ? { alpha: o.alpha } : {}),
    }
  } catch {
    return undefined
  }
}
