import { STROKE_WIDTH } from './view-config'

export type DrawRgb = [number, number, number]

/** Pen / shape stroke widths offered on the Annotate tab (PDF pt) */
export const STROKE_WIDTHS = [1, 2, 3, 5, 8] as const

const COLOR_KEY = 'genoffice-pdf-draw-color'
const WIDTH_KEY = 'genoffice-pdf-draw-width'

const read = (key: string): string | null => {
  try {
    return localStorage.getItem(key)
  } catch {
    return null
  }
}

const write = (key: string, value: string): void => {
  try {
    localStorage.setItem(key, value)
  } catch {
    // storage unavailable: the choice lasts for this session only
  }
}

export function loadDrawColor(fallback: DrawRgb): DrawRgb {
  try {
    const parsed: unknown = JSON.parse(read(COLOR_KEY) ?? 'null')
    if (
      Array.isArray(parsed) &&
      parsed.length === 3 &&
      parsed.every((v) => typeof v === 'number' && v >= 0 && v <= 1)
    )
      return parsed as DrawRgb
  } catch {
    // malformed entry: fall through
  }
  return fallback
}

export const saveDrawColor = (color: DrawRgb): void => write(COLOR_KEY, JSON.stringify(color))

export function loadDrawWidth(): number {
  const saved = Number(read(WIDTH_KEY))
  return (STROKE_WIDTHS as readonly number[]).includes(saved) ? saved : STROKE_WIDTH
}

export const saveDrawWidth = (width: number): void => write(WIDTH_KEY, String(width))
