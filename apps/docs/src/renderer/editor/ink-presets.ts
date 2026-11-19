import type { InkPenKind } from './ink'

/** one slot of the Draw tab's pen gallery (Word remembers a row of pens, each with its own look) */
export interface InkPreset {
  id: string
  kind: InkPenKind
  /** hex without '#' */
  color: string
  width: number
}

export const PEN_WIDTHS = [1, 2, 3.5, 5, 8]
export const HIGHLIGHTER_WIDTHS = [6, 10, 16, 22, 30]

export const PEN_COLORS = [
  '000000',
  '7F7F7F',
  'C00000',
  'FF0000',
  'FFC000',
  'FFFF00',
  '92D050',
  '00B050',
  '00B0F0',
  '0070C0',
  '002060',
  '7030A0',
  'FF66CC',
  '00CC99',
  'FF6600',
  'FFFFFF',
]

/** Word's highlighter grid: bright hues that stay readable under 45% alpha */
export const HIGHLIGHTER_COLORS = [
  'FFFF00',
  '00FF00',
  '00FFFF',
  'FF00FF',
  'FF8000',
  '92D050',
  '99D9FF',
  'FFB3D9',
  'E6004C',
  '00B050',
  '0070C0',
  'DDAAFF',
  'FF3300',
  'E6E6E6',
  '999999',
  '000000',
]

export const widthsFor = (kind: InkPenKind) =>
  kind === 'highlighter' ? HIGHLIGHTER_WIDTHS : PEN_WIDTHS
export const colorsFor = (kind: InkPenKind) =>
  kind === 'highlighter' ? HIGHLIGHTER_COLORS : PEN_COLORS

export const DEFAULT_INK_PRESETS: InkPreset[] = [
  { id: 'p1', kind: 'pen', color: '000000', width: 2 },
  { id: 'p2', kind: 'pen', color: 'C00000', width: 2 },
  { id: 'p3', kind: 'pen', color: '0070C0', width: 3.5 },
  { id: 'p4', kind: 'pencil', color: '000000', width: 3.5 },
  { id: 'p5', kind: 'highlighter', color: 'FFFF00', width: 10 },
]

const STORAGE_KEY = 'aidocs.inkPresets'

export function loadInkPresets(): InkPreset[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return DEFAULT_INK_PRESETS
    const list = JSON.parse(raw) as unknown
    if (!Array.isArray(list) || list.length === 0) return DEFAULT_INK_PRESETS
    const ok = list.filter(
      (p): p is InkPreset =>
        !!p &&
        typeof p === 'object' &&
        typeof (p as InkPreset).id === 'string' &&
        ['pen', 'highlighter', 'pencil'].includes((p as InkPreset).kind) &&
        /^[0-9A-Fa-f]{6}$/.test((p as InkPreset).color) &&
        typeof (p as InkPreset).width === 'number' &&
        (p as InkPreset).width > 0,
    )
    return ok.length > 0 ? ok : DEFAULT_INK_PRESETS
  } catch {
    return DEFAULT_INK_PRESETS
  }
}

export function saveInkPresets(presets: InkPreset[]): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(presets))
  } catch {
    // quota / private mode: the gallery simply resets next launch
  }
}

let presetSeq = 0
export const newInkPreset = (kind: InkPenKind): InkPreset => ({
  id: `p${Date.now()}-${presetSeq++}`,
  kind,
  color: kind === 'highlighter' ? 'FFFF00' : '000000',
  width: kind === 'highlighter' ? 10 : 2,
})
