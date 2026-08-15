/** The one preset palette shared by every color picker in the PDF app
    (text insert, text edit, draw color): 12 quick picks in a 6-column grid,
    same layout as the highlight presets — anything else comes from the
    gradient / hex field. */
export const COLOR_PRESETS = [
  '#000000',
  '#808080',
  '#FFFFFF',
  '#E53935',
  '#FB8C00',
  '#FDD835',
  '#7CB342',
  '#22A75A',
  '#00ACC1',
  '#2B66FF',
  '#7E57C2',
  '#D81B60',
] as const

export const isHexColor = (value: string): boolean => /^#[0-9a-f]{6}$/i.test(value)

export const hexTo255 = (hex: string): [number, number, number] => [
  parseInt(hex.slice(1, 3), 16),
  parseInt(hex.slice(3, 5), 16),
  parseInt(hex.slice(5, 7), 16),
]

export const rgb255ToHex = (c: readonly [number, number, number]): string =>
  `#${c.map((v) => v.toString(16).padStart(2, '0')).join('')}`

/** 0-255 RGB → [hue 0-360, saturation 0-1, value 0-1] */
export const rgbToHsv = (r: number, g: number, b: number): [number, number, number] => {
  const rn = r / 255
  const gn = g / 255
  const bn = b / 255
  const max = Math.max(rn, gn, bn)
  const min = Math.min(rn, gn, bn)
  const d = max - min
  let h = 0
  if (d > 0) {
    if (max === rn) h = 60 * (((gn - bn) / d + 6) % 6)
    else if (max === gn) h = 60 * ((bn - rn) / d + 2)
    else h = 60 * ((rn - gn) / d + 4)
  }
  return [h, max === 0 ? 0 : d / max, max]
}

/** [hue 0-360, saturation 0-1, value 0-1] → 0-255 RGB */
export const hsvToRgb = (h: number, s: number, v: number): [number, number, number] => {
  const f = (n: number): number => {
    const k = (n + h / 60) % 6
    return v - v * s * Math.max(0, Math.min(k, 4 - k, 1))
  }
  return [Math.round(f(5) * 255), Math.round(f(3) * 255), Math.round(f(1) * 255)]
}
