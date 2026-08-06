/**
 * Shared WordArt gallery presets consumed by the docs and slides ribbons, so the
 * Insert → WordArt gallery is identical across apps (same recipes, same order,
 * same 4-column unlabeled letter grid — style names live in per-cell tooltips).
 *
 * 12 presets modeled on the PowerPoint WordArt gallery: solid bold / white text
 * with colored stroke / colored text with dark stroke / black-gold, metallic
 * silver, etc. (stroke widths in EMU: 12700 = 1pt, 19050 = 1.5pt).
 *
 * `nameKey` is the tooltip i18n key; both apps define the same keys in their
 * ribbon strings.
 */
export interface WordArtPreset {
  id: string
  nameKey: string
  fill: string
  outline?: { color: string; widthEmu: number }
  bold?: boolean
  italic?: boolean
}

export const WORDART_PRESETS: WordArtPreset[] = [
  // Solid bold series
  { id: 'blue', nameKey: 'ribbonWordArtPresetBlue', fill: '#4472C4', bold: true },
  { id: 'gold', nameKey: 'ribbonWordArtPresetGold', fill: '#FFC000', bold: true },
  { id: 'red', nameKey: 'ribbonWordArtPresetRed', fill: '#C00000', bold: true },
  { id: 'purple', nameKey: 'ribbonWordArtPresetPurple', fill: '#7030A0', bold: true },
  {
    id: 'green-italic',
    nameKey: 'ribbonWordArtPresetGreenItalic',
    fill: '#70AD47',
    bold: true,
    italic: true,
  },
  // White text + colored stroke (outline text effect)
  {
    id: 'white-orange',
    nameKey: 'ribbonWordArtPresetWhiteOrange',
    fill: '#FFFFFF',
    outline: { color: '#ED7D31', widthEmu: 19050 },
    bold: true,
  },
  {
    id: 'white-red',
    nameKey: 'ribbonWordArtPresetWhiteRed',
    fill: '#FFFFFF',
    outline: { color: '#C00000', widthEmu: 19050 },
    bold: true,
  },
  // Colored text + dark stroke (3D feel)
  {
    id: 'gold-brown',
    nameKey: 'ribbonWordArtPresetGoldBrown',
    fill: '#FFC000',
    outline: { color: '#7F5F00', widthEmu: 12700 },
    bold: true,
  },
  {
    id: 'sky-navy',
    nameKey: 'ribbonWordArtPresetSkyNavy',
    fill: '#00B0F0',
    outline: { color: '#1F4E79', widthEmu: 12700 },
    bold: true,
  },
  // Dark text + bright stroke
  {
    id: 'navy-white',
    nameKey: 'ribbonWordArtPresetNavyWhite',
    fill: '#1F3864',
    outline: { color: '#FFFFFF', widthEmu: 12700 },
    bold: true,
  },
  {
    id: 'black-gold',
    nameKey: 'ribbonWordArtPresetBlackGold',
    fill: '#0D0D0D',
    outline: { color: '#FFC000', widthEmu: 19050 },
    bold: true,
  },
  // Metallic silver
  {
    id: 'silver-dark',
    nameKey: 'ribbonWordArtPresetSilverDark',
    fill: '#D9D9D9',
    outline: { color: '#595959', widthEmu: 12700 },
    bold: true,
  },
]

/** Gallery preview stroke: EMU line width → CSS px (12700 EMU = 1pt = 4/3 px). */
export function wordArtStrokePx(widthEmu: number): number {
  return Math.round((widthEmu / 12700) * (4 / 3) * 100) / 100
}

/**
 * Solid-color approximation for targets that cannot express the stroke
 * (e.g. the docx run color): light fills would be invisible on a white page,
 * so fall back to the outline color instead.
 */
export function wordArtSolidColor(preset: WordArtPreset): string {
  if (!preset.outline) return preset.fill
  const hex = preset.fill.replace('#', '')
  const r = parseInt(hex.slice(0, 2), 16)
  const g = parseInt(hex.slice(2, 4), 16)
  const b = parseInt(hex.slice(4, 6), 16)
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
  return luminance > 0.8 ? preset.outline.color : preset.fill
}
