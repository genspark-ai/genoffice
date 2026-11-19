/**
 * Built-in theme presets (Design tab theme gallery) — full color + font schemes.
 *
 * Each set = the 12 OOXML clrScheme slots + latin fonts for headings/body. On apply, the main
 * process rewrites theme*.xml in the package (scheme-referenced colors follow), and remaps the
 * deck's explicit srgbClr colors onto the new palette (neutrals
 * move along the dk1<->lt1 axis, chromatic colors map to accent1..6 by frequency, changing hue
 * only while keeping lightness) — so decks with explicit colors can also be reskinned in one click.
 *
 * Slot semantics: dk1=body text color, lt1=page background, dk2/lt2=secondary dark/light,
 * accent1..6=accent colors, hlink/folHlink=hyperlink colors.
 */

export interface SlideThemePreset {
  id: string
  name: string
  colors: Record<string, string>
  majorFont?: string
  minorFont?: string
}

const scheme = (
  dk1: string,
  lt1: string,
  dk2: string,
  lt2: string,
  accents: [string, string, string, string, string, string],
  hlink: string,
): Record<string, string> => ({
  dk1,
  lt1,
  dk2,
  lt2,
  accent1: accents[0],
  accent2: accents[1],
  accent3: accents[2],
  accent4: accents[3],
  accent5: accents[4],
  accent6: accents[5],
  hlink,
  folHlink: '954F72',
})

export const THEME_PRESETS: SlideThemePreset[] = [
  {
    id: 'office',
    name: 'Office',
    colors: scheme(
      '000000',
      'FFFFFF',
      '44546A',
      'E7E6E6',
      ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'],
      '0563C1',
    ),
    majorFont: 'Calibri Light',
    minorFont: 'Calibri',
  },
  {
    id: 'ember',
    name: 'Ember',
    colors: scheme(
      '2B1B14',
      'FFFFFF',
      '632B1A',
      'F7EBE6',
      ['C43E1C', 'E97132', 'FFC000', '8A3B12', 'D98F73', 'A33517'],
      'C43E1C',
    ),
    majorFont: 'Trebuchet MS',
    minorFont: 'Calibri',
  },
  {
    id: 'indigo',
    name: 'Indigo',
    colors: scheme(
      '1F2A44',
      'FFFFFF',
      '3B4C77',
      'E8ECF6',
      ['2E4FA3', '5B79C7', '8FA6DE', 'E97132', '31A5A0', '7030A0'],
      '2E4FA3',
    ),
    majorFont: 'Segoe UI',
    minorFont: 'Segoe UI',
  },
  {
    id: 'forest',
    name: 'Forest',
    colors: scheme(
      '1E2B20',
      'FFFFFF',
      '375E43',
      'E9F2EB',
      ['217346', '4EA72E', '92D050', 'FFC000', '3E8E8B', '70AD47'],
      '217346',
    ),
    majorFont: 'Candara',
    minorFont: 'Candara',
  },
  {
    id: 'cream',
    name: 'Cream',
    colors: scheme(
      '3B3226',
      'FBF6EC',
      '6E5F49',
      'F0E7D5',
      ['C0A062', '8A6D3B', 'B65C33', '6E8B5E', '4E7CA1', '9E5D74'],
      '8A6D3B',
    ),
    majorFont: 'Georgia',
    minorFont: 'Georgia',
  },
  {
    id: 'rose',
    name: 'Rose',
    colors: scheme(
      '3A1F33',
      'FFFFFF',
      '6E3A62',
      'F8ECF4',
      ['B44582', 'D96FA8', 'E9A6C9', '7030A0', 'E97132', '4E7CA1'],
      'B44582',
    ),
    majorFont: 'Trebuchet MS',
    minorFont: 'Calibri',
  },
  {
    id: 'graphite',
    name: 'Graphite',
    colors: scheme(
      'F2F2F2',
      '1E1E1E',
      'D9D9D9',
      '333333',
      ['4FC3F7', 'FFB74D', '81C784', 'BA68C8', 'E57373', '90A4AE'],
      '4FC3F7',
    ),
    majorFont: 'Segoe UI',
    minorFont: 'Segoe UI',
  },
  {
    id: 'midnight',
    name: 'Midnight',
    colors: scheme(
      'EAF2FF',
      '0F1C2E',
      'BCD0EC',
      '1E3350',
      ['4A9EDE', '63C7B2', 'F2C14E', 'E4718D', '9B8CDE', '7FB069'],
      '6FB3EC',
    ),
    majorFont: 'Calibri Light',
    minorFont: 'Calibri',
  },
]
