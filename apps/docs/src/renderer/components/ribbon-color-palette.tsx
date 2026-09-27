import { ColorPicker } from '@genoffice/ui'
import { useI18n, type StringKey } from '../i18n/locale'

const THEME_COLORS: Array<{ nameKey: StringKey; hex: string }> = [
  { nameKey: 'ribbonColorWhite', hex: 'FFFFFF' },
  { nameKey: 'ribbonColorBlack', hex: '000000' },
  { nameKey: 'ribbonColorLightGray', hex: 'E7E6E6' },
  { nameKey: 'ribbonColorBlueGray', hex: '0E2841' },
  { nameKey: 'ribbonColorBlue', hex: '156082' },
  { nameKey: 'ribbonColorOrange', hex: 'E97132' },
  { nameKey: 'ribbonColorGreen', hex: '196B24' },
  { nameKey: 'ribbonColorSkyBlue', hex: '0F9ED5' },
  { nameKey: 'ribbonColorPurple', hex: 'A02B93' },
  { nameKey: 'ribbonColorLightGreenAlt', hex: '4EA72E' },
]

/** Word standard colors */
const STANDARD_COLORS: Array<{ nameKey: StringKey; hex: string }> = [
  { nameKey: 'ribbonColorDarkRed', hex: 'C00000' },
  { nameKey: 'ribbonColorRed', hex: 'FF0000' },
  { nameKey: 'ribbonColorOrange', hex: 'FFC000' },
  { nameKey: 'ribbonColorYellow', hex: 'FFFF00' },
  { nameKey: 'ribbonColorLightGreen', hex: '92D050' },
  { nameKey: 'ribbonColorGreen', hex: '00B050' },
  { nameKey: 'ribbonColorLightBlue', hex: '00B0F0' },
  { nameKey: 'ribbonColorBlue', hex: '0070C0' },
  { nameKey: 'ribbonColorDarkBlue', hex: '002060' },
  { nameKey: 'ribbonColorPurple', hex: '7030A0' },
]

/** Translated tooltip names for the shared picker's named swatches */
const COLOR_NAME_KEYS: Record<string, StringKey> = Object.fromEntries(
  [...THEME_COLORS, ...STANDARD_COLORS].map((c) => [c.hex, c.nameKey]),
)

/** Word-style theme + standard color palette (shared panel, docs anchor positioning) */
export function RibbonColorPalette({
  current,
  noneLabel,
  onPick,
}: {
  current: string | null
  noneLabel: string
  onPick: (hex: string | null) => void
}) {
  const { t } = useI18n()
  // data-rb-panel marks the picker as "inside" for the unified dismissal
  // guard; display:contents keeps the wrapper out of layout so the panel's
  // anchor positioning still resolves against the trigger wrap.
  return (
    <div data-rb-panel="" style={{ display: 'contents' }}>
      <ColorPicker
        className="docs-color-pop"
        value={current ? `#${current}` : null}
        strings={{
          auto: noneLabel,
          themeColors: t('ribbonThemeColorsSection'),
          standardColors: t('ribbonStandardColors'),
          moreColors: t('ribbonMoreColors'),
          shadeTip: (r, c) => t('ribbonThemeColorShadeTip', { r, c }),
          colorName: (s) => {
            const key = COLOR_NAME_KEYS[s.hex]
            return key ? t(key) : s.name
          },
        }}
        onPick={(hex) => onPick(hex ? hex.slice(1) : null)}
      />
    </div>
  )
}
