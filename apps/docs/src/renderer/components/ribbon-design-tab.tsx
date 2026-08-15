import { useState } from 'react'
import { type SectionSettings, type ThemeColors, type ThemeFonts } from '@genoffice/docx-engine'
import { PromptModal } from './PromptModal'
import { useI18n, type StringKey } from '../i18n/locale'
import {
  IconCaret,
  IconPageBorders,
  IconPageColor,
  IconTheme,
  IconThemeColors,
  IconThemeFonts,
  IconWatermark,
} from './icons'

/** icon size for the big icon-over-label ribbon buttons (slides ribbon parity) */
import { BIG, TabProps, toggleDropdown } from './ribbon-tabs'

const PAGE_COLORS: Array<{ nameKey: StringKey; hex: string | null }> = [
  { nameKey: 'ribbonColorWhite', hex: null },
  { nameKey: 'ribbonColorLightYellow', hex: 'FFF9E6' },
  { nameKey: 'ribbonColorBeige', hex: 'F5F0E6' },
  { nameKey: 'ribbonColorLightGreen', hex: 'EAF5EA' },
  { nameKey: 'ribbonColorLightBlue', hex: 'E8F1FB' },
  { nameKey: 'ribbonColorLightPurple', hex: 'F3EEFB' },
  { nameKey: 'ribbonColorLightGray', hex: 'F2F2F2' },
  { nameKey: 'ribbonColorDarkGray', hex: '333333' },
  { nameKey: 'ribbonColorBlack', hex: '000000' },
]

/** Word-like theme presets: font pair + color scheme applied together */
const THEME_PRESETS: Array<{ nameKey: StringKey; fonts: ThemeFonts; colors: ThemeColors }> = [
  {
    nameKey: 'ribbonThemeOffice',
    fonts: { major: 'Calibri Light', minor: 'Calibri', eastAsia: '等线' },
    colors: {
      name: 'Office',
      dk2: '44546A',
      lt2: 'E7E6E6',
      accent1: '4472C4',
      accent2: 'ED7D31',
      accent3: 'A5A5A5',
      accent4: 'FFC000',
      accent5: '5B9BD5',
      accent6: '70AD47',
    },
  },
  {
    nameKey: 'ribbonThemeFacet',
    fonts: { major: 'Trebuchet MS', minor: 'Trebuchet MS', eastAsia: '微软雅黑' },
    colors: {
      name: 'Facet',
      dk2: '3E3D2D',
      lt2: 'E1DFDD',
      accent1: '90C226',
      accent2: '54A021',
      accent3: 'E6B91E',
      accent4: 'E76618',
      accent5: 'C42F1A',
      accent6: '918655',
    },
  },
  {
    nameKey: 'ribbonThemeIon',
    fonts: { major: 'Century Gothic', minor: 'Century Gothic', eastAsia: '黑体' },
    colors: {
      name: 'Ion',
      dk2: '1B587C',
      lt2: 'E8ECEE',
      accent1: '4E8542',
      accent2: '604878',
      accent3: 'C19859',
      accent4: 'CA7B4F',
      accent5: '415D6B',
      accent6: '8EA5AF',
    },
  },
  {
    nameKey: 'ribbonThemeRetrospect',
    fonts: { major: 'Calibri Light', minor: 'Calibri', eastAsia: '仿宋' },
    colors: {
      name: 'Retrospect',
      dk2: '637052',
      lt2: 'CCDDEA',
      accent1: 'E48312',
      accent2: 'BD582C',
      accent3: '865640',
      accent4: '9B8357',
      accent5: 'C2BC80',
      accent6: '94A088',
    },
  },
]

const THEME_FONT_PRESETS: Array<{ nameKey: StringKey; fonts: ThemeFonts }> = [
  {
    nameKey: 'ribbonThemeFontOffice',
    fonts: { major: 'Calibri Light', minor: 'Calibri', eastAsia: '等线' },
  },
  { nameKey: 'ribbonThemeFontSongti', fonts: { major: '黑体', minor: '宋体', eastAsia: '宋体' } },
  {
    nameKey: 'ribbonThemeFontYahei',
    fonts: { major: '微软雅黑', minor: '微软雅黑', eastAsia: '微软雅黑' },
  },
  { nameKey: 'ribbonThemeFontKaiti', fonts: { major: '楷体', minor: '楷体', eastAsia: '楷体' } },
  {
    nameKey: 'ribbonThemeFontTimes',
    fonts: { major: 'Times New Roman', minor: 'Times New Roman', eastAsia: '宋体' },
  },
  { nameKey: 'ribbonThemeFontArial', fonts: { major: 'Arial', minor: 'Arial', eastAsia: '黑体' } },
]

const THEME_COLOR_PRESETS: Array<{ nameKey: StringKey; colors: ThemeColors }> = [
  {
    nameKey: 'ribbonThemeOffice',
    colors: {
      name: 'Office',
      dk2: '44546A',
      lt2: 'E7E6E6',
      accent1: '4472C4',
      accent2: 'ED7D31',
      accent3: 'A5A5A5',
      accent4: 'FFC000',
      accent5: '5B9BD5',
      accent6: '70AD47',
    },
  },
  {
    nameKey: 'ribbonPaletteGrayscale',
    colors: {
      name: 'Grayscale',
      dk2: '3F3F3F',
      lt2: 'B3B3B3',
      accent1: '5B5B5B',
      accent2: '8C8C8C',
      accent3: 'A6A6A6',
      accent4: '737373',
      accent5: '404040',
      accent6: 'D9D9D9',
    },
  },
  {
    nameKey: 'ribbonColorBlue',
    colors: {
      name: 'Blue',
      dk2: '17406D',
      lt2: 'DBEFF9',
      accent1: '0F6FC6',
      accent2: '009DD9',
      accent3: '0BD0D9',
      accent4: '10CF9B',
      accent5: '7CCA62',
      accent6: 'A5C249',
    },
  },
  {
    nameKey: 'ribbonColorGreen',
    colors: {
      name: 'Green',
      dk2: '455F51',
      lt2: 'E3DED1',
      accent1: '549E39',
      accent2: '8AB833',
      accent3: 'C0CF3A',
      accent4: '029676',
      accent5: '4AB5C4',
      accent6: '0989B1',
    },
  },
  {
    nameKey: 'ribbonPaletteRedViolet',
    colors: {
      name: 'Red Violet',
      dk2: '5F5F5F',
      lt2: 'FFFDF7',
      accent1: 'A83D4C',
      accent2: 'D2554B',
      accent3: 'E1734C',
      accent4: '8D3B4B',
      accent5: '54387F',
      accent6: '7A99AC',
    },
  },
  {
    nameKey: 'ribbonColorOrange',
    colors: {
      name: 'Orange',
      dk2: '637052',
      lt2: 'CCDDEA',
      accent1: 'E48312',
      accent2: 'BD582C',
      accent3: '865640',
      accent4: '9B8357',
      accent5: 'C2BC80',
      accent6: '94A088',
    },
  },
]

const WATERMARK_KEYS: StringKey[] = [
  'ribbonWatermarkConfidential',
  'ribbonWatermarkNoCopy',
  'ribbonWatermarkDraft',
  'ribbonWatermarkSample',
]

interface DesignTabProps extends TabProps {
  pageColor: string | null
  onPageColor: (hex: string | null) => void
  section: SectionSettings | null
  onSection: (next: SectionSettings) => void
  watermark: string | null
  onWatermark: (text: string | null) => void
  themeFonts: ThemeFonts | null
  onThemeFonts: (fonts: ThemeFonts) => void
  onThemeColors: (colors: ThemeColors) => void
}

export function DesignTab({
  hasDoc,
  dropdown,
  setDropdown,
  pageColor,
  onPageColor,
  section,
  onSection,
  watermark,
  onWatermark,
  themeFonts,
  onThemeFonts,
  onThemeColors,
}: DesignTabProps) {
  const { t, lang } = useI18n()
  const watermarkPresets = [
    ...WATERMARK_KEYS.map((key) => t(key)),
    ...(lang !== 'en' ? ['CONFIDENTIAL', 'DRAFT'] : []),
  ]
  const [watermarkPromptOpen, setWatermarkPromptOpen] = useState(false)
  const applyTheme = (preset: (typeof THEME_PRESETS)[number]) => {
    onThemeFonts(preset.fonts)
    onThemeColors(preset.colors)
    setDropdown(() => null)
  }
  const customWatermark = () => {
    setWatermarkPromptOpen(true)
    setDropdown(() => null)
  }
  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonThemesTip')}
              onClick={() => toggleDropdown(setDropdown, 'theme')}
            >
              <span className="rb-big-icon">
                <IconTheme size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonThemes')}</span>
            </button>
            {dropdown === 'theme' && (
              <div className="layout-menu">
                {THEME_PRESETS.map((p) => (
                  <button key={p.nameKey} onClick={() => applyTheme(p)}>
                    <span className="theme-accent-row">
                      {['accent1', 'accent2', 'accent3', 'accent4'].map((k) => (
                        <span
                          key={k}
                          className="theme-accent-dot"
                          style={{ background: `#${p.colors[k as 'accent1']}` }}
                        />
                      ))}
                    </span>
                    {t(p.nameKey)}
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonThemeFontsTip')}
              onClick={() => toggleDropdown(setDropdown, 'themefonts')}
            >
              <span className="rb-big-icon">
                <IconThemeFonts size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonFonts')}</span>
            </button>
            {dropdown === 'themefonts' && (
              <div className="layout-menu">
                {THEME_FONT_PRESETS.map((p) => (
                  <button
                    key={p.nameKey}
                    className={
                      themeFonts?.minor === p.fonts.minor && themeFonts?.major === p.fonts.major
                        ? 'active'
                        : ''
                    }
                    onClick={() => {
                      onThemeFonts(p.fonts)
                      setDropdown(() => null)
                    }}
                  >
                    {t(p.nameKey)}
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonThemeColorsTip')}
              onClick={() => toggleDropdown(setDropdown, 'themecolors')}
            >
              <span className="rb-big-icon">
                <IconThemeColors size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonColors')}</span>
            </button>
            {dropdown === 'themecolors' && (
              <div className="layout-menu">
                {THEME_COLOR_PRESETS.map((p) => (
                  <button
                    key={p.nameKey}
                    onClick={() => {
                      onThemeColors(p.colors)
                      setDropdown(() => null)
                    }}
                  >
                    <span className="theme-accent-row">
                      {['accent1', 'accent2', 'accent3', 'accent4'].map((k) => (
                        <span
                          key={k}
                          className="theme-accent-dot"
                          style={{ background: `#${p.colors[k as 'accent1']}` }}
                        />
                      ))}
                    </span>
                    {t(p.nameKey)}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupDocFormatting')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonPageColorTip')}
              onClick={() => toggleDropdown(setDropdown, 'pagecolor')}
            >
              <span className="rb-big-icon rb-big-icon-colored">
                <IconPageColor size={BIG} />
                <span
                  className="rb-color-bar"
                  style={{ background: pageColor ? `#${pageColor}` : '#fff' }}
                />
                <IconCaret />
              </span>
              <span>{t('ribbonPageColor')}</span>
            </button>
            {dropdown === 'pagecolor' && (
              <div className="color-palette color-palette-page">
                {PAGE_COLORS.map((c) => (
                  <button
                    key={c.hex ?? 'auto'}
                    className={`color-swatch ${(pageColor ?? null) === c.hex ? 'selected' : ''}`}
                    data-tip={t(c.nameKey)}
                    aria-label={t(c.nameKey)}
                    style={{ background: c.hex ? `#${c.hex}` : '#ffffff' }}
                    onClick={() => {
                      onPageColor(c.hex)
                      setDropdown(() => null)
                    }}
                  />
                ))}
                <button
                  className="color-none"
                  onClick={() => {
                    onPageColor(null)
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonNoColor')}
                </button>
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${watermark ? 'active' : ''}`}
              disabled={!hasDoc}
              data-tip={t('ribbonWatermarkTip')}
              onClick={() => toggleDropdown(setDropdown, 'watermark')}
            >
              <span className="rb-big-icon">
                <IconWatermark size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonWatermark')}</span>
            </button>
            {dropdown === 'watermark' && (
              <div className="layout-menu">
                {watermarkPresets.map((text) => (
                  <button
                    key={text}
                    className={watermark === text ? 'active' : ''}
                    onClick={() => {
                      onWatermark(text)
                      setDropdown(() => null)
                    }}
                  >
                    {text}
                  </button>
                ))}
                <button onClick={customWatermark}>{t('ribbonWatermarkCustom')}</button>
                <button
                  className={!watermark ? 'active' : ''}
                  onClick={() => {
                    onWatermark(null)
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonWatermarkRemove')}
                </button>
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${section?.pageBorder ? 'active' : ''}`}
              disabled={!hasDoc || !section}
              data-tip={t('ribbonPageBordersTip')}
              onClick={() => toggleDropdown(setDropdown, 'pgborders')}
            >
              <span className="rb-big-icon">
                <IconPageBorders size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonPageBorders')}</span>
            </button>
            {dropdown === 'pgborders' && section && (
              <div className="layout-menu">
                <button
                  className={section.pageBorder ? 'active' : ''}
                  onClick={() => {
                    onSection({ ...section, pageBorder: true })
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonPageBorderBox')}
                </button>
                <button
                  className={!section.pageBorder ? 'active' : ''}
                  onClick={() => {
                    onSection({ ...section, pageBorder: false })
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonNone')}
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupPageBackground')}</div>
      </div>

      {watermarkPromptOpen && (
        <PromptModal
          title={t('ribbonWatermarkCustomTitle')}
          placeholder={t('ribbonWatermarkPh')}
          initial={watermark ?? ''}
          allowEmpty
          onSubmit={(text) => onWatermark(text || null)}
          onClose={() => setWatermarkPromptOpen(false)}
        />
      )}
    </>
  )
}

/* ================= Layout ================= */
