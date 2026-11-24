import { useEffect, useState, type KeyboardEvent, type MouseEvent } from 'react'
import {
  OFFICE_THEMES,
  OFFICE_THEME_COLORS,
  OFFICE_THEME_FONTS,
  type HfImage,
  type PictureWatermarkSpec,
  type SectionSettings,
  type StyleSetId,
  type ThemeColors,
  type ThemeFonts,
} from '@genoffice/docx-engine'
import { useI18n, type StringKey } from '../i18n/locale'
import type { ActionCommandId } from '../../shared/commands'
import { CUSTOM_STYLE_SET, CUSTOM_THEME, NONE } from '../commands/menu-commands'
import type { SavedTheme } from '../saved-themes'
import { cssFontFamily } from '../text-layout/font-rules'
import { RbCommand, useRunCommand } from './ribbon-command'
import {
  IconCaret,
  IconPageBorders,
  IconPageColor,
  IconTheme,
  IconThemeColors,
  IconThemeFonts,
  IconWatermark,
  IconLineSpacing,
  IconSetDefault,
} from './icons'

/** icon size for the big icon-over-label ribbon buttons (slides ribbon parity) */
import { BIG, TabProps, toggleDropdown } from './ribbon-tabs'
import { PARAGRAPH_SPACINGS, type ParagraphSpacing } from '../commands/paragraph-spacing'
import { RibbonColorPalette } from './ribbon-color-palette'
import { PageBorderDialog, type PageBorderScope } from './PageBorderDialog'
import { styleLabel } from '../style-gallery'
import {
  styleSetLineCss,
  styleSetPreviews,
  type StyleSetPreview,
  type StyleSetSource,
} from '../style-set-previews'

const STYLE_SET_NAMES: Record<StyleSetId, StringKey> = {
  distinctive: 'ribbonStyleSetDistinctive',
  elegant: 'ribbonStyleSetElegant',
  fancy: 'ribbonStyleSetFancy',
  formal: 'ribbonStyleSetFormal',
  manuscript: 'ribbonStyleSetManuscript',
  modern: 'ribbonStyleSetModern',
  newsprint: 'ribbonStyleSetNewsprint',
  perspective: 'ribbonStyleSetPerspective',
  simple: 'ribbonStyleSetSimple',
  thatch: 'ribbonStyleSetThatch',
  traditional: 'ribbonStyleSetTraditional',
  word2010: 'ribbonStyleSetWord2010',
}

/** the cards of the style set gallery: the document as it stands, then Word's sets on its theme */
function useStyleSetPreviews(source: StyleSetSource | null): StyleSetPreview[] {
  const [previews, setPreviews] = useState<StyleSetPreview[]>([])
  useEffect(() => {
    if (!source) return setPreviews([])
    let cancelled = false
    styleSetPreviews(source)
      .then((next) => {
        if (!cancelled) setPreviews(next)
      })
      .catch(() => {})
    return () => {
      cancelled = true
    }
  }, [source])
  return previews
}

/** Word-like theme presets: font pair + color scheme applied together */
const SPACING_LABELS: Record<ParagraphSpacing['id'], StringKey> = {
  none: 'ribbonPsNone',
  compact: 'ribbonPsCompact',
  tight: 'ribbonPsTight',
  open: 'ribbonPsOpen',
  relaxed: 'ribbonPsRelaxed',
  double: 'ribbonPsDouble',
}

const ACCENTS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const
const SWATCHES = ['dk2', 'lt2', ...ACCENTS] as const

/** a scheme's colours side by side; they are the document's, not the app's */
function ColorStrip({
  colors,
  slots,
}: {
  colors: ThemeColors
  slots: readonly (keyof ThemeColors)[]
}) {
  return (
    <span className="theme-strip">
      {slots.map((slot) => (
        <span key={slot} style={{ background: `#${colors[slot]}` }} />
      ))}
    </span>
  )
}

/** arrow keys walk a gallery's cards the way they sit on screen, across its sections */
function galleryKeys(e: KeyboardEvent<HTMLElement>) {
  if (!['ArrowRight', 'ArrowLeft', 'ArrowDown', 'ArrowUp'].includes(e.key)) return
  const cards = [
    ...e.currentTarget.querySelectorAll<HTMLElement>(
      '.theme-card, .style-set-card, .theme-list > button, .theme-foot > button',
    ),
  ]
  const at = cards.indexOf(document.activeElement as HTMLElement)
  if (at < 0) return
  let next: HTMLElement | undefined
  if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
    next = cards[at + (e.key === 'ArrowRight' ? 1 : -1)]
  } else {
    const box = (card: HTMLElement) => card.getBoundingClientRect()
    const here = box(cards[at]!)
    const down = e.key === 'ArrowDown'
    const beyond = cards.filter((c) => (down ? box(c).top > here.top : box(c).top < here.top))
    // the nearest row that way, and in it the card most nearly above or below this one
    const row = beyond.reduce<number | null>(
      (best, c) =>
        best === null || (down ? box(c).top < best : box(c).top > best) ? box(c).top : best,
      null,
    )
    next = beyond
      .filter((c) => box(c).top === row)
      .sort((a, b) => Math.abs(box(a).left - here.left) - Math.abs(box(b).left - here.left))[0]
  }
  if (!next) return
  e.preventDefault()
  next.focus()
}

/** the preset texts of the Watermark gallery; other languages also get the two English ones */
export const WATERMARK_KEYS: StringKey[] = [
  'ribbonWatermarkConfidential',
  'ribbonWatermarkNoCopy',
  'ribbonWatermarkDraft',
  'ribbonWatermarkSample',
]

export const watermarkPresetsOf = (t: (key: StringKey) => string, lang: string): string[] => [
  ...WATERMARK_KEYS.map((key) => t(key)),
  ...(lang !== 'en' ? ['CONFIDENTIAL', 'DRAFT'] : []),
]

interface DesignTabProps extends TabProps {
  pageColor: string | null
  section: SectionSettings | null
  onSection: (next: SectionSettings) => void
  /** Page Border dialog > Apply to: Whole document */
  onPageBorderAll: (props: SectionSettings['pageBorderProps'] | null) => void
  watermark: string | null
  watermarkPicture: PictureWatermarkSpec | null
  /** the saved document's watermark shape, for reopening the dialog on it */
  watermarkDocImage: HfImage | null
  themeFonts: ThemeFonts | null
  themeColors: ThemeColors | null
  styleSetSource: StyleSetSource | null
  savedThemes: readonly SavedTheme[]
}

export function DesignTab({
  hasDoc,
  dropdown,
  setDropdown,
  pageColor,
  section,
  onSection,
  onPageBorderAll,
  watermark,
  watermarkPicture,
  watermarkDocImage,
  themeFonts,
  themeColors,
  styleSetSource,
  savedThemes,
}: DesignTabProps) {
  const { t, lang } = useI18n()
  const [pageBorderDialog, setPageBorderDialog] = useState(false)
  const applyPageBorder = (
    props: SectionSettings['pageBorderProps'] | null,
    scope: PageBorderScope,
  ) => {
    if (scope === 'whole') return onPageBorderAll(props)
    if (!section) return
    const { pageBorderProps: _drop, ...rest } = section
    onSection(
      props
        ? { ...rest, pageBorder: true, pageBorderProps: props }
        : { ...rest, pageBorder: false },
    )
  }
  const watermarkPresets = watermarkPresetsOf(t, lang)
  const runCommand = useRunCommand()
  /** a pick from a gallery: run the command, close the gallery */
  const pick = (id: ActionCommandId, payload?: string) => {
    runCommand(id, payload)
    setDropdown(() => null)
  }
  const onFonts = (fonts: ThemeFonts) =>
    themeFonts?.minor === fonts.minor && themeFonts?.major === fonts.major
  const onColors = (colors: ThemeColors) =>
    !!themeColors && SWATCHES.every((k) => themeColors[k]?.toUpperCase() === colors[k])
  const styleSets = useStyleSetPreviews(hasDoc ? styleSetSource : null)
  const [setMenu, setSetMenu] = useState<{ id: string; x: number; y: number } | null>(null)
  const [themeMenu, setThemeMenu] = useState<{ file: string; x: number; y: number } | null>(null)
  const themeCard = (
    theme: { name: string; fonts: ThemeFonts; colors: ThemeColors },
    testId: string,
    payload: string,
    file?: string,
  ) => {
    const active = onFonts(theme.fonts) && onColors(theme.colors)
    return (
      <button
        key={payload}
        className={`theme-card${active ? ' active' : ''}`}
        data-testid={testId}
        aria-pressed={active}
        onClick={() => pick('document-theme', payload)}
        // command-exempt: opens the card's own right-click menu
        onContextMenu={(e) => {
          if (!file) return
          e.preventDefault()
          setThemeMenu({ file, x: e.clientX, y: e.clientY })
        }}
      >
        <span
          className="theme-thumb"
          style={{ background: `#${theme.colors.lt1}`, color: `#${theme.colors.dk2}` }}
        >
          <span style={{ fontFamily: cssFontFamily(theme.fonts.major) }}>A</span>
          <span style={{ fontFamily: cssFontFamily(theme.fonts.minor) }}>a</span>
          <ColorStrip colors={theme.colors} slots={ACCENTS} />
        </span>
        <span className="theme-name">{theme.name}</span>
      </button>
    )
  }
  const styleSetCard = (p: StyleSetPreview) => {
    const key = p.custom ? `${CUSTOM_STYLE_SET}${p.custom.id}` : (p.id ?? 'current')
    const name = p.custom?.name ?? t(p.id ? STYLE_SET_NAMES[p.id] : 'ribbonStyleSetThisDoc')
    const line = (info: StyleSetPreview['title'], builtIn: string) => (
      <span style={styleSetLineCss(info ?? p.body, p.docDefaults)}>
        {styleLabel(info ?? { name: builtIn }, t)}
      </span>
    )
    // command-exempt: opens the card's own right-click menu
    const menu = (e: MouseEvent) => {
      if (!p.custom) return
      e.preventDefault()
      setSetMenu({ id: p.custom.id, x: e.clientX, y: e.clientY })
    }
    return (
      <button
        key={key}
        className={`style-set-card${key === 'current' ? ' active' : ''}`}
        data-testid={`style-set-${p.custom ? `custom-${p.custom.name}` : key}`}
        data-tip={name}
        aria-label={name}
        onClick={() => key !== 'current' && pick('design-style-set', key)}
        onContextMenu={menu}
      >
        <span className="style-set-page" style={{ background: '#fff', color: '#000' }}>
          {line(p.title, 'Title')}
          {line(p.heading, 'heading 1')}
          <span style={styleSetLineCss(p.body, p.docDefaults)}>{t('ribbonStyleSetSample')}</span>
        </span>
      </button>
    )
  }
  const customCards = styleSets.filter((p) => p.custom).map(styleSetCard)
  const builtInCards = styleSets.filter((p) => p.id).map(styleSetCard)
  const currentCard = styleSets.filter((p) => !p.id && !p.custom).map(styleSetCard)
  return (
    <>
      <div className="ribbon-group style-set-group">
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
              <div
                data-rb-panel=""
                className="theme-gallery"
                role="group"
                aria-label={t('ribbonThemes')}
                onKeyDown={galleryKeys}
              >
                {savedThemes.length > 0 && (
                  <>
                    <div className="layout-menu-head">{t('ribbonStyleSetCustom')}</div>
                    <div className="theme-grid">
                      {savedThemes.map((saved) =>
                        themeCard(
                          { ...saved.theme, name: saved.name },
                          `theme-custom-${saved.name}`,
                          `${CUSTOM_THEME}${saved.file}`,
                          saved.file,
                        ),
                      )}
                    </div>
                    <div className="layout-menu-head">Office</div>
                  </>
                )}
                <div className="theme-grid">
                  {OFFICE_THEMES.map((theme) => themeCard(theme, `theme-${theme.id}`, theme.id))}
                </div>
                <div className="rb-menu-sep" />
                <div className="style-set-foot theme-foot">
                  <button data-testid="theme-reset" onClick={() => pick('theme-reset')}>
                    {t('ribbonThemeReset')}
                  </button>
                  <button data-testid="theme-browse" onClick={() => pick('theme-browse')}>
                    {t('ribbonThemeBrowse')}
                  </button>
                  <button data-testid="theme-save" onClick={() => pick('theme-save-current')}>
                    {t('ribbonThemeSave')}
                  </button>
                </div>
              </div>
            )}
            {themeMenu && (
              <>
                {/* command-exempt: closes the card's own right-click menu */}
                <div className="style-set-ctx-backdrop" onMouseDown={() => setThemeMenu(null)} />
                <div className="style-set-ctx" style={{ left: themeMenu.x, top: themeMenu.y }}>
                  <button
                    data-testid="theme-delete"
                    onClick={() => {
                      runCommand('theme-custom-delete', themeMenu.file)
                      setThemeMenu(null)
                    }}
                  >
                    {t('appNavDelete')}
                  </button>
                </div>
              </>
            )}
          </div>
          <div className="rb-split-wrap style-set-wrap">
            <div className="style-set-strip">
              {currentCard}
              {customCards}
              {builtInCards}
            </div>
            <button
              className="rb-styles-more"
              disabled={!hasDoc}
              data-testid="style-set-more"
              data-tip={t('ribbonStyleSets')}
              aria-label={t('ribbonStyleSets')}
              onClick={() => toggleDropdown(setDropdown, 'stylesets')}
            >
              <IconCaret />
            </button>
            {dropdown === 'stylesets' && (
              <div data-rb-panel="" className="style-set-menu" onKeyDown={galleryKeys}>
                <div className="layout-menu-head">{t('ribbonStyleSetThisDoc')}</div>
                <div className="style-set-grid">{currentCard}</div>
                {customCards.length > 0 && (
                  <>
                    <div className="layout-menu-head">{t('ribbonStyleSetCustom')}</div>
                    <div className="style-set-grid">{customCards}</div>
                  </>
                )}
                <div className="layout-menu-head">{t('ribbonStyleSetBuiltIn')}</div>
                <div className="style-set-grid">{builtInCards}</div>
                <div className="rb-menu-sep" />
                <div className="style-set-foot">
                  <button
                    data-testid="style-set-reset"
                    onClick={() => pick('design-style-set-reset')}
                  >
                    {t('ribbonStyleSetReset')}
                  </button>
                  <button
                    data-testid="style-set-save"
                    onClick={() => pick('design-style-set-save')}
                  >
                    {t('ribbonStyleSetSave')}
                  </button>
                </div>
              </div>
            )}
            {setMenu && (
              <>
                {/* command-exempt: closes the card's own right-click menu */}
                <div className="style-set-ctx-backdrop" onMouseDown={() => setSetMenu(null)} />
                <div className="style-set-ctx" style={{ left: setMenu.x, top: setMenu.y }}>
                  <button
                    data-testid="style-set-delete"
                    onClick={() => {
                      runCommand('design-style-set-delete', setMenu.id)
                      setSetMenu(null)
                    }}
                  >
                    {t('appNavDelete')}
                  </button>
                </div>
              </>
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
              <div data-rb-panel="" className="layout-menu theme-list" onKeyDown={galleryKeys}>
                {OFFICE_THEME_COLORS.map((scheme) => (
                  <button
                    key={scheme.id}
                    className={onColors(scheme.colors) ? 'active' : ''}
                    data-testid={`theme-colors-${scheme.id}`}
                    onClick={() => pick('theme-colors', scheme.id)}
                  >
                    <ColorStrip colors={scheme.colors} slots={SWATCHES} />
                    {scheme.name}
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
              <div data-rb-panel="" className="layout-menu theme-list" onKeyDown={galleryKeys}>
                {OFFICE_THEME_FONTS.map((scheme) => (
                  <button
                    key={scheme.id}
                    className={onFonts(scheme.fonts) ? 'active' : ''}
                    data-testid={`theme-fonts-${scheme.id}`}
                    onClick={() => pick('theme-fonts', scheme.id)}
                  >
                    <span className="theme-font-sample">
                      <span style={{ fontFamily: cssFontFamily(scheme.fonts.major) }}>A</span>
                      <span style={{ fontFamily: cssFontFamily(scheme.fonts.minor) }}>a</span>
                    </span>
                    <span className="theme-font-names">
                      {scheme.name}
                      <span>
                        {scheme.fonts.major} / {scheme.fonts.minor}
                      </span>
                    </span>
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="rb-col rb-pair">
            <div className="rb-split-wrap">
              <button
                className="rb-small"
                disabled={!hasDoc}
                data-tip={t('ribbonParagraphSpacing')}
                onClick={() => toggleDropdown(setDropdown, 'paraspacing')}
              >
                <IconLineSpacing size={14} />
                <span className="rb-small-label">{t('ribbonParagraphSpacing')}</span>
                <IconCaret />
              </button>
              {dropdown === 'paraspacing' && (
                <div data-rb-panel="" className="layout-menu">
                  {PARAGRAPH_SPACINGS.map((s) => (
                    <button key={s.id} onClick={() => pick('design-paragraph-spacing', s.id)}>
                      {t(SPACING_LABELS[s.id])}
                    </button>
                  ))}
                  <button onClick={() => pick('design-custom-paragraph-spacing')}>
                    {t('ribbonPsCustom')}
                  </button>
                </div>
              )}
            </div>
            <RbCommand cmd="design-set-default" look="small">
              <IconSetDefault size={14} />
            </RbCommand>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupDocFormatting')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${watermark || watermarkPicture || watermarkDocImage ? 'active' : ''}`}
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
              <div data-rb-panel="" className="layout-menu">
                {watermarkPresets.map((text) => (
                  <button
                    key={text}
                    className={watermark === text ? 'active' : ''}
                    onClick={() => pick('watermark', text)}
                  >
                    {text}
                  </button>
                ))}
                <button onClick={() => pick('insert-watermark')}>
                  {t('ribbonWatermarkCustom')}
                </button>
                <button
                  className={!watermark && !watermarkPicture && !watermarkDocImage ? 'active' : ''}
                  onClick={() => pick('watermark', NONE)}
                >
                  {t('ribbonWatermarkRemove')}
                </button>
              </div>
            )}
          </div>
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
              <RibbonColorPalette
                current={pageColor}
                noneLabel={t('ribbonNoColor')}
                onPick={(hex) => pick('page-color', hex ?? NONE)}
              />
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
              <div data-rb-panel="" className="layout-menu">
                <button
                  className={section.pageBorder && !section.pageBorderProps ? 'active' : ''}
                  onClick={() => pick('page-border', 'box')}
                >
                  {t('ribbonPageBorderBox')}
                </button>
                <button
                  className={!section.pageBorder ? 'active' : ''}
                  onClick={() => pick('page-border', NONE)}
                >
                  {t('ribbonNone')}
                </button>
                <div className="rb-menu-sep" />
                <button
                  onClick={() => {
                    setPageBorderDialog(true)
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonPageBorderDialog')}
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupPageBackground')}</div>
      </div>

      {pageBorderDialog && section && (
        <PageBorderDialog
          section={section}
          onApply={applyPageBorder}
          onClose={() => setPageBorderDialog(false)}
        />
      )}
    </>
  )
}

/* ================= Layout ================= */
