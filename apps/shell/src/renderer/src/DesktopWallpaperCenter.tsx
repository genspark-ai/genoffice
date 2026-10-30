import { useCallback, useEffect, useState } from 'react'
import type { CSSProperties, ReactElement } from 'react'
import type { DesktopPrefs, HomeApi } from '../../shared/home-api'
import type { DesktopStrings } from './desktop-strings'

/**
 * The wallpaper center dialog, opened off the desktop's
 * empty-surface menu: the whole look-and-feel in one place — the animated and
 * static wallpaper tiles plus the custom-image picker, the fit row an image
 * wallpaper fills by, the label-color row with its presets, native picker and
 * live-validated HEX field, and the timed carousel row.
 *
 * Every choice applies the moment it is made, through the parent's
 * applyPrefs / applyWallpaper write-throughs; the center itself owns only the
 * HEX draft, a half-typed value that is marked invalid without ever reaching
 * the prefs. The desktop view keeps everything the backdrop shares with the
 * dialog — the open state, the prefs object, the wallpaper bytes, whether an
 * image is actually showing — and hands it all in through the props. The
 * words come from desktop-strings, the look from desktop.css.
 */

/** the carousel row: off and the four rotation periods (minutes) */
const CAROUSEL_ITEMS = [
  { minutes: 0, labelKey: 'carouselOff' },
  { minutes: 1, labelKey: 'carousel1m' },
  { minutes: 10, labelKey: 'carousel10m' },
  { minutes: 30, labelKey: 'carousel30m' },
  { minutes: 60, labelKey: 'carousel1h' },
] as const satisfies ReadonlyArray<{ minutes: number; labelKey: keyof DesktopStrings }>

/** the wallpaper fit row: how a wallpaper image fills the surface, Windows-style */
const WALLPAPER_FIT_ITEMS = [
  { fit: 'cover', labelKey: 'fitCover' },
  { fit: 'contain', labelKey: 'fitContain' },
  { fit: 'stretch', labelKey: 'fitStretch' },
  { fit: 'tile', labelKey: 'fitTile' },
  { fit: 'center', labelKey: 'fitCenter' },
] as const satisfies ReadonlyArray<{
  fit: DesktopPrefs['wallpaperFit']
  labelKey: keyof DesktopStrings
}>

/** the chosen wallpaper's thumbnail preview of the same fit (img, so the
 * mapping rides object-fit; tile previews a natural-size copy at the origin) */
const FIT_OBJECT_FIT: Record<DesktopPrefs['wallpaperFit'], CSSProperties['objectFit']> = {
  cover: 'cover',
  contain: 'contain',
  stretch: 'fill',
  tile: 'none',
  center: 'none',
}
const FIT_OBJECT_POSITION: Partial<
  Record<DesktopPrefs['wallpaperFit'], CSSProperties['objectPosition']>
> = {
  tile: 'left top',
}

/** the label-color row's presets: automatic, white, dark gray and the four brand
 * hues (logo.svg, the same values the aurora tokens paint). The swatch's own
 * fill is the color data itself, like a file badge — not UI chrome. */
const LABEL_PRESETS = [
  { value: '', labelKey: 'labelAuto' },
  { value: '#ffffff', labelKey: 'labelWhite' },
  { value: '#333333', labelKey: 'labelDark' },
  { value: '#0f7fff', labelKey: 'labelBlue' },
  { value: '#16a34a', labelKey: 'labelGreen' },
  { value: '#f97316', labelKey: 'labelOrange' },
  { value: '#ef4444', labelKey: 'labelRed' },
] as const satisfies ReadonlyArray<{ value: string; labelKey: keyof DesktopStrings }>

/** a valid label color: '' (automatic) or #rgb / #rrggbb, mirroring the store */
const LABEL_HEX_RE = /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/

/** the shipped wallpapers, as tools/gen-wallpapers.mjs's manifest lists them */
const BUILTIN_WALLPAPERS = [
  { id: 'brand-aurora', nameKey: 'wpNameAurora' },
  { id: 'starfield-night', nameKey: 'wpNameStarfield' },
  { id: 'warm-dusk', nameKey: 'wpNameDusk' },
  { id: 'dot-grid-dark', nameKey: 'wpNameDots' },
  { id: 'minimal-light', nameKey: 'wpNameMinimal' },
  { id: 'teal-deep', nameKey: 'wpNameTeal' },
] as const satisfies ReadonlyArray<{ id: string; nameKey: keyof DesktopStrings }>

/** the two animated wallpapers: the live CSS aurora, and the starfield PNG
 * with the canvas star-twinkle layer riding over it. The view imports this
 * to preload the catalog's bytes while the center is open. */
export const ANIMATED_WALLPAPERS = ['default', 'builtin:starfield-night'] as const

export interface DesktopWallpaperCenterProps {
  /** the desktop's words, already resolved for the UI language */
  readonly strings: DesktopStrings
  /** the live look-and-feel prefs: wallpaper, fit, label color, carousel */
  readonly prefs: DesktopPrefs
  /** apply a prefs patch now and persist the whole object (the view's write-through) */
  readonly applyPrefs: (patch: Partial<DesktopPrefs>) => void
  /** apply a picked wallpaper id and close the center — the backdrop re-renders
   * behind the dialog on its way out */
  readonly applyWallpaper: (id: string) => void
  /** wallpaper data URLs by id, shared with the backdrop's crossfade layers */
  readonly wallpaperImages: Readonly<Record<string, string>>
  /** whether an image wallpaper is actually showing: the live CSS aurora has
   * no image to fill, so the fit row only exists for an image */
  readonly fitApplies: boolean
  /** close the center: the scrim, the close button, Escape, or an applied choice */
  readonly onClose: () => void
}

export function DesktopWallpaperCenter({
  strings,
  prefs,
  applyPrefs,
  applyWallpaper,
  wallpaperImages,
  fitApplies,
  onClose,
}: DesktopWallpaperCenterProps): ReactElement {
  const s = strings

  // ── the HEX field: a live-validated draft of labelColor ──
  // The draft follows every applied color (preset, picker, launch); typing only
  // replaces it once the text is a whole valid hex, so half-typed values are
  // marked invalid without ever reaching the prefs.
  const [hexDraft, setHexDraft] = useState(prefs.labelColor)
  useEffect(() => {
    setHexDraft(prefs.labelColor)
  }, [prefs.labelColor])

  /** type into the HEX field: '' resets to automatic, a valid #rgb/#rrggbb
   * applies at once, anything else is marked and dropped */
  const onHexInput = useCallback(
    (value: string) => {
      setHexDraft(value)
      const trimmed = value.trim()
      if (trimmed === '' || LABEL_HEX_RE.test(trimmed)) {
        applyPrefs({ labelColor: trimmed === '' ? '' : trimmed.toLowerCase() })
      }
    },
    [applyPrefs],
  )
  const hexInvalid = hexDraft !== '' && !LABEL_HEX_RE.test(hexDraft.trim())

  /** the selected thumbnail's live preview of the current fit */
  const fitPreviewStyle = (active: boolean): CSSProperties | undefined =>
    active
      ? {
          objectFit: FIT_OBJECT_FIT[prefs.wallpaperFit],
          objectPosition: FIT_OBJECT_POSITION[prefs.wallpaperFit],
        }
      : undefined

  /** the custom-image tile: pick, copy, and apply whatever came back */
  const pickCustomWallpaper = useCallback(() => {
    const request = (window.aiOffice as Partial<HomeApi>).wallpaperPickCustom?.()
    if (!request) return
    void request.then(
      (id) => {
        if (id !== '') applyWallpaper(id)
      },
      () => undefined,
    )
  }, [applyWallpaper])

  return (
    <div
      className="set-overlay"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onClose()
      }}
    >
      <div
        className="set-dialog desktop-wp-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={s.menuWallpaper}
      >
        <div className="set-header">
          <div>
            <h2 className="set-title">{s.menuWallpaper}</h2>
          </div>
          <button className="set-close" aria-label={s.wpClose} onClick={() => onClose()}>
            <svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true">
              <path
                d="M2 2l10 10M12 2L2 12"
                stroke="currentColor"
                strokeWidth="1.5"
                strokeLinecap="round"
              />
            </svg>
          </button>
        </div>
        <div className="desktop-wp-body">
          {fitApplies && (
            <section className="desktop-wp-group">
              <h3 className="desktop-wp-group-title">{s.wpFit}</h3>
              <div className="desktop-wp-fit">
                {WALLPAPER_FIT_ITEMS.map(({ fit, labelKey }) => (
                  <button
                    key={fit}
                    type="button"
                    className={prefs.wallpaperFit === fit ? 'active' : undefined}
                    onClick={() => applyPrefs({ wallpaperFit: fit })}
                  >
                    {s[labelKey]}
                  </button>
                ))}
              </div>
            </section>
          )}
          <section className="desktop-wp-group">
            <h3 className="desktop-wp-group-title">{s.wpLabelColor}</h3>
            <div className="desktop-wp-label-color">
              {LABEL_PRESETS.map(({ value, labelKey }) => (
                <button
                  key={value}
                  type="button"
                  className={`desktop-wp-swatch${value === '' ? ' auto' : ''}${prefs.labelColor === value ? ' active' : ''}`}
                  style={value === '' ? undefined : { backgroundColor: value }}
                  aria-label={s[labelKey]}
                  aria-pressed={prefs.labelColor === value}
                  title={s[labelKey]}
                  onClick={() => applyPrefs({ labelColor: value })}
                />
              ))}
              <label className="desktop-wp-color-chip" title={s.labelCustom}>
                <input
                  type="color"
                  value={prefs.labelColor === '' ? '#ffffff' : prefs.labelColor}
                  aria-label={s.labelCustom}
                  onChange={(event) => applyPrefs({ labelColor: event.target.value })}
                />
              </label>
              <input
                type="text"
                className={`desktop-wp-hex${hexInvalid ? ' invalid' : ''}`}
                aria-label={s.labelHex}
                aria-invalid={hexInvalid || undefined}
                spellCheck={false}
                value={hexDraft}
                placeholder="#RRGGBB"
                onChange={(event) => onHexInput(event.target.value)}
              />
            </div>
          </section>
          <section className="desktop-wp-group">
            <h3 className="desktop-wp-group-title">{s.wpCarousel}</h3>
            <div className="desktop-wp-carousel">
              {CAROUSEL_ITEMS.map(({ minutes, labelKey }) => (
                <button
                  key={minutes}
                  type="button"
                  className={prefs.carouselMinutes === minutes ? 'active' : undefined}
                  onClick={() => applyPrefs({ carouselMinutes: minutes })}
                >
                  {s[labelKey]}
                </button>
              ))}
            </div>
          </section>
          <section className="desktop-wp-group">
            <h3 className="desktop-wp-group-title">{s.wpAnimated}</h3>
            <div className="desktop-wp-grid">
              {ANIMATED_WALLPAPERS.map((id) => (
                <button
                  key={id}
                  type="button"
                  className={`desktop-wp-tile${prefs.wallpaper === id ? ' active' : ''}`}
                  onClick={() => applyWallpaper(id)}
                >
                  <span className="desktop-wp-thumb">
                    {id === 'default' ? (
                      <span className="desktop-wp-preview-aurora" />
                    ) : (
                      wallpaperImages[id] !== undefined && (
                        <img
                          src={wallpaperImages[id]}
                          alt=""
                          style={fitPreviewStyle(prefs.wallpaper === id)}
                        />
                      )
                    )}
                  </span>
                  <span className="desktop-wp-name">
                    {id === 'default' ? s.wpNameDefault : s.wpNameStarfield}
                  </span>
                </button>
              ))}
            </div>
          </section>
          <section className="desktop-wp-group">
            <h3 className="desktop-wp-group-title">{s.wpStatic}</h3>
            <div className="desktop-wp-grid">
              {BUILTIN_WALLPAPERS.filter(({ id }) => id !== 'starfield-night').map(
                ({ id, nameKey }) => (
                  <button
                    key={id}
                    type="button"
                    className={`desktop-wp-tile${prefs.wallpaper === `builtin:${id}` ? ' active' : ''}`}
                    onClick={() => applyWallpaper(`builtin:${id}`)}
                  >
                    <span className="desktop-wp-thumb">
                      {wallpaperImages[`builtin:${id}`] !== undefined && (
                        <img
                          src={wallpaperImages[`builtin:${id}`]}
                          alt=""
                          style={fitPreviewStyle(prefs.wallpaper === `builtin:${id}`)}
                        />
                      )}
                    </span>
                    <span className="desktop-wp-name">{s[nameKey]}</span>
                  </button>
                ),
              )}
            </div>
          </section>
          <section className="desktop-wp-group">
            <h3 className="desktop-wp-group-title">{s.wpCustom}</h3>
            <button
              type="button"
              className={`desktop-wp-tile desktop-wp-custom${prefs.wallpaper.startsWith('custom:') ? ' active' : ''}`}
              onClick={pickCustomWallpaper}
            >
              <span className="desktop-wp-custom-glyph">
                <PlusGlyph />
              </span>
              <span className="desktop-wp-custom-text">
                <span className="desktop-wp-name">{s.wpCustom}</span>
                <span className="desktop-wp-custom-hint">{s.wpCustomHint}</span>
              </span>
            </button>
          </section>
        </div>
      </div>
    </div>
  )
}

// ── glyphs ──────────────────────────────────────────────

function PlusGlyph() {
  return (
    <svg width="20" height="20" viewBox="0 0 14 14" fill="none" aria-hidden="true">
      <path d="M7 2.5v9M2.5 7h9" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
    </svg>
  )
}
