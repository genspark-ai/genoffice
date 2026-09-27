import {
  useRef,
  useState,
  type InputHTMLAttributes,
  type KeyboardEvent,
  type ReactElement,
} from 'react'

/** One named palette entry; `name` is the English color name (tooltip fallback). */
export interface ColorSwatch {
  name: string
  hex: string
}

/** Office theme colors (top row of the Word-style picker). */
export const THEME_COLORS: readonly ColorSwatch[] = [
  { name: 'White', hex: 'FFFFFF' },
  { name: 'Black', hex: '000000' },
  { name: 'Light Gray', hex: 'E7E6E6' },
  { name: 'Blue Gray', hex: '0E2841' },
  { name: 'Blue', hex: '156082' },
  { name: 'Orange', hex: 'E97132' },
  { name: 'Green', hex: '196B24' },
  { name: 'Sky Blue', hex: '0F9ED5' },
  { name: 'Purple', hex: 'A02B93' },
  { name: 'Light Green', hex: '4EA72E' },
]

/** 5 tint/shade rows under the theme colors, one column per theme color. */
export const THEME_COLOR_SHADES: readonly (readonly string[])[] = [
  [
    'F2F2F2',
    '7F7F7F',
    'D0CECE',
    'DDEBF7',
    'DDEBF7',
    'FCE4D6',
    'E2F0D9',
    'DDEBF7',
    'E4DFEC',
    'E2F0D9',
  ],
  [
    'D9D9D9',
    '595959',
    'AEAAAA',
    'BDD7EE',
    '9DC3E6',
    'F8CBAD',
    'C6E0B4',
    '9DC3E6',
    'D9E1F2',
    'C6E0B4',
  ],
  [
    'BFBFBF',
    '3F3F3F',
    '757171',
    '8EA9DB',
    '5B9BD5',
    'F4B084',
    'A9D18E',
    '5B9BD5',
    'B4C6E7',
    'A9D18E',
  ],
  [
    'A6A6A6',
    '262626',
    '3A3838',
    '4472C4',
    '2E75B6',
    'C65911',
    '70AD47',
    '00B0F0',
    '8064A2',
    '70AD47',
  ],
  [
    '808080',
    '0D0D0D',
    '171616',
    '203864',
    '1F4E78',
    '843C0C',
    '375623',
    '0070C0',
    '5B315E',
    '385723',
  ],
]

/** Word standard colors (bottom row). */
export const STANDARD_COLORS: readonly ColorSwatch[] = [
  { name: 'Dark Red', hex: 'C00000' },
  { name: 'Red', hex: 'FF0000' },
  { name: 'Orange', hex: 'FFC000' },
  { name: 'Yellow', hex: 'FFFF00' },
  { name: 'Light Green', hex: '92D050' },
  { name: 'Green', hex: '00B050' },
  { name: 'Light Blue', hex: '00B0F0' },
  { name: 'Blue', hex: '0070C0' },
  { name: 'Dark Blue', hex: '002060' },
  { name: 'Purple', hex: '7030A0' },
]

export interface ColorPickerStrings {
  themeColors: string
  standardColors: string
  /** Section title over the recentColors row (required to show that section). */
  recentColors?: string | undefined
  /** Omit to hide the "More Colors…" native-picker row. */
  moreColors?: string | undefined
  /** Label of the full-width top button ("Automatic" / "No Fill"…); omit to hide it. */
  auto?: string | undefined
  /** Tooltip for a theme shade cell (1-based row/column); defaults to the hex value. */
  shadeTip?: ((row: number, column: number) => string) | undefined
  /** Tooltip for a named swatch; defaults to the English name. */
  colorName?: ((swatch: ColorSwatch) => string) | undefined
}

export interface ColorPickerProps {
  /** Current color as #RRGGBB (any case, # optional); null/undefined = automatic/none. */
  value?: string | null | undefined
  strings: ColorPickerStrings
  /** Extra classes on the panel root (typically the app's popover positioning class). */
  className?: string | undefined
  /** Recently used colors (#RRGGBB), most recent first; shown when strings.recentColors is set. */
  recentColors?: readonly string[] | undefined
  /** Named/shade/custom pick emits "#RRGGBB" (uppercase); the auto button emits null. */
  onPick: (hex: string | null) => void
  /** Merged onto the hidden native input; lets callers override onChange (debounce,
      selection restore) or hook onPointerDown (e.g. arming the input). */
  moreInputProps?: InputHTMLAttributes<HTMLInputElement> | undefined
}

const normalizeHex = (hex: string): string => `#${hex.replace(/^#/, '').toUpperCase()}`

/** The shared Word-style color picker panel: Automatic/None, theme colors with
    tint/shade grid, standard colors and a "More Colors…" native picker entry.
    Callers own the dropdown open state and anchor positioning (via className). */
export function ColorPicker({
  value,
  strings,
  className,
  recentColors,
  onPick,
  moreInputProps,
}: ColorPickerProps): ReactElement {
  const current = value ? normalizeHex(value) : null
  const isSelected = (hex: string): boolean => current === `#${hex}`
  const rootRef = useRef<HTMLDivElement>(null)
  const [focusPos, setFocusPos] = useState<string | null>(null)

  interface Cell {
    hex: string
    title: string
    key: string
  }
  const named = (c: ColorSwatch): Cell => ({
    hex: c.hex,
    title: strings.colorName?.(c) ?? c.name,
    key: c.hex,
  })
  const rows: Cell[][] = [
    THEME_COLORS.map(named),
    ...THEME_COLOR_SHADES.map((row, r) =>
      row.map((hex, c) => ({
        hex,
        title: strings.shadeTip?.(r + 1, c + 1) ?? `#${hex}`,
        key: `${r}-${c}-${hex}`,
      })),
    ),
    STANDARD_COLORS.map(named),
  ]
  const showRecent = Boolean(strings.recentColors && recentColors && recentColors.length > 0)
  if (showRecent) {
    rows.push(
      recentColors!.map((hex, i) => {
        const bare = hex.replace(/^#/, '').toUpperCase()
        return { hex: bare, title: `#${bare}`, key: `recent-${i}-${bare}` }
      }),
    )
  }

  // one tab stop: the selected swatch, else the first; arrows rove from there
  let selectedPos = '0-0'
  rows.some((row, r) =>
    row.some((cell, c) => {
      if (!isSelected(cell.hex)) return false
      selectedPos = `${r}-${c}`
      return true
    }),
  )
  const activePos = focusPos ?? selectedPos

  const moveFocus = (e: KeyboardEvent<HTMLButtonElement>, r: number, c: number): void => {
    const len = (i: number): number => rows[i]?.length ?? 0
    let nr = r
    let nc = c
    switch (e.key) {
      case 'ArrowRight':
        nc += 1
        if (nc >= len(r)) {
          nr = (r + 1) % rows.length
          nc = 0
        }
        break
      case 'ArrowLeft':
        nc -= 1
        if (nc < 0) {
          nr = (r + rows.length - 1) % rows.length
          nc = len(nr) - 1
        }
        break
      case 'ArrowDown':
        nr = (r + 1) % rows.length
        break
      case 'ArrowUp':
        nr = (r + rows.length - 1) % rows.length
        break
      case 'Home':
        nc = 0
        break
      case 'End':
        nc = len(r) - 1
        break
      default:
        return
    }
    e.preventDefault()
    nc = Math.min(nc, len(nr) - 1)
    const pos = `${nr}-${nc}`
    setFocusPos(pos)
    rootRef.current?.querySelector<HTMLElement>(`[data-pos="${pos}"]`)?.focus()
  }

  const gridRow = (r: number): ReactElement => (
    <div key={r} role="row" className="gcp-row">
      {(rows[r] ?? []).map((cell, c) => {
        const pos = `${r}-${c}`
        const selected = isSelected(cell.hex)
        return (
          <button
            key={cell.key}
            type="button"
            role="gridcell"
            data-pos={pos}
            tabIndex={pos === activePos ? 0 : -1}
            aria-selected={selected}
            aria-label={cell.title}
            className={`gcp-swatch ${selected ? 'selected' : ''}`}
            title={cell.title}
            style={{ background: `#${cell.hex}` }}
            onMouseDown={(e) => e.preventDefault()}
            onFocus={() => setFocusPos(pos)}
            onKeyDown={(e) => moveFocus(e, r, c)}
            onClick={() => onPick(`#${cell.hex}`)}
          />
        )
      })}
    </div>
  )

  return (
    <div ref={rootRef} className={`gcp-palette${className ? ` ${className}` : ''}`}>
      {strings.auto && (
        <button
          type="button"
          className={`gcp-auto ${!current ? 'selected' : ''}`}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => onPick(null)}
        >
          {strings.auto}
        </button>
      )}
      <div className="gcp-section-title">{strings.themeColors}</div>
      <div className="gcp-theme-base" role="grid" aria-label={strings.themeColors}>
        {gridRow(0)}
      </div>
      <div className="gcp-theme-shades" role="grid" aria-label={strings.themeColors}>
        {THEME_COLOR_SHADES.map((_, r) => gridRow(r + 1))}
      </div>
      <div className="gcp-section-title">{strings.standardColors}</div>
      <div className="gcp-standard-row" role="grid" aria-label={strings.standardColors}>
        {gridRow(THEME_COLOR_SHADES.length + 1)}
      </div>
      {showRecent && (
        <>
          <div className="gcp-section-title">{strings.recentColors}</div>
          <div className="gcp-standard-row" role="grid" aria-label={strings.recentColors}>
            {gridRow(rows.length - 1)}
          </div>
        </>
      )}
      {strings.moreColors && (
        <label className="gcp-more">
          <span className="gcp-more-icon">
            <svg
              width="16"
              height="16"
              viewBox="0 0 16 16"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.2"
              aria-hidden="true"
            >
              <path d="M8 12.98a4.98 4.98 0 1 1 4.98-4.98c0 2.44-1.74 2.49-2.74 2.49-.8 0-1.25.5-1.25 1.25 0 .7-.45 1.25-1 1.25Z" />
              <circle cx="8.83" cy="4.93" r="0.71" fill="currentColor" stroke="none" />
              <circle cx="11.07" cy="6.71" r="0.71" fill="currentColor" stroke="none" />
              <circle cx="6.09" cy="5.51" r="0.71" fill="currentColor" stroke="none" />
              <circle cx="4.93" cy="8.25" r="0.71" fill="currentColor" stroke="none" />
            </svg>
          </span>
          {strings.moreColors}
          <input
            type="color"
            value={(current ?? '#4472C4').toLowerCase()}
            onChange={(e) => onPick(normalizeHex(e.currentTarget.value))}
            {...moreInputProps}
          />
        </label>
      )}
    </div>
  )
}
