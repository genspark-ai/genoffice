import type { ReactElement, ReactNode } from 'react'

/// Ribbon icon set drawn to the shared icon standard: 24×24 canvas with
/// 1.5-unit strokes, round caps and round joins. The rendered size comes
/// from CSS, which also pins the PAINTED stroke to ~1.5px (20px+ glyphs) /
/// ~1.25px (13-16px glyphs) via stroke-width overrides in styles.css —
/// proportional scaling would leave small glyphs hairline and big ones fat.
function Icon({ children }: { readonly children: ReactNode }): ReactElement {
  return (
    <svg
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.5}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {children}
    </svg>
  )
}

/// Thin dropdown chevron, copied from the slides ribbon's
/// RbCaret (replaces the ▾ text glyph).
export function CaretIcon(): ReactElement {
  return (
    <svg className="chev" width="10" height="10" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M5.5 9.25 12 15.75l6.5-6.5"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

/// Quick-access save button — standard floppy glyph shared across all apps.
export function SaveIcon(): ReactElement {
  return (
    <Icon>
      <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
      <path d="M17 21v-8H7v8" />
      <path d="M7 3v5h8V3" />
    </Icon>
  )
}

/// Quick-access undo/redo — design-supplied geometry shared with the docs
/// and slides ribbons, so the three apps' tab-row arrows read identically.
export function UndoIcon(): ReactElement {
  return (
    <Icon>
      <path d="M5.91026 4L2.5 7.14791L5.91026 10.8205" />
      <path d="M3.96154 7.41028H15.1636C18.5169 7.41028 21.3646 10.1484 21.4953 13.5C21.6334 17.0416 18.707 20.0769 15.1636 20.0769H6.88384" />
    </Icon>
  )
}

export function RedoIcon(): ReactElement {
  return (
    <Icon>
      <path d="M18.0897 4L21.5 7.14791L18.0897 10.8205" />
      <path d="M20.0385 7.41028H8.83636C5.4831 7.41028 2.63537 10.1484 2.5047 13.5C2.36657 17.0416 5.29296 20.0769 8.83636 20.0769H17.1162" />
    </Icon>
  )
}

// drawings shared by several glyph keys
const SPARKLE_ICON = (
  <Icon>
    <path d="M12 4.75 13.9 10.1 19.25 12l-5.35 1.9L12 19.25l-1.9-5.35L4.75 12l5.35-1.9L12 4.75Z" />
  </Icon>
)
const CLOCK_ICON = (
  <Icon>
    <circle cx="12" cy="12" r="7.5" />
    <path d="M12 7.75V12l3 2.25" />
  </Icon>
)
const REFRESH_ICON = (
  <Icon>
    <path d="M18.74 12A6.74 6.74 0 1 1 16.65 7.12" />
    <path d="M16.88 4.1v3.25h-3.25" />
  </Icon>
)

/// Icons keyed by the legacy glyph strings, so ribbon call sites stay
/// unchanged; symbols without an entry render as plain text (letterforms
/// such as $, ?, θ, ƒx are typography, not drawn icons).
export const RIBBON_GLYPH_ICONS: Record<string, ReactElement> = {
  // ---- alignment / rows / columns ----
  '≡': (
    <Icon>
      <path d="M4.75 4.75h14.5M4.75 9.5h9.5M4.75 14.25h14.5M4.75 19h9.5" />
    </Icon>
  ),
  '≣': (
    <Icon>
      <path d="M4.75 4.75h14.5M7.25 9.5h9.5M4.75 14.25h14.5M7.25 19h9.5" />
    </Icon>
  ),
  '☰': (
    <Icon>
      <path d="M4.75 4.75h14.5M9.75 9.5h9.5M4.75 14.25h14.5M9.75 19h9.5" />
    </Icon>
  ),
  '⤒': (
    <Icon>
      <path d="M5 4.75h14" />
      <path d="M12 19.25V9.5M8.5 13 12 9.5l3.5 3.5" />
    </Icon>
  ),
  '⤓': (
    <Icon>
      <path d="M5 19.25h14" />
      <path d="M12 4.75v9.75M8.5 11l3.5 3.5 3.5-3.5" />
    </Icon>
  ),
  '↩': (
    <Icon>
      <path d="M4.75 5.75h14.5" />
      <path d="M4.75 11.25h10.25a3 3 0 0 1 0 6h-3.25" />
      <path d="M14 15 11.75 17.25 14 19.5" />
    </Icon>
  ),
  '⇤': (
    <Icon>
      <path d="M5 5v14" />
      <path d="M19 12H9.5M13 8.5 9.5 12l3.5 3.5" />
    </Icon>
  ),
  '⇥': (
    <Icon>
      <path d="M19 5v14" />
      <path d="M5 12h9.5M11 8.5 14.5 12 11 15.5" />
    </Icon>
  ),
  // ---- mini chart-type grid ----
  '▮▬': (
    <Icon>
      <path d="M4.5 19.5h15" />
      <path d="M8 19.5V12M12.5 19.5V8.5M17 19.5V15" />
    </Icon>
  ),
  '📈': (
    <Icon>
      <path d="M4.5 4.5v15h15" />
      <path d="m7.5 15.5 3.5-4 2.7 2.5 4.3-5.5" />
    </Icon>
  ),
  '◪': (
    <Icon>
      <path d="M4.75 19.25V13.5l5.5-6 4 5.5 5-4.25v10.5H4.75Z" />
    </Icon>
  ),
  '◔': (
    <Icon>
      <circle cx="12" cy="12" r="7.5" />
      <path d="M12 12V4.5M12 12l6.75 3.25" />
    </Icon>
  ),
  '∴': (
    <Icon>
      <path d="M4.5 4.5v15h15" />
      <circle cx="9" cy="15" r="0.5" fill="currentColor" />
      <circle cx="12" cy="9.5" r="0.5" fill="currentColor" />
      <circle cx="16" cy="12.5" r="0.5" fill="currentColor" />
      <circle cx="18" cy="7" r="0.5" fill="currentColor" />
    </Icon>
  ),
  '✳': (
    <Icon>
      <path d="M12 4.25 18.75 8.1v7.8L12 19.75 5.25 15.9V8.1L12 4.25Z" />
      <path d="M12 12V4.25M12 12l6.75 3.9M12 12l-6.75 3.9" />
    </Icon>
  ),
  '◍': (
    <Icon>
      <circle cx="12" cy="12" r="7.25" />
      <circle cx="12" cy="12" r="3" />
    </Icon>
  ),
  '𝄜': (
    <Icon>
      <path d="M4.5 19.5h15" />
      <path d="M7 19.5v-5.5M11 19.5V10" />
      <path d="m5 8.5 4.5-3.25L14 8l5-3.75" />
    </Icon>
  ),
  // ---- clipboard ----
  // paste shares the docs ribbon's clipboard geometry (docs 16-canvas × 1.5);
  // paste special swaps the text lines for option sliders on the same board
  '📋': (
    <Icon>
      <path d="M 8.14 18.16 H 5.89 C 4.85 18.16 4 17.31 4 16.26 V 6.79 C 4 5.74 4.85 4.89 5.89 4.89 H 7.32 M 17.26 7.74 V 6.79 C 17.26 5.74 16.41 4.89 15.37 4.89 H 13.95" />
      <rect x="7.79" y="3" width="5.68" height="2.84" rx="0.95" />
      <path d="M 18.21 7.74 H 9.68 C 8.64 7.74 7.79 8.59 7.79 9.63 V 19.11 C 7.79 20.15 8.64 21 9.68 21 H 15.49 L 20.11 16.03 V 9.63 C 20.11 8.59 19.26 7.74 18.21 7.74 Z" />
      <path d="M 10.63 11.05 H 17.26 M 10.63 14.37 H 14.42" />
      <path d="M 15.37 21 V 16.74 C 15.37 16.21 15.79 15.79 16.32 15.79 H 20.11" />
    </Icon>
  ),
  '📑': (
    <Icon>
      <path d="M 14.99 5.02 H 16.98 C 18.09 5.02 18.98 5.92 18.98 7.04 L 18.88 19.02 C 18.87 20.11 17.98 21 16.88 21 H 6.88 C 5.77 21 4.87 20.09 4.88 18.98 L 4.98 7 C 4.98 5.91 5.88 5.02 6.98 5.02 H 8.99" />
      <rect x="9.16" y="3.5" width="5.68" height="2.84" rx="0.95" />
      <path d="M 8 10 H 16 M 8 13.5 H 16 M 8 17 H 13" />
    </Icon>
  ),
  // cut / copy / format painter share one 24-canvas design across all three
  // apps (the docs ribbon carries the same geometry × 2/3 on its 16 canvas)
  '✂': (
    <Icon>
      <path d="M 7.91 19.27 L 8.92 17.57 L 16.98 3.59 M 7.02 3.5 L 15.08 17.47 L 16.09 19.27" />
      <circle cx="5.83" cy="18.13" r="2.37" />
      <circle cx="18.17" cy="18.13" r="2.37" />
    </Icon>
  ),
  '⧉': (
    <Icon>
      <rect x="7" y="7" width="14" height="14" rx="3" />
      <path d="M 14.5 4 H 7 C 5.34 4 4 5.34 4 7 V 14.5" />
    </Icon>
  ),
  '🖌': (
    <Icon>
      <rect x="10.65" y="4.05" width="2.7" height="5.1" rx="1.35" />
      <rect x="4.5" y="9.15" width="15" height="10.8" rx="1.5" />
      <path d="M 4.5 13.35 H 19.5" />
      <path d="M 9.3 16.35 V 18.15 M 14.7 16.35 V 18.15" />
    </Icon>
  ),
  // ---- find / filter ----
  '🔍': (
    <Icon>
      <circle cx="10.5" cy="10.5" r="6" />
      <path d="m15 15 4.5 4.5" />
    </Icon>
  ),
  '▽': (
    <Icon>
      <path d="M 4.17 4.71 h 15.66 l -5.94 6.8 v 7.45 l -3.78 -2.38 v -5.08 l -5.94 -6.8 Z" />
    </Icon>
  ),
  '⊘': (
    <Icon>
      <path d="M4 5.25h11.5l-4.35 5v5.4l-2.8-1.75v-3.65L4 5.25Z" />
      <path d="m15.75 14.75 4.5 4.5m0-4.5-4.5 4.5" />
    </Icon>
  ),
  '⌖': (
    <Icon>
      <circle cx="12" cy="12" r="5.75" />
      <path d="M12 4.5v2.5M12 17v2.5M4.5 12H7M17 12h2.5" />
    </Icon>
  ),
  // ---- comments / review ----
  '🗨': (
    // Square-corner comment bubble with text lines, shared with docs/slides IconComment
    <Icon>
      <path d="M 4.51 4.87 h 14.99 v 10.22 h -8.18 L 7.23 19.18 v -4.09 h -2.73 z" />
      <path d="M 7.91 8.28 h 8.18 M 7.91 11.68 h 5.45" />
    </Icon>
  ),
  '🗑': (
    <Icon>
      <path d="M 5.48 7.05 h 13.05" />
      <path d="M 9.75 7.05 v -0.9 a 1.35 1.35 0 0 1 1.35 -1.35 h 1.8 a 1.35 1.35 0 0 1 1.35 1.35 v 0.9" />
      <path d="m 6.82 7.05 0.81 11.03 a 1.8 1.8 0 0 0 1.8 1.67 h 5.13 a 1.8 1.8 0 0 0 1.8 -1.66 l 0.81 -11.02" />
      <path d="M 10.2 10.65 v 5.4 M 13.8 10.65 v 5.4" />
    </Icon>
  ),
  '✓': (
    <Icon>
      <path d="m4.75 13.75 4.75 5L19.25 5.75" />
    </Icon>
  ),
  '☑': (
    <Icon>
      <rect x="4.75" y="4.75" width="14.5" height="14.5" rx="2" />
      <path d="m8.5 12.5 2.5 2.5 4.75-5.5" />
    </Icon>
  ),
  '⚠': (
    <Icon>
      <path d="M10.7 5.75 4.6 16.5a1.5 1.5 0 0 0 1.3 2.25h12.2a1.5 1.5 0 0 0 1.3-2.25L13.3 5.75a1.5 1.5 0 0 0-2.6 0Z" />
      <path d="M12 9.75v3.75" />
      <circle cx="12" cy="16.15" r="0.4" fill="currentColor" stroke="none" />
    </Icon>
  ),
  '❔': (
    <Icon>
      <path d="M 9.36 9.27 A 2.64 2.64 0 0 1 12 6.72 c 1.46 0 2.64 1.1 2.64 2.55 0 0.97 -0.53 1.58 -1.32 2.2 -0.79 0.62 -1.32 1.1 -1.32 2.11 v 0.4" />
      <circle cx="12" cy="16.84" r="0.44" fill="currentColor" stroke="none" />
      <circle cx="12" cy="12" r="7.26" />
    </Icon>
  ),
  '💡': (
    <Icon>
      <path d="M12 4.25a5.75 5.75 0 0 0-3.25 10.5c.62.43 1 1.05 1 1.78v.72h4.5v-.72c0-.73.38-1.35 1-1.78A5.75 5.75 0 0 0 12 4.25Z" />
      <path d="M10 19.75h4" />
    </Icon>
  ),
  '👓': (
    <Icon>
      <circle cx="7.25" cy="14.25" r="3.25" />
      <circle cx="16.75" cy="14.25" r="3.25" />
      <path d="M10.5 13.75a1.5 1.5 0 0 1 3 0" />
      <path d="M4 14.25c0-2.1.5-4.5 1.5-6.5M20 14.25c0-2.1-.5-4.5-1.5-6.5" />
    </Icon>
  ),
  '✒': (
    <Icon>
      <path d="M13.75 5.5 18.5 10.25 9.5 19.25H4.75V14.5l9-9Z" />
      <path d="m12 7.25 4.75 4.75" />
    </Icon>
  ),
  // ---- charts / analysis ----
  '📊': (
    <Icon>
      <path d="M4.5 19.5h15" />
      <path d="M8 19.5V9.5M12.5 19.5V4.75M17 19.5V12.5" />
    </Icon>
  ),
  '🗠': (
    <Icon>
      <path d="M4.5 4.5v15h15" />
      <path d="m7.5 15.5 3.5-4 2.7 2.5 4.3-5.5" />
    </Icon>
  ),
  '〜': (
    <Icon>
      <path d="m4.5 19 5-7.5 3.5 3.5 6.5-9" />
      <path d="M15.25 6h4.25v4.25" />
    </Icon>
  ),
  '✦': SPARKLE_ICON,
  '✧': SPARKLE_ICON,
  '∑': (
    <Icon>
      <path d="M17.75 7.5V4.75H6.25L12.25 12l-6 7.25h11.5V16.5" />
    </Icon>
  ),
  Σ: (
    <Icon>
      <path d="M17.75 7.5V4.75H6.25L12.25 12l-6 7.25h11.5V16.5" />
    </Icon>
  ),
  '🧮': (
    <Icon>
      <rect x="6.08" y="4.49" width="11.83" height="15.02" rx="1.82" />
      <path d="M 8.59 7.22 h 6.83 v 2.73 h -6.82 z" />
      <circle cx="9.04" cy="13.14" r="0.36" fill="currentColor" stroke="none" />
      <circle cx="12" cy="13.14" r="0.36" fill="currentColor" stroke="none" />
      <circle cx="14.96" cy="13.14" r="0.36" fill="currentColor" stroke="none" />
      <circle cx="9.04" cy="16.32" r="0.36" fill="currentColor" stroke="none" />
      <circle cx="12" cy="16.32" r="0.36" fill="currentColor" stroke="none" />
      <circle cx="14.96" cy="16.32" r="0.36" fill="currentColor" stroke="none" />
    </Icon>
  ),
  // ---- grids / tables / layout ----
  '▦': (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
      <path d="M4.5 9.5h15M4.5 14.5h15M9.5 4.5v15M14.5 4.5v15" />
    </Icon>
  ),
  '⊞': (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
      <path d="M4.5 9.5h15M10.5 9.5v10" />
    </Icon>
  ),
  '▤': (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
      <path d="M4.5 9.5h15M4.5 14.5h15" />
    </Icon>
  ),
  '⊟': (
    <Icon>
      <rect x="4.75" y="4.75" width="14.5" height="14.5" rx="2" />
      <path d="M8.75 12h6.5" />
    </Icon>
  ),
  '⊕': (
    <Icon>
      <circle cx="12" cy="12" r="7.25" />
      <path d="M12 8.5v7M8.5 12h7" />
    </Icon>
  ),
  '⊙': (
    <Icon>
      <circle cx="12" cy="12" r="7.25" />
      <circle cx="12" cy="12" r="1.9" fill="currentColor" stroke="none" />
    </Icon>
  ),
  '⿴': (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
      <rect x="8.25" y="8.25" width="7.5" height="7.5" rx="0.75" />
    </Icon>
  ),
  '⤢': (
    <Icon>
      <rect x="4.75" y="4.75" width="8.75" height="12" rx="1.5" />
      <rect x="9" y="13.5" width="10.25" height="5.75" rx="1.5" />
    </Icon>
  ),
  '▭': (
    <Icon>
      <rect x="6.75" y="4.5" width="10.5" height="15" rx="1.5" />
    </Icon>
  ),
  '⬚': (
    <Icon>
      <rect x="5" y="5" width="14" height="14" rx="1" strokeDasharray="3 2.7" />
    </Icon>
  ),
  '┆': (
    <Icon>
      <path d="M6.5 8.75v-2.5a1.5 1.5 0 0 1 1.5-1.5h8a1.5 1.5 0 0 1 1.5 1.5v2.5" />
      <path d="M6.5 15.25v2.5a1.5 1.5 0 0 0 1.5 1.5h8a1.5 1.5 0 0 0 1.5-1.5v-2.5" />
      <path d="M4.75 12h14.5" strokeDasharray="2.5 2.5" />
    </Icon>
  ),
  '❄': (
    <Icon>
      <path d="M 12 4.64 v 14.72 M 5.62 8.32 l 12.75 7.36 M 18.38 8.32 5.62 15.68" />
      <path d="m 9.7 6.39 2.3 1.84 2.3 -1.84 M 9.7 17.61 l 2.3 -1.84 2.3 1.84" />
    </Icon>
  ),
  // ---- media / objects ----
  '🎨': (
    <Icon>
      <path d="M12 4.5c-4.7 0-8.5 3.4-8.5 7.5s3.8 7.5 8.5 7.5c1.05 0 1.9-.85 1.9-1.9 0-.5-.2-.95-.5-1.3-.3-.35-.5-.8-.5-1.25 0-1.05.85-1.9 1.9-1.9h2.2c2 0 3.5-1.5 3.5-3.4C20.5 7.2 16.7 4.5 12 4.5Z" />
      <circle cx="8.1" cy="10.1" r="0.5" fill="currentColor" stroke="none" />
      <circle cx="11.6" cy="8.1" r="0.5" fill="currentColor" stroke="none" />
      <circle cx="15.4" cy="9.3" r="0.5" fill="currentColor" stroke="none" />
    </Icon>
  ),
  '🖼': (
    <Icon>
      <rect x="3.9" y="4.98" width="16.2" height="14.04" rx="2.16" />
      <circle cx="8.76" cy="9.57" r="1.51" />
      <path d="m 4.44 16.86 4.21 -4.21 3.51 3.51 3.62 -3.62 3.78 3.78" />
    </Icon>
  ),
  '◇': (
    <Icon>
      <rect x="4.75" y="4.75" width="9.5" height="9.5" rx="1.5" />
      <circle cx="15" cy="15" r="4.25" />
    </Icon>
  ),
  '🔗': (
    <Icon>
      <path d="M 9.6 16 H 8 A 4 4 0 0 1 8 8 h 1.6" />
      <path d="M 14.4 8 h 1.6 a 4 4 0 1 1 0 8 h -1.6" />
      <path d="M 9 12 h 6" />
    </Icon>
  ),
  '🗎': (
    <Icon>
      <path d="M 13.37 4.49 H 7.68 c -0.76 0 -1.36 0.61 -1.36 1.37 v 12.29 c 0 0.76 0.61 1.37 1.37 1.37 h 8.65 c 0.76 0 1.37 -0.61 1.37 -1.36 V 8.82 L 13.37 4.49 Z" />
      <path d="M 13.37 4.49 V 8.82 h 4.32" />
    </Icon>
  ),
  '🗒': (
    <Icon>
      <path d="M 13.37 4.49 H 7.68 c -0.76 0 -1.36 0.61 -1.36 1.37 v 12.29 c 0 0.76 0.61 1.37 1.37 1.37 h 8.65 c 0.76 0 1.37 -0.61 1.37 -1.36 V 8.82 L 13.37 4.49 Z" />
      <path d="M 13.37 4.49 V 8.82 h 4.32" />
      <path d="M 9.27 12.46 h 5.46 M 9.27 15.19 h 3.64" />
    </Icon>
  ),
  '🕮': (
    <Icon>
      <path d="M12 6.25c-1.8-1.3-4-1.9-7.25-1.75V17.75c3.25-.15 5.45.45 7.25 1.75 1.8-1.3 4-1.9 7.25-1.75V4.5C16 4.35 13.8 4.95 12 6.25Z" />
      <path d="M12 6.25V19.5" />
    </Icon>
  ),
  '🛢': (
    <Icon>
      <ellipse cx="12" cy="6.5" rx="7" ry="2.5" />
      <path d="M 5.49 6.89 v 10.23 c 0 1.3 2.91 2.33 6.51 2.33 s 6.51 -1.02 6.51 -2.33 v -10.23" />
      <path d="M 5.49 12 c 0 1.3 2.91 2.33 6.51 2.33 s 6.51 -1.02 6.51 -2.33" />
    </Icon>
  ),
  '🌐': (
    <Icon>
      <circle cx="12" cy="12" r="7.5" />
      <path d="M12 4.5c2.5 2 3.75 4.5 3.75 7.5S14.5 17.5 12 19.5c-2.5-2-3.75-4.5-3.75-7.5S9.5 6.5 12 4.5Z" />
      <path d="M4.75 12h14.5" />
    </Icon>
  ),
  '🏷': (
    <Icon>
      <path d="M12.6 4.75H7a2.25 2.25 0 0 0-2.25 2.25v5.6c0 .6.24 1.17.66 1.59l5.9 5.9a2.25 2.25 0 0 0 3.18 0l4.86-4.86a2.25 2.25 0 0 0 0-3.18l-5.9-5.9a2.25 2.25 0 0 0-1.59-.66Z" />
      <circle cx="9" cy="9" r="1.2" />
    </Icon>
  ),
  '🔒': (
    <Icon>
      <rect x="5.75" y="10.5" width="12.5" height="9" rx="2" />
      <path d="M8.75 10.5V8a3.25 3.25 0 0 1 6.5 0v2.5" />
      <circle cx="12" cy="15" r="1" fill="currentColor" stroke="none" />
    </Icon>
  ),
  '🔐': (
    <Icon>
      <rect x="5.75" y="10.5" width="12.5" height="9" rx="2" />
      <path d="M8.75 10.5V8a3.25 3.25 0 0 1 6.5 0v2.5" />
      <circle cx="12" cy="15" r="1" fill="currentColor" stroke="none" />
    </Icon>
  ),
  '▶': (
    <Icon>
      <path d="M 7.8 4.5 v 15 L 20.1 12 7.8 4.5 Z" />
    </Icon>
  ),
  '⚡': (
    <Icon>
      <path d="M13 4.25 6.5 13.25h4.5L11 19.75l6.5-9h-4.5L13 4.25Z" />
    </Icon>
  ),
  '♿': (
    <Icon>
      <circle cx="12" cy="6.15" r="1.8" />
      <path d="M 5.48 9.75 c 4.32 1.17 8.73 1.17 13.05 0" />
      <path d="M 12 10.65 v 3.38 l -3.06 4.95 M 12 14.03 l 3.06 4.95" />
    </Icon>
  ),
  '🕘': CLOCK_ICON,
  '🕐': CLOCK_ICON,
  '🕒': CLOCK_ICON,
  '⟳': REFRESH_ICON,
  '↻': REFRESH_ICON,
  '✕': (
    <Icon>
      <path d="m6 6 12 12M18 6 6 18" />
    </Icon>
  ),
  // ---- sort / move ----
  '↓': (
    <Icon>
      <path d="M8 5v13.5M4.75 15.25 8 18.5l3.25-3.25" />
      <path d="M13.5 6.5h6M13.5 10.5h4.5M13.5 14.5h3" />
    </Icon>
  ),
  '↑': (
    <Icon>
      <path d="M8 19V5.5M4.75 8.75 8 5.5l3.25 3.25" />
      <path d="M13.5 9.5h3M13.5 13.5h4.5M13.5 17.5h6" />
    </Icon>
  ),
  '⇄': (
    <Icon>
      <path d="M4.75 8.25h14.5M16 5l3.25 3.25L16 11.5" />
      <path d="M19.25 15.75H4.75M8 12.5l-3.25 3.25L8 19" />
    </Icon>
  ),
  '⇅': (
    <Icon>
      <path d="M8.5 19V5.5M5.25 8.75 8.5 5.5l3.25 3.25" />
      <path d="M15.5 5v13.5M12.25 15.25l3.25 3.25L18.75 15.25" />
    </Icon>
  ),
  '⇶': (
    <Icon>
      <path d="M 5.48 7.05 h 10.35 M 13.35 4.57 l 2.93 2.48 -2.92 2.48" />
      <path d="M 5.48 12 h 10.35 M 13.35 9.53 16.28 12 l -2.92 2.48" />
      <path d="M 5.48 16.95 h 10.35 M 13.35 14.48 l 2.93 2.48 -2.92 2.48" />
    </Icon>
  ),
  // ---- text ----
  A: (
    <Icon>
      <path d="M5.75 7V4.75h12.5V7M12 4.75v14.5M9.5 19.25h5" />
    </Icon>
  ),
  // ---- windows / panes ----
  '🗔': (
    <Icon>
      <rect x="5.17" y="5.63" width="13.65" height="12.74" rx="1.37" />
      <path d="M 5.17 9.27 h 13.65" />
    </Icon>
  ),
  '◫': (
    <Icon>
      <rect x="4.5" y="5" width="15" height="14" rx="1.5" />
      <path d="M12 5v14" />
    </Icon>
  ),
  '▥': (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
      <path d="M9.5 4.5v15M14.5 4.5v15" />
    </Icon>
  ),
  '⬡': (
    <Icon>
      <path d="M12 4.25 18.75 8.1v7.8L12 19.75 5.25 15.9V8.1L12 4.25Z" />
    </Icon>
  ),
  // ---- editing arrows ----
  '⌫': (
    <Icon>
      <path d="M9.25 5.5h9a1.5 1.5 0 0 1 1.5 1.5v10a1.5 1.5 0 0 1-1.5 1.5h-9L4.25 12l5-6.5Z" />
      <path d="m11.75 9.75 4.5 4.5m0-4.5-4.5 4.5" />
    </Icon>
  ),
  '⇢': (
    <Icon>
      <path d="M4.75 12h11.5" strokeDasharray="2.5 2.5" />
      <path d="M15.5 8.25 19.25 12l-3.75 3.75" />
    </Icon>
  ),
  '⇠': (
    <Icon>
      <path d="M19.25 12H7.75" strokeDasharray="2.5 2.5" />
      <path d="M8.5 8.25 4.75 12l3.75 3.75" />
    </Icon>
  ),
  '→': (
    <Icon>
      <path d="M4.75 12h14.5M15.5 8.25 19.25 12l-3.75 3.75" />
    </Icon>
  ),
  '←': (
    <Icon>
      <path d="M19.25 12H4.75M8.5 8.25 4.75 12l3.75 3.75" />
    </Icon>
  ),
  '↔': (
    <Icon>
      <path d="M4.75 12h14.5M8 8.75 4.75 12 8 15.25M16 8.75 19.25 12 16 15.25" />
    </Icon>
  ),
  '↕': (
    <Icon>
      <path d="M12 4.75v14.5M8.75 8 12 4.75 15.25 8M8.75 16 12 19.25 15.25 16" />
    </Icon>
  ),
  'A↑': (
    <Icon>
      <path d="M3.25 19 8.5 5.25 13.75 19M5.1 14.25h6.8" />
      <path d="M18 17.5V6.75M14.9 9.85 18 6.75l3.1 3.1" />
    </Icon>
  ),
  'A↓': (
    <Icon>
      <path d="M3.25 19 8.5 5.25 13.75 19M5.1 14.25h6.8" />
      <path d="M18 6.75V17.5M14.9 14.4l3.1 3.1 3.1-3.1" />
    </Icon>
  ),
  // ---- home formatting (borders / fill / merge / orientation) ----
  '⊡': (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" strokeDasharray="2.6 2.3" />
      <path d="M12 4.5v15M4.5 12h15" />
    </Icon>
  ),
  '◧': (
    <Icon>
      <path d="M11.25 4.75 4.6 11.4a1.5 1.5 0 0 0 0 2.1l4.9 4.9a1.5 1.5 0 0 0 2.1 0l6.65-6.65z" />
      <path d="M19.4 13.9c.75 1 1.15 1.85 1.15 2.5a1.45 1.45 0 0 1-2.9 0c0-.65.4-1.5 1.15-2.5Z" />
    </Icon>
  ),
  '∅': (
    <Icon>
      <rect x="4.75" y="4.75" width="14.5" height="14.5" rx="2" />
      <path d="m6.5 17.5 11-11" />
    </Icon>
  ),
  '⇔': (
    <Icon>
      <rect x="4.5" y="4.75" width="15" height="14.5" rx="1.5" />
      <path d="M8 12h8" />
      <path d="M10.25 9.75 8 12l2.25 2.25M13.75 9.75 16 12l-2.25 2.25" />
    </Icon>
  ),
  '⤴': (
    <Icon>
      <text
        x="2.5"
        y="17.5"
        fontSize="12"
        fill="currentColor"
        stroke="none"
        fontFamily="Segoe UI, sans-serif"
      >
        ab
      </text>
      <path d="m15.75 11.75 5.25-5.25" />
      <path d="M17.25 6.5H21v3.75" />
    </Icon>
  ),
}

/** Genspark brand mark (rounded-square sparkle badge), inline so it renders
 * crisply at device resolution instead of going through <img> rasterization */
export function GensparkMark({ size = 18 }: { readonly size?: number }): ReactElement {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 130 130.025"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      aria-hidden
    >
      <path
        d="M105.115 0H24.6428C11.0443 0 0 11.0686 0 24.6915V105.334C0 118.981 11.0199 130.025 24.6428 130.025H105.115C118.714 130.025 129.758 118.957 129.758 105.334V24.6915C129.758 11.0443 118.714 0 105.115 0ZM71.5201 35.2735C85.5078 33.1571 86.7729 31.9164 88.865 17.88C88.938 17.4421 89.3028 17.1259 89.7407 17.1259C90.1786 17.1259 90.5435 17.4421 90.6164 17.88C92.7328 31.8921 93.9735 33.1571 107.961 35.2735C108.399 35.3465 108.715 35.7114 108.715 36.1493C108.715 36.5871 108.399 36.952 107.961 37.025C93.9249 39.1414 92.7085 40.4064 90.5677 54.6131C90.5191 54.9537 90.2516 55.197 89.911 55.197C89.5704 55.197 89.3028 54.9537 89.2542 54.6131C87.1134 40.4064 85.5565 39.1658 71.4958 37.025C71.0579 36.952 70.7417 36.5871 70.7417 36.1493C70.7417 35.7114 71.0579 35.3465 71.4958 35.2735H71.5201ZM101.758 78.5261C101.758 78.8181 101.563 79.037 101.271 79.0856C92.3193 80.4236 91.5652 81.2264 90.2029 90.2759C90.1786 90.4948 89.9839 90.6408 89.7893 90.6408C89.5703 90.6408 89.4001 90.4948 89.3758 90.2759C88.0135 81.2507 87.0161 80.4479 78.0883 79.0856C77.7964 79.037 77.6017 78.7937 77.6017 78.5261C77.6017 78.2342 77.7964 78.0153 78.0883 77.9666C86.9918 76.6287 87.7703 75.8259 89.1326 66.898C89.1812 66.6061 89.4244 66.4115 89.692 66.4115C89.9839 66.4115 90.2028 66.6061 90.2515 66.898C91.5894 75.8259 92.3923 76.6043 101.296 77.9666C101.588 78.0153 101.782 78.2585 101.782 78.5261H101.758ZM16.5178 54.8077C16.5178 54.1023 17.0286 53.4941 17.7341 53.3968C40.1388 50.0154 42.1093 47.9963 45.4907 25.5672C45.588 24.8861 46.1961 24.3509 46.9016 24.3509C47.6071 24.3509 48.191 24.8617 48.3126 25.5672C51.694 47.9963 53.6887 50.0154 76.0691 53.3968C76.7503 53.4941 77.2855 54.1023 77.2855 54.8077C77.2855 55.5132 76.7746 56.1214 76.0691 56.2187C53.5914 59.6244 51.6696 61.6192 48.2639 84.3645C48.1909 84.8754 47.7287 85.2889 47.2179 85.2889C46.707 85.2889 46.2448 84.8997 46.1718 84.3645C42.7418 61.6435 40.2604 59.6244 17.7584 56.2187C17.0772 56.1214 16.542 55.5132 16.542 54.8077H16.5178ZM112.097 109.591C112.097 111.416 110.613 112.9 108.813 112.9H21.2614C19.4369 112.9 17.9774 111.416 17.9774 109.591V102.658C17.9774 100.834 19.4612 99.3497 21.2614 99.3497H108.813C110.637 99.3497 112.097 100.834 112.097 102.658V109.591Z"
        fill="currentColor"
      />
    </svg>
  )
}
