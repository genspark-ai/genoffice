/** Small monochrome SVG icons approximating Word's ribbon glyphs. */

import type { ReactNode } from 'react'
import type { AnimEffectKind } from '../../shared/ipc'

interface IconProps {
  size?: number
}

/** Constant painted stroke instead of proportional scaling: ~1.5px lines on
 *  20px+ glyphs, ~1.25px on the 13-19px ones, ~1.1px below (a proportional
 *  1.5-unit stroke would paint 1.75px at 28px and hairlines at small sizes).
 *  stroke-width is in 24-canvas units: units = painted-px × 24 / rendered-px. */
function pinnedStroke(size: number): number {
  const painted = size >= 20 ? 1.5 : size >= 13 ? 1.25 : 1.1
  return (painted * 24) / size
}

function Svg({ size = 24, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={pinnedStroke(size)}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

/** Renders gallery glyph markup (inner SVG for a 24×24 viewBox, stroke=currentColor). */
export function IconGlyph({ body, size = 18 }: { body: string; size?: number }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={pinnedStroke(size)}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
      dangerouslySetInnerHTML={{ __html: body }}
    />
  )
}

export function IconFind(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="10.13" cy="10.13" r="5.61" />
      <path d="M 14.41 14.41 L 19.48 19.48" />
    </Svg>
  )
}

export function IconBullets(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="5.48" cy="6.67" r="1.3" fill="currentColor" stroke="none" />
      <circle cx="5.48" cy="12" r="1.3" fill="currentColor" stroke="none" />
      <circle cx="5.48" cy="17.33" r="1.3" fill="currentColor" stroke="none" />
      <path d="M 9.63 6.67 h 9.48 M 9.63 12 h 9.48 M 9.63 17.33 h 9.48" />
    </Svg>
  )
}

export function IconNumbered(props: IconProps) {
  return (
    <Svg {...props}>
      <text
        x="1.5"
        y="8.1"
        fontSize="8.1"
        fill="currentColor"
        stroke="none"
        fontFamily="Segoe UI, sans-serif"
      >
        1
      </text>
      <text
        x="1.5"
        y="15.6"
        fontSize="8.1"
        fill="currentColor"
        stroke="none"
        fontFamily="Segoe UI, sans-serif"
      >
        2
      </text>
      <text
        x="1.5"
        y="23.1"
        fontSize="8.1"
        fill="currentColor"
        stroke="none"
        fontFamily="Segoe UI, sans-serif"
      >
        3
      </text>
      <path d="M9.75 5.25 h12 M9.75 12.75 h12 M9.75 20.25 h12" />
    </Svg>
  )
}

export function IconIndentDec(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.53 5.15 h 14.94 M 12 9.26 h 7.47 M 12 12.62 h 7.47 M 12 15.98 h 7.47 M 4.53 19.47 h 14.94" />
      <path d="M 8.51 9.26 4.78 12.62 l 3.74 3.36 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconIndentInc(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.53 5.15 h 14.94 M 12 9.26 h 7.47 M 12 12.62 h 7.47 M 12 15.98 h 7.47 M 4.53 19.47 h 14.94" />
      <path d="M 4.78 9.26 l 3.74 3.36 -3.73 3.36 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconAlignLeft(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.53 5.78 h 14.94 M 4.53 9.51 h 9.96 M 4.53 13.25 h 14.94 M 4.53 16.98 h 9.96" />
    </Svg>
  )
}

export function IconAlignCenter(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.53 5.78 h 14.94 M 7.02 9.51 h 9.96 M 4.53 13.25 h 14.94 M 7.02 16.98 h 9.96" />
    </Svg>
  )
}

export function IconAlignRight(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.53 5.78 h 14.94 M 9.51 9.51 h 9.96 M 4.53 13.25 h 14.94 M 9.51 16.98 h 9.96" />
    </Svg>
  )
}

export function IconAlignJustify(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.53 5.78 h 14.94 M 4.53 9.51 h 14.94 M 4.53 13.25 h 14.94 M 4.53 16.98 h 14.94" />
    </Svg>
  )
}

export function IconLineSpacing(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 12 5.85 h 7.38 M 12 10.03 h 7.38 M 12 14.21 h 7.38 M 12 18.4 h 7.38" />
      <path d="M 6.47 6.1 v 11.81 M 4.37 8.56 l 2.09 -2.46 2.09 2.46 M 4.37 15.44 l 2.09 2.46 2.09 -2.46" />
    </Svg>
  )
}

export function IconClearFormat(props: IconProps) {
  return (
    <Svg {...props}>
      {/* letter A with a wiped-off stroke at its top left */}
      <path d="M3.75 18.75 8.25 6l4.5 12.75" />
      <path d="M5.55 14.25h5.4" />
      <path d="M4.35 8.85l1.8-1.8" />
      {/* compact diagonal eraser at the lower right (the old diamond's spot), outline only, band facing the A */}
      <g transform="rotate(45 17.4 17.4)">
        <rect
          x="13.65"
          y="14.7"
          width="7.5"
          height="5.4"
          rx="0.75"
          stroke="var(--ribbon-accent-2, #A33FB5)"
        />
        <path d="M15.75 14.7v5.4" stroke="var(--ribbon-accent-2, #A33FB5)" />
      </g>
    </Svg>
  )
}

export function IconGrowFont(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.25 19 8.5 5.25 13.75 19M5.1 14.25h6.8" />
      <path d="M18 17.5V6.75M14.9 9.85 18 6.75l3.1 3.1" />
    </Svg>
  )
}

export function IconShrinkFont(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.25 19 8.5 5.25 13.75 19M5.1 14.25h6.8" />
      <path d="M18 6.75V17.5M14.9 14.4l3.1 3.1 3.1-3.1" />
    </Svg>
  )
}

/** Character spacing (MS-style): AV above a double-headed arrow */
export function IconCharSpacing(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={4} y={13.5} s={13}>
        AV
      </TextGlyph>
      <path
        d="M4.5 18.75 h15 M7.2 16.05 4.5 18.75 l2.7 2.7 M16.8 16.05 19.5 18.75 l-2.7 2.7"
        stroke="var(--ribbon-accent, #2B7CD3)"
      />
    </Svg>
  )
}

/** MS-style text highlighter: marker nib only; the color bar is rendered by the button */
export function IconTextHighlight(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M4.5 16.5 14.25 6.75 a2.1 2.1 0 0 1 3 0 l0.75 0.75 a2.1 2.1 0 0 1 0 3 L8.25 20.25 H4.5 z" />
    </Svg>
  )
}

export function IconHighlight(props: IconProps) {
  return (
    <Svg {...props}>
      <path
        d="M 4.95 15.38 14.5 5.83 a 2.06 2.06 0 0 1 2.94 0 l 0.73 0.73 a 2.06 2.06 0 0 1 0 2.94 L 8.62 19.05 H 4.95 z"
        fill="none"
      />
      <rect
        x="3.78"
        y="17.87"
        width="5.87"
        height="2.35"
        rx="1.17"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

/* ---------- shared shapes ---------- */

/** page outline used by many icons */
const PAGE = <path d="M6 2.25 h9 l3.75 3.75 v15.75 h-12.75 z" />

function TextGlyph({
  x,
  y,
  s,
  children,
  bold,
}: {
  x: number
  y: number
  s: number
  children: string
  bold?: boolean
}) {
  return (
    <text
      x={x}
      y={y}
      fontSize={s}
      fill="currentColor"
      stroke="none"
      fontFamily="Segoe UI, sans-serif"
      fontWeight={bold ? 700 : 400}
    >
      {children}
    </text>
  )
}

/* ---------- clipboard (Home) ---------- */

export function IconPaste(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 8.14 18.16 H 5.89 C 4.85 18.16 4 17.31 4 16.26 V 6.79 C 4 5.74 4.85 4.89 5.89 4.89 H 7.32 M 17.26 7.74 V 6.79 C 17.26 5.74 16.41 4.89 15.37 4.89 H 13.95" />
      <rect x="7.79" y="3" width="5.68" height="2.84" rx="0.95" />
      <path d="M 18.21 7.74 H 9.68 C 8.64 7.74 7.79 8.59 7.79 9.63 V 19.11 C 7.79 20.15 8.64 21 9.68 21 H 15.49 L 20.11 16.03 V 9.63 C 20.11 8.59 19.26 7.74 18.21 7.74 Z" />
      <path d="M 10.63 11.05 H 17.26 M 10.63 14.37 H 14.42" />
      <path d="M 15.37 21 V 16.74 C 15.37 16.21 15.79 15.79 16.32 15.79 H 20.11" />
    </Svg>
  )
}

export function IconCut(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 7.91 19.27 L 8.92 17.57 L 16.98 3.59 M 7.02 3.5 L 15.08 17.47 L 16.09 19.27" />
      <circle cx="5.83" cy="18.13" r="2.37" />
      <circle cx="18.17" cy="18.13" r="2.37" />
    </Svg>
  )
}

export function IconCopy(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="7" y="7" width="14" height="14" rx="3" />
      <path d="M 14.5 4 H 7 C 5.34 4 4 5.34 4 7 V 14.5" />
    </Svg>
  )
}

export function IconFormatPainter(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 16.5 4 H 3.97 V 10.71 H 16.5 Z" />
      <path d="M 16.51 7.13 H 20.53 V 14.33 L 9.79 15.72 V 21" />
    </Svg>
  )
}

/** Slide layout: slide frame + title line + two content placeholders */
export function IconSlideLayout(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="5.15" width="14.94" height="13.69" rx="1" />
      <path d="M 7.64 8.89 h 8.72" />
      <rect x="7.64" y="11.38" width="3.74" height="4.36" rx="0.5" />
      <rect x="12.62" y="11.38" width="3.74" height="4.36" rx="0.5" />
    </Svg>
  )
}

/* ---------- Insert ---------- */

export function IconTable(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="5.15" width="14.94" height="13.69" rx="1" />
      <path d="M 4.53 9.76 h 14.94 M 4.53 14.37 h 14.94 M 9.51 5.15 v 13.69 M 14.49 5.15 v 13.69" />
    </Svg>
  )
}

export function IconPicture(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="5.78" width="14.94" height="12.45" rx="1" />
      <circle cx="8.76" cy="9.76" r="1.37" />
      <path d="M 5.15 16.98 10.13 12 l 3.74 3.74 2.49 -2.49 2.49 2.49" />
    </Svg>
  )
}

export function IconShapes(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="9.42" cy="9.42" r="4.64" />
      <rect x="11.36" y="11.36" width="8.39" height="8.39" rx="1.03" fill="var(--surface, #fff)" />
    </Svg>
  )
}

export function IconLink(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 10.35 13.65 13.65 10.35" />
      <path d="M 11.31 7.59 13.38 5.52 a 3.59 3.59 0 0 1 5.1 5.1 L 16.41 12.69" />
      <path d="M 12.69 16.41 10.62 18.48 a 3.59 3.59 0 0 1 -5.1 -5.1 l 2.07 -2.07" />
    </Svg>
  )
}

export function IconComment(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.51 4.87 h 14.99 v 10.22 h -8.18 L 7.23 19.18 v -4.09 h -2.73 z" />
      <path d="M 7.91 8.28 h 8.18 M 7.91 11.68 h 5.45" />
    </Svg>
  )
}

export function IconPageBreak(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 7.38 4.49 h 9.24 v 5.2 M 7.38 4.49 v 5.2 M 7.38 19.51 h 9.24 v -5.2 M 7.38 19.51 v -5.2" />
      <path d="M 4.49 12 h 2.31 M 8.54 12 h 2.31 M 12.58 12 h 2.31 M 16.62 12 h 2.89" />
    </Svg>
  )
}

export function IconHeader(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="6.23" y="4.49" width="11.55" height="15.02" rx="0.92" />
      <path d="M 7.96 7.38 h 8.09 M 7.96 9.46 h 8.09" opacity="0.9" />
    </Svg>
  )
}

export function IconFooter(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="6.23" y="4.49" width="11.55" height="15.02" rx="0.92" />
      <path d="M 7.96 14.54 h 8.09 M 7.96 16.62 h 8.09" opacity="0.9" />
    </Svg>
  )
}

export function IconPageNumber(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="6.25" y="4.5" width="11.5" height="15" rx="0.95" />
      <TextGlyph x={9.2} y={15.5} s={8}>
        #
      </TextGlyph>
    </Svg>
  )
}

export function IconSymbol(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={3.4} y={20.4} s={24}>
        Ω
      </TextGlyph>
    </Svg>
  )
}

export function IconEquation(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={5.4} y={17.75} s={22}>
        π
      </TextGlyph>
    </Svg>
  )
}

/* ---------- Design ---------- */

export function IconSlideMaster(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.2" y="4.8" width="12" height="8.4" rx="1.2" />
      <path d="M 6.6 7.8 h 7.2 M 6.6 10.2 h 4.8" />
      <rect x="9.6" y="15" width="9.6" height="4.8" rx="0.96" />
    </Svg>
  )
}

export function IconTheme(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={2.25} y={17.25} s={16.5}>
        A
      </TextGlyph>
      <TextGlyph x={12.75} y={17.25} s={12}>
        a
      </TextGlyph>
      <rect
        x="3.75"
        y="19.35"
        width="16.5"
        height="2.7"
        rx="1.35"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconThemeFonts(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={3} y={18} s={16.5}>
        F
      </TextGlyph>
      <path d="M14.25 18 18 6.75 21.75 18 M15.45 14.4 h5.1" />
    </Svg>
  )
}

export function IconThemeColors(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="7.32" cy="7.63" r="2.81" />
      <circle cx="16.68" cy="7.63" r="2.81" />
      <circle cx="7.32" cy="16.32" r="2.81" />
      <circle cx="16.68" cy="16.32" r="2.81" fill="currentColor" />
    </Svg>
  )
}

export function IconPageColor(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 12.84 5.04 6.6 11.28 a 1.56 1.56 0 0 0 0 2.16 l 3.96 3.96 a 1.56 1.56 0 0 0 2.16 0 l 6.24 -6.24 z" />
      <path d="M 12.84 5.04 10.8 7.2" />
      <path
        d="M 18.72 15.12 s 1.68 2.04 1.68 3.24 a 1.68 1.68 0 0 1 -3.36 0 c 0 -1.2 1.68 -3.24 1.68 -3.24 z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconWatermark(props: IconProps) {
  return (
    <Svg {...props}>
      {PAGE}
      <path d="M7.8 17.25 16.2 7.5" opacity="0.45" />
    </Svg>
  )
}

export function IconPageBorders(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="4.53" width="14.94" height="14.94" rx="1" />
      <rect x="7.52" y="7.52" width="8.96" height="8.96" />
    </Svg>
  )
}

/* ---------- Layout ---------- */

export function IconMargins(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="6.23" y="4.49" width="11.55" height="15.02" rx="0.92" />
      <rect x="8.77" y="7.03" width="6.47" height="9.93" strokeDasharray="2.4 2.1" />
    </Svg>
  )
}

export function IconOrientation(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.8" y="7.2" width="9" height="12" rx="0.96" />
      <rect x="9" y="11.4" width="10.8" height="7.8" rx="0.96" fill="var(--surface, #fff)" />
      <path d="M 15.6 5.04 a 6 6 0 0 1 3.6 3.12 M 19.2 5.4 v 3 h -3" />
    </Svg>
  )
}

export function IconPageSize(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="6.23" y="4.49" width="11.55" height="15.02" rx="0.92" />
      <path d="M 9.11 12 h 5.78 M 12 9.11 v 5.78 M 10.5 10.5 9.11 9.11 m 5.78 0 -1.39 1.39 m 0 3 1.39 1.39 m -5.78 0 1.39 -1.39" />
    </Svg>
  )
}

export function IconColumns(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.47 5.15 h 6.16 M 4.47 8.58 h 6.16 M 4.47 12 h 6.16 M 4.47 15.42 h 6.16 M 4.47 18.85 h 6.16" />
      <path d="M 13.37 5.15 h 6.16 M 13.37 8.58 h 6.16 M 13.37 12 h 6.16 M 13.37 15.42 h 6.16 M 13.37 18.85 h 6.16" />
    </Svg>
  )
}

/* ---------- References ---------- */

export function IconToc(props: IconProps) {
  return (
    <Svg {...props}>
      {PAGE}
      <path d="M8.25 7.5 h7.5 M10.05 10.95 h5.7 M10.05 14.4 h5.7 M8.25 17.85 h7.5" />
    </Svg>
  )
}

export function IconRefresh(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 19.02 9.98 a 7.29 7.29 0 0 0 -13.5 -1.62 M 4.98 14.03 a 7.29 7.29 0 0 0 13.5 1.62" />
      <path d="M 19.43 4.57 v 4.05 h -4.05 M 4.57 19.43 v -4.05 h 4.05" />
    </Svg>
  )
}

export function IconFootnote(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={2.25} y={18} s={13.5}>
        AB
      </TextGlyph>
      <TextGlyph x={17.25} y={12} s={10.5} bold>
        1
      </TextGlyph>
    </Svg>
  )
}

export function IconEndnote(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={2.25} y={18} s={13.5}>
        AB
      </TextGlyph>
      <TextGlyph x={16.95} y={12} s={10.5} bold>
        n
      </TextGlyph>
    </Svg>
  )
}

export function IconCitation(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={3} y={18.75} s={21} bold>
        “”
      </TextGlyph>
    </Svg>
  )
}

export function IconBook(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 12 6.13 C 10.43 4.95 7.82 4.43 4.82 4.69 v 13.57 c 3 -0.26 5.61 0.26 7.18 1.44 1.57 -1.17 4.18 -1.7 7.18 -1.44 V 4.69 c -3 -0.26 -5.61 0.26 -7.18 1.44 z" />
      <path d="M 12 6.13 v 13.57" />
    </Svg>
  )
}

export function IconCaption(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.53 7.64 h 10.58 L 19.47 12 l -4.36 4.36 H 4.53 z" />
      <circle cx="8.27" cy="12" r="1.12" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconIndex(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={2.7} y={9.75} s={9.75}>
        A
      </TextGlyph>
      <TextGlyph x={2.7} y={20.25} s={9.75}>
        B
      </TextGlyph>
      <path d="M12 6.75 h9 M12 12 h9 M12 17.25 h9" />
    </Svg>
  )
}

/* ---------- Review ---------- */

export function IconWordCount(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={2.4} y={12} s={12}>
        123
      </TextGlyph>
      <path d="M3 16.5 h18 M3 20.25 h12" />
    </Svg>
  )
}

export function IconSpellcheck(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={2.1} y={12.75} s={11.25}>
        abc
      </TextGlyph>
      <path d="M9 17.25 12.75 20.25 19.5 11.25" />
    </Svg>
  )
}

export function IconSparkle(props: IconProps) {
  return (
    <Svg {...props}>
      <path
        d="M 12 4.55 C 12 8.67 15.33 12 19.45 12 C 15.33 12 12 15.33 12 19.45 C 12 15.33 8.67 12 4.55 12 C 8.67 12 12 8.67 12 4.55 Z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconWand(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.4 18.6 14.4 9.6" />
      <path
        d="M 16.56 4.2 l 0.84 2.28 2.28 0.84 -2.28 0.84 -0.84 2.28 -0.84 -2.28 -2.28 -0.84 2.28 -0.84 z"
        fill="currentColor"
        stroke="none"
      />
      <path
        d="M 18.6 12.6 l 0.48 1.32 1.32 0.48 -1.32 0.48 -0.48 1.32 -0.48 -1.32 -1.32 -0.48 1.32 -0.48 z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconTranslate(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={1.8} y={13.5} s={12.75}>
        文
      </TextGlyph>
      <path d="M13.2 20.25 17.25 9.75 21.3 20.25 M14.55 16.95 h5.4" />
    </Svg>
  )
}

export function IconTrackChanges(props: IconProps) {
  return (
    <Svg {...props}>
      {PAGE}
      <path d="M8.25 8.25 h7.5 M8.25 12 h4.5" />
      <path d="M12.75 19.8 20.4 12.15 l1.8 1.8 -7.65 7.65 -2.7 0.9 z" fill="var(--surface, #fff)" />
    </Svg>
  )
}

export function IconAccept(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.75 12.75 9 18 l11.25 -12" />
    </Svg>
  )
}

export function IconReject(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.75 3.75 20.25 20.25 M20.25 3.75 l-16.5 16.5" />
    </Svg>
  )
}

export function IconCompare(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="6.23" width="6.35" height="11.55" rx="0.92" />
      <rect x="13.16" y="6.23" width="6.35" height="11.55" rx="0.92" />
      <path d="M 9.69 12 h 4.62 M 12.69 10.38 14.31 12 l -1.62 1.62" />
    </Svg>
  )
}

export function IconLock(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="6.4" y="10.76" width="11.21" height="9.34" rx="1.24" />
      <path d="M 8.89 10.76 V 8.27 a 3.11 3.11 0 0 1 6.23 0 v 2.49" />
      <circle cx="12" cy="15.11" r="1.24" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/* ---------- View ---------- */

function Magnifier({ children }: { children?: ReactNode }) {
  return (
    <>
      <circle cx="10.5" cy="10.5" r="7.2" />
      <path d="M15.9 15.9 21 21" />
      {children}
    </>
  )
}

export function IconZoomOut(props: IconProps) {
  return (
    <Svg {...props}>
      <Magnifier>
        <path d="M7.2 10.5 h6.6" />
      </Magnifier>
    </Svg>
  )
}

export function IconZoomIn(props: IconProps) {
  return (
    <Svg {...props}>
      <Magnifier>
        <path d="M7.2 10.5 h6.6 M10.5 7.2 v6.6" />
      </Magnifier>
    </Svg>
  )
}

export function IconZoom100(props: IconProps) {
  return (
    <Svg {...props}>
      <Magnifier>
        <TextGlyph x={5.7} y={13.5} s={7.5} bold>
          1:1
        </TextGlyph>
      </Magnifier>
    </Svg>
  )
}

export function IconPageWidth(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="4.53" width="14.94" height="14.94" rx="1" />
      <path d="M 7.02 12 h 9.96 M 9.26 9.76 7.02 12 l 2.24 2.24 M 14.74 9.76 16.98 12 l -2.24 2.24" />
    </Svg>
  )
}

export function IconWholePage(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="6.23" y="4.49" width="11.55" height="15.02" rx="0.92" />
      <path d="M 9.11 12 h 5.78 M 12 9.11 v 5.78" />
    </Svg>
  )
}

export function IconAiPanel(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="15.02" height="12.71" rx="0.92" />
      <path d="M 14.08 5.65 v 12.71" />
      <path
        d="M 15.47 9.92 l 0.58 1.5 1.5 0.58 -1.5 0.58 -0.58 1.5 -0.58 -1.5 -1.5 -0.58 1.5 -0.58 z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconMoon(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 18.79 14.35 A 7.57 7.57 0 0 1 9.65 5.21 a 7.57 7.57 0 1 0 9.14 9.14 z" />
    </Svg>
  )
}

export function IconReadMode(props: IconProps) {
  return <IconBook {...props} />
}

export function IconOutlineView(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="5.48" cy="6.13" r="1.31" fill="currentColor" stroke="none" />
      <path d="M 8.74 6.13 h 10.44" />
      <circle cx="8.74" cy="12" r="1.31" fill="currentColor" stroke="none" />
      <path d="M 12 12 h 7.18" />
      <circle cx="8.74" cy="17.87" r="1.31" fill="currentColor" stroke="none" />
      <path d="M 12 17.87 h 7.18" />
    </Svg>
  )
}

export function IconRuler(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="9.11" width="15.02" height="5.78" rx="0.92" />
      <path d="M 7.96 9.11 v 2.31 M 10.85 9.11 v 3.47 M 13.73 9.11 v 2.31 M 16.62 9.11 v 3.47" />
    </Svg>
  )
}

export function IconNavPane(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="15.02" height="12.71" rx="0.92" />
      <path d="M 9.69 5.65 v 12.71" />
      <path d="M 5.99 8.54 h 2.31 M 5.99 11.42 h 2.31 M 5.99 14.31 h 2.31" />
    </Svg>
  )
}

export function IconSplit(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="4.53" width="14.94" height="14.94" rx="1" />
      <path d="M 4.53 12 h 14.94" />
    </Svg>
  )
}

export function IconPrintLayout(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="6.8" y="4.49" width="10.4" height="15.02" rx="0.92" />
      <path d="M 9.11 7.96 h 5.78 M 9.11 10.85 h 5.78 M 9.11 13.73 h 5.78 M 9.11 16.62 h 3.47" />
    </Svg>
  )
}

export function IconWebLayout(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="15.02" height="12.71" rx="0.92" />
      <path d="M 4.49 8.54 h 15.02" />
      <path d="M 6.8 11.42 h 10.4 M 6.8 13.73 h 10.4 M 6.8 16.04 h 6.93" />
    </Svg>
  )
}

export function IconGridlines(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="4.53" width="14.94" height="14.94" rx="1" />
      <path d="M 4.53 9.51 h 14.94 M 4.53 14.49 h 14.94 M 9.51 4.53 v 14.94 M 14.49 4.53 v 14.94" />
    </Svg>
  )
}

export function IconNewWindow(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="7.96" width="11.55" height="11.55" rx="0.92" />
      <path d="M 7.96 7.96 v -2.31 a 1.16 1.16 0 0 1 1.16 -1.15 h 9.24 a 1.16 1.16 0 0 1 1.16 1.16 v 9.24 a 1.16 1.16 0 0 1 -1.15 1.16 h -2.31" />
      <path d="M 10.27 13.73 h 4.62 M 12.58 11.42 v 4.62" />
    </Svg>
  )
}

export function IconArrangeAll(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.07" width="15.02" height="6.01" rx="0.92" />
      <rect x="4.49" y="12.92" width="15.02" height="6.01" rx="0.92" />
    </Svg>
  )
}

export function IconSwitchWindows(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="9.11" width="10.4" height="9.24" rx="0.92" />
      <path d="M 8.54 9.11 v -2.31 a 1.16 1.16 0 0 1 1.16 -1.15 h 8.66 a 1.16 1.16 0 0 1 1.16 1.16 v 8.09 a 1.16 1.16 0 0 1 -1.15 1.16 h -3.46" />
    </Svg>
  )
}

export function IconPosition(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="4.49" width="15.02" height="15.02" rx="1.16" />
      <rect x="8.54" y="8.54" width="6.93" height="6.93" />
    </Svg>
  )
}

export function IconWrapText(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="8.54" width="6.93" height="6.93" />
      <path d="M 13.73 5.07 h 5.78 M 13.73 8.54 h 5.78 M 13.73 12 h 5.78 M 13.73 15.47 h 5.78 M 4.49 18.93 h 15.02 M 4.49 5.07 h 6.93" />
    </Svg>
  )
}

export function IconDoc(props: IconProps) {
  return (
    <Svg {...props}>
      {PAGE}
      <path d="M15 2.25 V6 h3.75" />
      <path d="M8.25 9.75 h7.5 M8.25 13.5 h7.5 M8.25 17.25 h5.25" />
    </Svg>
  )
}

/* ---------- AI panel ---------- */

export function IconSend(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.52 12 19.48 5.03 15.87 18.97 11.48 14.06 z" />
      <path d="M 11.48 14.06 19.48 5.03" />
    </Svg>
  )
}

export function IconStop(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="5" y="5" width="14" height="14" rx="2.625" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconGear(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="12" cy="12" r="2.67" />
      <path d="M 12 4.47 v 2.43 M 12 17.1 v 2.43 M 19.53 12 h -2.43 M 6.9 12 h -2.43 M 17.35 6.65 l -1.7 1.7 M 8.36 15.65 l -1.7 1.7 M 17.35 17.35 15.65 15.65 M 8.36 8.36 6.65 6.65" />
    </Svg>
  )
}

/** collapse the right sidebar: panel outline + arrow pushing into it */
export function IconSidebarCollapse(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="15.02" height="12.71" rx="1.16" />
      <path d="M 14.89 5.65 v 12.71" />
      <path d="M 6.8 12 h 5.08 M 9.92 9.57 12.35 12 l -2.43 2.43" />
    </Svg>
  )
}

/** Mirror of IconSidebarCollapse for the LEFT-docked AI panel (right panes keep the original).
 *  Sheets-parity glyph (16-canvas, 1.2/1.3 stroke), self-contained so the shared
 *  Svg wrapper's 24-canvas pinned stroke doesn't alter its weight. */
export function IconSidebarCollapseLeft({ size = 24 }: IconProps) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.2"
      strokeLinecap="round"
      aria-hidden
    >
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
      <path d="M5.5 2.5v11" />
      <path d="M12.5 8H8.1M9.8 5.9 7.7 8l2.1 2.1" strokeWidth="1.3" strokeLinejoin="round" />
    </svg>
  )
}

export function IconClock(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="12" cy="12" r="7.47" />
      <path d="M 12 8.02 V 12 l 2.86 1.99" />
    </Svg>
  )
}

export function IconPaperclip(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 18.75 10.92 12.27 17.4 a 4.59 4.59 0 0 1 -6.48 -6.48 l 6.75 -6.75 a 3.11 3.11 0 0 1 4.32 4.32 l -6.75 6.75 a 1.49 1.49 0 0 1 -2.16 -2.16 l 6.21 -6.21" />
    </Svg>
  )
}

export function IconNewChat(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 19.01 10.98 v -3.82 A 2.17 2.17 0 0 0 16.85 4.99 H 7.15 a 2.17 2.17 0 0 0 -2.17 2.17 v 7.78 a 2.17 2.17 0 0 0 2.17 2.17 h 1.4 v 2.55 l 3.32 -2.55 h 1.66" />
      <path d="M 17.36 13.79 v 5.1 M 14.81 16.34 h 5.1" />
    </Svg>
  )
}

/* ---------- titlebar quick access ---------- */

/** Design-supplied glyphs on the 1:16 stroke:canvas ratio (24-canvas / 1.5 stroke):
 *  the stroke scales proportionally with size instead of the pinnedStroke policy. */
function SvgRatio({ size = 24, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.5}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

export function IconSave(props: IconProps) {
  return (
    <SvgRatio {...props}>
      <path d="M3 4.5C3 3.67158 3.67158 3 4.5 3H17.1407L21 6.60325V19.5C21 20.3285 20.3285 21 19.5 21H4.5C3.67158 21 3 20.3285 3 19.5V4.5Z" />
      <path d="M12.0042 3L12 6.6923C12 6.86225 11.7761 7 11.5 7H7.5C7.22385 7 7 6.86225 7 6.6923V3H12.0042Z" />
      <path d="M7 13H17" />
      <path d="M7 17H12.0042" />
    </SvgRatio>
  )
}

export function IconUndo(props: IconProps) {
  return (
    <SvgRatio {...props}>
      <path d="M5.91026 4L2.5 7.14791L5.91026 10.8205" />
      <path d="M3.96154 7.41028H15.1636C18.5169 7.41028 21.3646 10.1484 21.4953 13.5C21.6334 17.0416 18.707 20.0769 15.1636 20.0769H6.88384" />
    </SvgRatio>
  )
}

export function IconRedo(props: IconProps) {
  return (
    <SvgRatio {...props}>
      <path d="M18.0897 4L21.5 7.14791L18.0897 10.8205" />
      <path d="M20.0385 7.41028H8.83636C5.4831 7.41028 2.63537 10.1484 2.5047 13.5C2.36657 17.0416 5.29296 20.0769 8.83636 20.0769H17.1162" />
    </SvgRatio>
  )
}

export function IconCursor(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.97 4.52 18.03 12.8 l -5.12 1.21 L 10.49 19.43 5.97 4.52 Z" />
    </Svg>
  )
}

export function IconPen(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m 3.78 20.22 1.17 -4.41 L 14.94 5.83 a 2.06 2.06 0 0 1 2.94 0 l 0.29 0.29 a 2.06 2.06 0 0 1 0 2.94 L 8.18 19.05 3.78 20.22 Z" />
      <path d="M 13.47 7.3 16.7 10.53" />
    </Svg>
  )
}

export function IconHighlighterPen(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 8.85 13.89 15.53 7.21 a 1.64 1.64 0 0 1 2.39 0 l -1.13 -1.13 1.13 1.13 a 1.64 1.64 0 0 1 0 2.39 L 11.24 16.28 l -3.28 0.88 0.88 -3.28 Z" />
      <rect
        x="5.7"
        y="18.05"
        width="12.6"
        height="3.02"
        rx="1.51"
        fill="currentColor"
        stroke="none"
        opacity="0.5"
      />
    </Svg>
  )
}

export function IconEraser(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m12.495 5.29 6.765 6.765 a1.98 1.98 0 0 1 0 2.805 L14.64 19.48 H10.02 L4.74 14.2 a1.98 1.98 0 0 1 0 -2.805 l4.95 -4.95 a1.98 1.98 0 0 1 2.805 0 Z" />
      <path d="M7.875 8.92 15.63 16.675" />
      <path d="M10.02 19.48 h10.56" />
    </Svg>
  )
}

export function IconTextBox(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="15.02" height="12.71" rx="1.16" />
      <path d="M 8.54 9.11 h 6.93 M 12 9.11 v 6.35" />
    </Svg>
  )
}

/** New slide (MS-style): slide frame with a title band and a split content area */
export function IconNewSlide(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="6.4" width="14.94" height="11.21" rx="0.62" />
      <path d="M 4.53 10.13 h 14.94" />
      <path d="M 13.25 10.13 V 17.6" />
    </Svg>
  )
}

/** Section: divider + disclosure triangle, slide thumbnails grouped beneath */
export function IconSection(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.53 5.78 h 14.94" />
      <path d="M 4.78 8.89 l 3.24 2.37 -3.24 2.37 z" fill="currentColor" stroke="none" />
      <rect x="10.13" y="8.89" width="9.34" height="4.36" rx="0.62" />
      <rect x="10.13" y="15.11" width="9.34" height="4.36" rx="0.62" />
    </Svg>
  )
}

export function IconRect(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="6.8" width="15.02" height="10.4" />
    </Svg>
  )
}

export function IconRoundRect(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="6.8" width="15.02" height="10.4" rx="2.89" />
    </Svg>
  )
}

export function IconEllipse(props: IconProps) {
  return (
    <Svg {...props}>
      <ellipse cx="12" cy="12" rx="7.51" ry="5.2" />
    </Svg>
  )
}

/** Slide show: from beginning (screen + play triangle) */
export function IconPlayFromStart(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="15.02" height="10.4" rx="0.92" />
      <path d="M 10.27 8.3 v 5.08 l 4.39 -2.54 z" fill="currentColor" stroke="none" />
      <path d="M 12 16.04 v 2.31 M 9.11 18.35 h 5.78" />
    </Svg>
  )
}

/** Slide show: from current slide (half screen + play triangle) */
export function IconPlayCurrent(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.49 10.27 v -3.7 a 0.92 0.92 0 0 1 0.92 -0.92 h 13.17 a 0.92 0.92 0 0 1 0.92 0.92 v 9.7 a 0.92 0.92 0 0 1 -0.92 0.92 H 12" />
      <path d="M 4.49 13.39 h 4.62 M 4.49 16.27 h 4.62 M 4.49 19.16 h 4.62" />
      <path d="M 12.23 9.23 v 5.08 l 4.39 -2.54 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/** Presenter view: main screen + small presenter screen */
export function IconPresenterView(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="10.4" height="8.09" rx="0.92" />
      <rect x="12.58" y="11.42" width="6.93" height="5.78" rx="0.92" fill="var(--surface, #fff)" />
      <circle cx="16.04" cy="13.4" r="1.15" fill="currentColor" stroke="none" />
      <path d="M 13.84 16.35 a 2.2 1.45 0 0 1 4.4 0 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/** Custom show: screen + gear dots */
export function IconCustomShow(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="15.02" height="10.4" rx="0.92" />
      <path d="M 7.96 9.11 h 8.09 M 7.96 12 h 4.62" />
      <path d="M 12 16.04 v 2.31 M 9.11 18.35 h 5.78" />
    </Svg>
  )
}

/** Set up slide show: screen + wrench slash */
export function IconSetupShow(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="15.02" height="10.4" rx="0.92" />
      <path d="M 8.54 13.73 13.73 8.54 M 13.16 8.54 h 1.73 v 1.73" />
      <path d="M 12 16.04 v 2.31 M 9.11 18.35 h 5.78" />
    </Svg>
  )
}

/** Hide slide: slide + slash */
export function IconHideSlide(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="5.65" y="6.8" width="12.71" height="10.4" rx="0.92" />
      <path d="M 4.49 19.51 19.51 4.49" />
    </Svg>
  )
}

/** Rehearse timings: stopwatch */
export function IconRehearse(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="12" cy="13.25" r="6.23" />
      <path d="M 12 10.13 V 13.25 l 2.24 1.74" />
      <path d="M 10.13 4.53 h 3.74 M 12 4.53 v 2.24" />
    </Svg>
  )
}

/** Record slide show: recording dot */
export function IconRecord(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="12" cy="12" r="7.48" />
      <circle cx="12" cy="12" r="3.1" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/** Chart: bar chart */
export function IconChart(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.53 4.53 v 14.94 h 14.94" />
      <rect x="7.64" y="12" width="2.74" height="4.98" fill="currentColor" stroke="none" />
      <rect x="12" y="8.27" width="2.74" height="8.72" fill="currentColor" stroke="none" />
      <rect x="16.36" y="10.13" width="2.74" height="6.85" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/** SmartArt: connected nodes */
export function IconSmartArt(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="9.11" y="4.49" width="5.78" height="4.16" rx="0.92" />
      <rect x="4.49" y="14.89" width="5.78" height="4.16" rx="0.92" />
      <rect x="13.73" y="14.89" width="5.78" height="4.16" rx="0.92" />
      <path d="M 12 8.65 v 2.77 M 12 11.42 L 7.38 14.89 M 12 11.42 l 4.62 3.47" />
    </Svg>
  )
}

/** WordArt: outlined A */
export function IconWordArt(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.2 19.48 L 12 4.52 l 6.8 14.97 M 7.92 14.45 h 8.16" />
    </Svg>
  )
}

/** Icon gallery: smiley */
export function IconIconLib(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="12" cy="12" r="7.47" />
      <circle cx="9.26" cy="10.13" r="0.87" fill="currentColor" stroke="none" />
      <circle cx="14.74" cy="10.13" r="0.87" fill="currentColor" stroke="none" />
      <path d="M 8.64 13.99 a 4.23 4.23 0 0 0 6.72 0" />
    </Svg>
  )
}

/** Video: film play */
export function IconVideo(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="6.8" width="15.02" height="10.4" rx="1.62" />
      <path d="M 10.5 9.69 l 3.93 2.31 -3.93 2.31 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/** Audio: speaker */
export function IconAudio(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.15 9.76 h 2.99 L 12.62 5.78 v 12.45 L 8.14 14.24 H 5.15 z" />
      <path d="M 15.49 9.01 a 4.23 4.23 0 0 1 0 5.98 M 17.98 6.77 a 7.47 7.47 0 0 1 0 10.46" />
    </Svg>
  )
}

/** Screen recording: screen + record dot */
export function IconScreenRec(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="15.02" height="10.4" rx="1.39" />
      <path d="M 9.11 18.93 h 5.78" />
      <circle cx="12" cy="10.85" r="2.31" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/** 3D model: cube */
export function Icon3d(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 12 4.47 l 6.68 3.77 v 7.53 L 12 19.53 l -6.68 -3.77 V 8.23 z" />
      <path d="M 12 12 l 6.68 -3.77 M 12 12 L 5.32 8.23 M 12 12 v 7.53" />
    </Svg>
  )
}

/** Zoom link: jump arrow + page */
export function IconZoomJump(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="4.53" width="9.96" height="7.47" rx="1" />
      <rect x="9.51" y="12" width="9.96" height="7.47" rx="1" />
      <path d="M 14.49 8.27 l 3.74 0 M 18.22 8.27 l -1.74 -1.74 M 18.22 8.27 l -1.74 1.74" />
    </Svg>
  )
}

/** Date-time: calendar + hour hand */
export function IconDateTime(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.93" y="6.3" width="10.83" height="10.83" rx="1.37" />
      <path d="M 4.93 9.72 h 10.83 M 8.01 4.93 V 7.44 M 13.14 4.93 V 7.44" />
      <circle cx="16.56" cy="15.99" r="3.42" fill="var(--surface, #fff)" />
      <path d="M 16.56 14.28 v 1.71 l 1.25 1.03" />
    </Svg>
  )
}

/** Object align/distribute (distinct from paragraph text alignment IconAlign*): color blocks + alignment baseline */
export function IconObjAlignLeft(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.65 4.49 v 15.02" />
      <rect x="7.96" y="6.23" width="10.4" height="4.16" rx="0.69" />
      <rect x="7.96" y="13.62" width="6.35" height="4.16" rx="0.69" />
    </Svg>
  )
}

export function IconObjAlignCenterH(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 12 4.49 v 15.02" />
      <rect x="6.23" y="6.23" width="11.55" height="4.16" rx="0.69" />
      <rect x="8.77" y="13.62" width="6.47" height="4.16" rx="0.69" />
    </Svg>
  )
}

export function IconObjAlignRight(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 18.35 4.49 v 15.02" />
      <rect x="5.65" y="6.23" width="10.4" height="4.16" rx="0.69" />
      <rect x="9.69" y="13.62" width="6.35" height="4.16" rx="0.69" />
    </Svg>
  )
}

export function IconObjAlignTop(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.49 5.65 h 15.02" />
      <rect x="6.23" y="7.96" width="4.16" height="10.4" rx="0.69" />
      <rect x="13.62" y="7.96" width="4.16" height="6.35" rx="0.69" />
    </Svg>
  )
}

export function IconObjAlignMiddle(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="6.23" y="6.23" width="4.16" height="11.55" rx="0.69" />
      <rect x="13.62" y="8.77" width="4.16" height="6.47" rx="0.69" />
      <path d="M 4.49 12 H 6.23 M 10.38 12 h 3.23 M 17.77 12 h 1.73" />
    </Svg>
  )
}

export function IconObjAlignBottom(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.49 18.35 h 15.02" />
      <rect x="6.23" y="5.65" width="4.16" height="10.4" rx="0.69" />
      <rect x="13.62" y="9.69" width="4.16" height="6.35" rx="0.69" />
    </Svg>
  )
}

export function IconObjDistributeH(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.65 4.49 v 15.02 M 18.35 4.49 v 15.02" />
      <rect x="9.92" y="6.4" width="4.16" height="11.21" rx="0.69" />
    </Svg>
  )
}

export function IconObjDistributeV(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.49 5.65 h 15.02 M 4.49 18.35 h 15.02" />
      <rect x="6.4" y="9.92" width="11.21" height="4.16" rx="0.69" />
    </Svg>
  )
}

export function IconObjFlipH(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 12 4.49 v 15.02" strokeDasharray="2.4 2.1" />
      <path d="M 9.1 6.9 v 10.2 h -4.4 z M 14.9 6.9 v 10.2 h 4.4 z" />
    </Svg>
  )
}

export function IconObjFlipV(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.49 12 h 15.02" strokeDasharray="2.4 2.1" />
      <path d="M 6.9 9.1 h 10.2 v -4.4 z M 6.9 14.9 h 10.2 v 4.4 z" />
    </Svg>
  )
}

export function IconSwitchRowCol(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 6.27 9.88 A 6.47 6.47 0 0 1 17.1 7.64" />
      <path d="M 17.73 4.53 v 3.36 H 14.37" />
      <path d="M 17.73 14.12 A 6.47 6.47 0 0 1 6.9 16.36" />
      <path d="M 6.27 19.47 v -3.36 h 3.36" />
    </Svg>
  )
}

export function IconEditChartData(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="5.3" y="5.3" width="10.15" height="10.15" rx="0.86" />
      <path d="M 5.3 8.65 h 10.15 M 8.65 5.3 v 10.15" />
      <path d="M 17.62 13.19 l 1.84 1.84 -4.64 4.64 -2.48 0.65 0.65 -2.48 z" />
    </Svg>
  )
}

export function IconChangeChartType(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 6.79 8.8 A 5.81 5.81 0 0 1 16.74 6.9" />
      <path d="M 17.45 4.3 v 3.08 H 14.37" />
      <rect x="5.6" y="13.9" width="2.96" height="5.45" fill="currentColor" stroke="none" />
      <rect x="10.34" y="11.29" width="2.96" height="8.06" fill="currentColor" stroke="none" />
      <rect x="15.08" y="12.59" width="2.96" height="6.75" fill="currentColor" stroke="none" />
    </Svg>
  )
}

// ── Animation effect icons (Animations tab; color comes from .rb-anim-* via currentColor) ──

const ANIM_EFFECT_BODIES: Record<AnimEffectKind, ReactNode> = {
  // entrance
  appear: (
    <>
      <rect x="4" y="5" width="16" height="14" rx="1.5" strokeDasharray="3 2.2" />
      <rect x="8.2" y="9.2" width="7.6" height="5.6" fill="currentColor" stroke="none" />
    </>
  ),
  fade: (
    <>
      <rect
        x="4.5"
        y="6.5"
        width="4.6"
        height="11"
        fill="currentColor"
        stroke="none"
        opacity="0.2"
      />
      <rect
        x="9.7"
        y="6.5"
        width="4.6"
        height="11"
        fill="currentColor"
        stroke="none"
        opacity="0.5"
      />
      <rect x="14.9" y="6.5" width="4.6" height="11" fill="currentColor" stroke="none" />
    </>
  ),
  flyIn: (
    <>
      <rect x="5.5" y="3.5" width="13" height="8" rx="1" />
      <path d="M12 21 V14.5 M8.8 17.4 L12 14.2 l3.2 3.2" />
    </>
  ),
  wipe: (
    <>
      <rect x="3.5" y="6" width="17" height="12" rx="1" />
      <rect x="3.5" y="6" width="8.2" height="12" fill="currentColor" stroke="none" />
      <path d="M13.3 12 h4.8 M15.9 9.8 l2.2 2.2 -2.2 2.2" />
    </>
  ),
  wipeDown: (
    <>
      <rect x="3.5" y="4.5" width="17" height="15" rx="1" />
      <rect x="3.5" y="4.5" width="17" height="6.4" fill="currentColor" stroke="none" />
      <path d="M12 12.6 V17.3 M9.9 15.3 L12 17.4 l2.1 -2.1" />
    </>
  ),
  splitIn: (
    <>
      <rect x="3.5" y="6" width="17" height="12" rx="1" />
      <path d="M5.8 12 h4.4 M8.3 9.9 l2.1 2.1 -2.1 2.1 M18.2 12 h-4.4 M15.7 9.9 l-2.1 2.1 2.1 2.1" />
    </>
  ),
  bounce: (
    <>
      <path d="M3.5 19.5 C6 7.5, 8.5 7.5, 11 19.5 C12.8 13, 14.6 13, 16.4 19.5" />
      <circle cx="19.3" cy="17.8" r="1.9" fill="currentColor" stroke="none" />
    </>
  ),
  flipIn: (
    <>
      <rect x="4" y="6" width="7" height="12" />
      <path d="M13.5 4.2 L20 6.4 V17.6 L13.5 19.8 Z" />
    </>
  ),
  zoom: (
    <>
      <rect x="9.2" y="9.2" width="5.6" height="5.6" fill="currentColor" stroke="none" />
      <path d="M8 8 L4.5 4.5 M4.5 8 V4.5 H8" />
      <path d="M16 16 L19.5 19.5 M19.5 16 V19.5 H16" />
    </>
  ),
  // emphasis
  pulse: (
    <>
      <circle cx="12" cy="12" r="2.2" fill="currentColor" stroke="none" />
      <circle cx="12" cy="12" r="5.6" opacity="0.65" />
      <circle cx="12" cy="12" r="9" opacity="0.35" />
    </>
  ),
  spin: (
    <>
      <path d="M12 5.4 A 6.6 6.6 0 1 1 5.4 12" />
      <path d="M3.3 13.9 L5.4 11.5 7.7 13.8" />
    </>
  ),
  grow: (
    <>
      <rect x="7.5" y="7.5" width="9" height="9" />
      <path d="M4.5 19.5 L19.5 4.5 M4.5 15.2 V19.5 H8.8 M19.5 8.8 V4.5 H15.2" />
    </>
  ),
  teeter: (
    <>
      <rect x="5" y="9.5" width="14" height="9" rx="1" transform="rotate(-8 12 14)" />
      <path d="M6 6.6 A 7.5 4.5 0 0 1 18 6.6" />
      <path d="M5.2 4.2 L6 6.8 8.6 6.1 M18.8 4.2 L18 6.8 15.4 6.1" />
    </>
  ),
  // exit
  disappear: <rect x="4" y="5" width="16" height="14" rx="1.5" strokeDasharray="3 2.2" />,
  fadeOut: (
    <>
      <rect x="4.5" y="6.5" width="4.6" height="11" fill="currentColor" stroke="none" />
      <rect
        x="9.7"
        y="6.5"
        width="4.6"
        height="11"
        fill="currentColor"
        stroke="none"
        opacity="0.5"
      />
      <rect
        x="14.9"
        y="6.5"
        width="4.6"
        height="11"
        fill="currentColor"
        stroke="none"
        opacity="0.2"
      />
    </>
  ),
  flyOut: (
    <>
      <rect x="5.5" y="3.5" width="13" height="8" rx="1" />
      <path d="M12 14.2 V20.7 M8.8 17.5 L12 20.7 l3.2 -3.2" />
    </>
  ),
  wipeOut: (
    <>
      <rect x="3.5" y="6" width="17" height="12" rx="1" />
      <rect x="12.3" y="6" width="8.2" height="12" fill="currentColor" stroke="none" />
      <path d="M10.7 12 H5.9 M8.1 9.8 L5.9 12 l2.2 2.2" />
    </>
  ),
  shrink: (
    <>
      <rect
        x="9.6"
        y="9.6"
        width="4.8"
        height="4.8"
        fill="currentColor"
        stroke="none"
        transform="rotate(15 12 12)"
      />
      <path d="M4.5 4.5 L8.3 8.3 M8.3 5.1 V8.3 H5.1" />
      <path d="M19.5 19.5 L15.7 15.7 M15.7 18.9 V15.7 H18.9" />
    </>
  ),
  zoomOut: (
    <>
      <rect x="3.5" y="3.5" width="17" height="17" rx="1" strokeDasharray="2.6 2" />
      <path d="M5.6 5.6 L9.4 9.4 M9.4 6.2 V9.4 H6.2" />
      <path d="M18.4 18.4 L14.6 14.6 M14.6 17.8 V14.6 H17.8" />
    </>
  ),
  motionPath: (
    <>
      <path d="M4.5 6 C10 3, 14 9, 18.6 15.8" strokeDasharray="3 2.2" />
      <path d="M18.9 10.9 l0.5 5.4 -5.4 -0.6" />
    </>
  ),
}

export function AnimEffectIcon({ kind, size }: { kind: AnimEffectKind; size?: number }) {
  return <Svg size={size}>{ANIM_EFFECT_BODIES[kind]}</Svg>
}

/* ---------- transition gallery (drawn to the shared icon standard, replacing
   the old unicode text glyphs whose size and weight came from the font) ---- */

export function IconTransNone(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="6" width="14.5" height="12" rx="1.5" />
      <path d="M6.5 16.5 17.5 7.5" />
    </Svg>
  )
}

export function IconTransMorph(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M4.75 9h11.5M13.5 6.25 16.25 9l-2.75 2.75" />
      <path d="M19.25 15H7.75M10.5 17.75 7.75 15l2.75-2.75" />
    </Svg>
  )
}

export function IconTransFade(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="4.75" width="11" height="9.5" rx="1.4" strokeDasharray="2.4 2.2" />
      <rect x="8.25" y="9.75" width="11" height="9.5" rx="1.4" />
    </Svg>
  )
}

export function IconTransPush(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="6" width="14.5" height="12" rx="1.5" />
      <path d="M12 15.5v-6M9.25 12.25 12 9.5l2.75 2.75" />
    </Svg>
  )
}

export function IconTransWipe(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="6" width="14.5" height="12" rx="1.5" />
      <path d="M7.5 12h6M11.25 9.75 13.5 12l-2.25 2.25" />
    </Svg>
  )
}

export function IconTransSplit(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="6" width="14.5" height="12" rx="1.5" />
      <path d="M12 8.5v7" />
      <path d="M9.5 12H7M8.25 10.75 7 12l1.25 1.25M14.5 12H17M15.75 10.75 17 12l-1.25 1.25" />
    </Svg>
  )
}

export function IconTransCircle(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="12" cy="12" r="7.25" />
      <circle cx="12" cy="12" r="3.25" />
    </Svg>
  )
}

export function IconTransCover(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M14.5 8.25V6.25a1.5 1.5 0 0 0-1.5-1.5H6.25a1.5 1.5 0 0 0-1.5 1.5V13a1.5 1.5 0 0 0 1.5 1.5h2" />
      <rect x="9.75" y="9.75" width="9.5" height="9.5" rx="1.5" />
    </Svg>
  )
}

export function IconTransPull(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="4.75" width="9.5" height="9.5" rx="1.5" />
      <path d="M9.5 15.75v2a1.5 1.5 0 0 0 1.5 1.5h6.75a1.5 1.5 0 0 0 1.5-1.5V11a1.5 1.5 0 0 0-1.5-1.5h-2" />
    </Svg>
  )
}

export function IconTransDissolve(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="6" width="14.5" height="12" rx="1.5" />
      <circle cx="9" cy="10" r="0.8" fill="currentColor" stroke="none" />
      <circle cx="13.75" cy="9.25" r="0.8" fill="currentColor" stroke="none" />
      <circle cx="15.75" cy="13.25" r="0.8" fill="currentColor" stroke="none" />
      <circle cx="8.25" cy="14" r="0.8" fill="currentColor" stroke="none" />
      <circle cx="12" cy="14.75" r="0.8" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconTransZoom(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M14 4.75h5.25V10M19 5 13.75 10.25" />
      <path d="M10 19.25H4.75V14M5 19 10.25 13.75" />
    </Svg>
  )
}

export function IconTransRandom(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="5" y="5" width="14" height="14" rx="2.5" />
      <circle cx="9.25" cy="9.25" r="1" fill="currentColor" stroke="none" />
      <circle cx="14.75" cy="9.25" r="1" fill="currentColor" stroke="none" />
      <circle cx="12" cy="12" r="1" fill="currentColor" stroke="none" />
      <circle cx="9.25" cy="14.75" r="1" fill="currentColor" stroke="none" />
      <circle cx="14.75" cy="14.75" r="1" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/* ---------- animation gallery chrome (star / none / motion paths) -------- */

export function IconAnimStar(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m12 4.5 2.15 4.9 5.35.5-4 3.6 1.15 5.25L12 16l-4.65 2.75 1.15-5.25-4-3.6 5.35-.5Z" />
    </Svg>
  )
}

export function IconAnimNone(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="12" cy="12" r="7.25" />
      <path d="M6.9 17.1 17.1 6.9" />
    </Svg>
  )
}

export function IconPathRight(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M4.75 12H18M14.75 8.75 18 12l-3.25 3.25" />
    </Svg>
  )
}

export function IconPathDown(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M12 4.75V18M8.75 14.75 12 18l3.25-3.25" />
    </Svg>
  )
}

export function IconPathDiagonal(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M5.5 5.5 18 18M18 13.4V18h-4.6" />
    </Svg>
  )
}

export function IconPathCircle(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="12" cy="12" r="6.75" />
      <path d="M15.5 3.9 12.7 5.2l1.3 2.75" />
    </Svg>
  )
}

export function IconPathZigzag(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M4.75 15.5 8.75 9.5l3.75 5.5 4.25-6.25" />
      <path d="M17.5 12.9V8.25h-4.6" />
    </Svg>
  )
}

export function IconCrop(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 4.75v11.25h11.25" />
      <path d="M4.75 8H16v11.25" />
    </Svg>
  )
}

export function IconNoneX(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m6.5 6.5 11 11M17.5 6.5l-11 11" />
    </Svg>
  )
}

/** AI feature glyphs shared by the ribbon Home tab and the canvas AI bar.
 * Fixed 1.5-unit stroke (not pinnedStroke) to keep the ribbon rendering,
 * where CSS sizes them; `size` is for other hosts. */
function AiFeatureSvg({ size, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

export function IconAiBeautify(props: IconProps) {
  return (
    <AiFeatureSvg {...props}>
      <path d="M10.4948 15.0904C9.91037 14.5539 9.12711 14.5196 9.12711 14.5196C9.12711 14.5196 8.37299 14.4408 7.8691 14.7887C7.03613 15.3629 7.02071 16.5283 7.01385 16.7014C6.995 17.2293 7.03785 17.8652 6.71563 18.4239C6.24088 19.2466 5 19.5894 5 19.5894C5 19.5894 5.21081 19.8448 5.51246 19.9939C5.51246 19.9939 6.47226 20.5423 8.45012 20.0727C8.99 19.9442 9.58816 19.7796 10.0732 19.4763C10.6662 19.1061 11.105 18.5799 11.2232 18.3194C11.7374 17.2053 11.4752 15.9902 10.4948 15.0904Z" />
      <path d="M18.9872 4.57628C18.9772 4.39067 18.7726 4.29501 18.7726 4.29501C18.7726 4.29501 18.5538 4.25444 18.3736 4.34695C18.0654 4.50498 17.6181 4.70138 15.4274 7.19175C13.9033 8.92525 12.0146 11.2553 10.9207 12.5777C10.1021 13.5619 10.0281 14.2654 10.1806 14.6721C10.2108 14.7514 10.5825 15.0484 10.9207 15.2661C11.1663 15.4249 11.3406 15.5257 11.6199 15.5437C12.0717 15.5739 12.5127 15.2661 12.5127 15.2661C12.5127 15.2661 12.7483 15.0623 13.1086 14.587C13.8466 13.6144 15.1527 11.7581 16.2918 9.96583C17.7356 7.70256 18.8308 5.34067 18.8308 5.34067C18.8308 5.34067 19.0033 4.88149 18.9872 4.57628Z" />
      <path d="M7 4L7.22106 4.59745C7.51094 5.38087 7.65589 5.77259 7.94166 6.05833C8.22743 6.34409 8.61914 6.48903 9.40257 6.77893L10 7L9.40257 7.22107C8.61914 7.51097 8.22743 7.65592 7.94166 7.94167C7.65589 8.22741 7.51094 8.61913 7.22106 9.40255L7 10L6.77894 9.40255C6.48906 8.61913 6.34411 8.22741 6.05834 7.94167C5.77257 7.65592 5.38086 7.51097 4.59743 7.22107L4 7L4.59743 6.77893C5.38086 6.48903 5.77257 6.34409 6.05834 6.05833C6.34411 5.77259 6.48906 5.38087 6.77894 4.59745L7 4Z" />
      <path d="M18 16L18.1474 16.3983C18.3406 16.9206 18.4373 17.1817 18.6278 17.3722C18.8183 17.5627 19.0794 17.6594 19.6017 17.8526L20 18L19.6017 18.1474C19.0794 18.3406 18.8183 18.4373 18.6278 18.6278C18.4373 18.8183 18.3406 19.0794 18.1474 19.6017L18 20L17.8526 19.6017C17.6594 19.0794 17.5627 18.8183 17.3722 18.6278C17.1817 18.4373 16.9206 18.3406 16.3983 18.1474L16 18L16.3983 17.8526C16.9206 17.6594 17.1817 17.5627 17.3722 17.3722C17.5627 17.1817 17.6594 16.9206 17.8526 16.3983L18 16Z" />
    </AiFeatureSvg>
  )
}

export function IconAiFactCheck(props: IconProps) {
  return (
    <AiFeatureSvg {...props}>
      <circle cx="11.9997" cy="12.0004" r="9.53571" />
      <path
        d="M18.2407 8.20112C18.5754 7.86675 19.118 7.86668 19.4526 8.20112C19.7873 8.53587 19.7873 9.07931 19.4526 9.41402L14.0727 14.7929C13.9119 14.9535 13.6936 15.044 13.4663 15.0439C13.2392 15.0437 13.0213 14.9535 12.8608 14.7929L11.3921 13.3222C11.1247 13.0544 11.0714 12.6529 11.2319 12.332L8.70555 14.8593C8.37087 15.194 7.82741 15.1939 7.49266 14.8593L4.53563 11.9023C4.20123 11.5676 4.20121 11.025 4.53563 10.6904C4.87039 10.3559 5.41388 10.3557 5.74852 10.6904L8.0991 13.041L12.8725 8.26753C13.2072 7.93333 13.7498 7.93331 14.0845 8.26753C14.4189 8.60218 14.4188 9.14572 14.0845 9.48042L11.6147 11.9492C11.936 11.7883 12.3371 11.8423 12.605 12.1103L13.4673 12.9736L18.2407 8.20112Z"
        fill="currentColor"
        stroke="none"
      />
    </AiFeatureSvg>
  )
}

export function IconAiImage(props: IconProps) {
  return (
    <AiFeatureSvg {...props}>
      <path d="M7.55042 9.62398C8.33741 9.62398 8.97538 8.986 8.97538 8.19902C8.97538 7.41203 8.33741 6.77405 7.55042 6.77405C6.76343 6.77405 6.12545 7.41203 6.12545 8.19902C6.12545 8.986 6.76343 9.62398 7.55042 9.62398Z" />
      <path d="M20.8474 11.05C20.8496 11.4966 20.8496 11.9708 20.8496 12.475C20.8496 16.7293 20.8496 18.8565 19.528 20.1782C18.2063 21.4998 16.0791 21.4998 11.8248 21.4998C7.57047 21.4998 5.44331 21.4998 4.12165 20.1782C2.8 18.8565 2.8 16.7293 2.8 12.475C2.8 8.22066 2.8 6.0935 4.12165 4.77184C5.44331 3.4502 7.57047 3.4502 11.8248 3.4502C12.3289 3.4502 12.8032 3.4502 13.2498 3.45239" />
      <path d="M18.4756 2.49915L18.7206 3.16131C19.0419 4.02958 19.2026 4.46372 19.5193 4.78041C19.836 5.09712 20.2701 5.25776 21.1384 5.57905L21.8006 5.82407L21.1384 6.06909C20.2701 6.39038 19.836 6.55103 19.5193 6.86772C19.2026 7.18442 19.0419 7.61856 18.7206 8.48683L18.4756 9.14899L18.2306 8.48683C17.9093 7.61856 17.7487 7.18442 17.432 6.86772C17.1153 6.55103 16.6811 6.39038 15.8128 6.06909L15.1507 5.82407L15.8128 5.57905C16.6811 5.25776 17.1153 5.09712 17.432 4.78041C17.7487 4.46372 17.9093 4.02958 18.2306 3.16131L18.4756 2.49915Z" />
      <path d="M5.17599 21.0236C9.32974 16.06 13.9862 9.51375 20.8483 13.9391" />
    </AiFeatureSvg>
  )
}

/** Genspark brand mark (rounded-square sparkle badge), inline so it renders
 * crisply at device resolution instead of going through <img> rasterization */
export function GensparkMark({ size = 30 }: { size?: number }) {
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
        fill="#000"
      />
    </svg>
  )
}
