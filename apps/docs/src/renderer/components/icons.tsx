/** Small monochrome SVG icons approximating Word's ribbon glyphs. */

import type { ReactNode } from 'react'

interface IconProps {
  size?: number
}

/** Constant painted stroke instead of proportional scaling: ~1.5px lines on
 *  20px+ glyphs, ~1.25px on the 13-19px ones, ~1.1px below (a proportional
 *  1-unit stroke would paint 1.75px at 28px and hairlines at small sizes).
 *  stroke-width is in 16-canvas units: units = painted-px × 16 / rendered-px. */
function pinnedStroke(size: number): number {
  const painted = size >= 20 ? 1.5 : size >= 13 ? 1.25 : 1.1
  return (painted * 16) / size
}

function Svg({ size = 20, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
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

export function IconBullets(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="3.66" cy="4.45" r="0.87" fill="currentColor" stroke="none" />
      <circle cx="3.66" cy="8" r="0.87" fill="currentColor" stroke="none" />
      <circle cx="3.66" cy="11.56" r="0.87" fill="currentColor" stroke="none" />
      <path d="M 6.42 4.45 h 6.32 M 6.42 8 h 6.32 M 6.42 11.56 h 6.32" />
    </Svg>
  )
}

export function IconNumbered(props: IconProps) {
  return (
    <Svg {...props}>
      <text
        x="1"
        y="5.4"
        fontSize="5.4"
        fill="currentColor"
        stroke="none"
        fontFamily="Segoe UI, sans-serif"
      >
        1
      </text>
      <text
        x="1"
        y="10.4"
        fontSize="5.4"
        fill="currentColor"
        stroke="none"
        fontFamily="Segoe UI, sans-serif"
      >
        2
      </text>
      <text
        x="1"
        y="15.4"
        fontSize="5.4"
        fill="currentColor"
        stroke="none"
        fontFamily="Segoe UI, sans-serif"
      >
        3
      </text>
      <path d="M6.5 3.5h8M6.5 8.5h8M6.5 13.5h8" />
    </Svg>
  )
}

export function IconMultilevel(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.61" y="3.84" width="1.39" height="1.39" fill="currentColor" stroke="none" />
      <path d="M 5.69 4.54 h 6.93" />
      <rect x="4.54" y="7.31" width="1.39" height="1.39" fill="currentColor" stroke="none" />
      <path d="M 7.62 8 h 5.01" />
      <rect x="6.46" y="10.77" width="1.39" height="1.39" fill="currentColor" stroke="none" />
      <path d="M 9.54 11.47 h 3.08" />
    </Svg>
  )
}

export function IconIndentDec(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.02 3.44 h 9.96 M 8 6.17 h 4.98 M 8 8.41 h 4.98 M 8 10.66 h 4.98 M 3.02 12.98 h 9.96" />
      <path d="M 5.68 6.17 3.19 8.41 l 2.49 2.24 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconIndentInc(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.02 3.44 h 9.96 M 8 6.17 h 4.98 M 8 8.41 h 4.98 M 8 10.66 h 4.98 M 3.02 12.98 h 9.96" />
      <path d="M 3.19 6.17 l 2.49 2.24 -2.49 2.24 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconAlignLeft(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.02 3.85 h 9.96 M 3.02 6.34 h 6.64 M 3.02 8.83 h 9.96 M 3.02 11.32 h 6.64" />
    </Svg>
  )
}

export function IconAlignCenter(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.02 3.85 h 9.96 M 4.68 6.34 h 6.64 M 3.02 8.83 h 9.96 M 4.68 11.32 h 6.64" />
    </Svg>
  )
}

export function IconAlignRight(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.02 3.85 h 9.96 M 6.34 6.34 h 6.64 M 3.02 8.83 h 9.96 M 6.34 11.32 h 6.64" />
    </Svg>
  )
}

export function IconAlignJustify(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.02 3.85 h 9.96 M 3.02 6.34 h 9.96 M 3.02 8.83 h 9.96 M 3.02 11.32 h 9.96" />
    </Svg>
  )
}

export function IconDirLtr(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.02 3.85 h 9.96 M 3.02 6.34 h 6.64 M 3.02 11.32 h 7.1" />
      <path d="M 9.8 9.4 l 2.9 1.92 -2.9 1.92 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconDirRtl(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.02 3.85 h 9.96 M 6.34 6.34 h 6.64 M 5.88 11.32 h 7.1" />
      <path d="M 6.2 9.4 l -2.9 1.92 2.9 1.92 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconLineSpacing(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 8 3.9 h 4.92 M 8 6.69 h 4.92 M 8 9.48 h 4.92 M 8 12.26 h 4.92" />
      <path d="M 4.31 4.06 v 7.87 M 2.92 5.7 l 1.39 -1.64 1.39 1.64 M 2.92 10.3 l 1.39 1.64 1.39 -1.64" />
    </Svg>
  )
}

export function IconClearFormat(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.17 12.67 5.67 3.5l3.5 9.17M3.4 9.5h4.55" />
      <path d="m10.4 9.1 3.4 3.4M13.8 9.1l-3.4 3.4" />
    </Svg>
  )
}

export function IconGrowFont(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.17 12.67 5.67 3.5l3.5 9.17M3.4 9.5h4.55" />
      <path d="M12 11.67V4.5M9.93 6.57 12 4.5l2.07 2.07" />
    </Svg>
  )
}

export function IconShrinkFont(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.17 12.67 5.67 3.5l3.5 9.17M3.4 9.5h4.55" />
      <path d="M12 4.5v7.17M9.93 9.6 12 11.67l2.07-2.07" />
    </Svg>
  )
}

export function IconHighlight(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3 10.5 9.5 4a1.4 1.4 0 0 1 2 0l0.5 0.5a1.4 1.4 0 0 1 0 2L5.5 13H3z" fill="none" />
      <path d="M2.2 13h4" strokeWidth="1" />
    </Svg>
  )
}

/* ---------- shared shapes ---------- */

/** page outline used by many icons */
const PAGE = <path d="M4.92 3h4.62l1.93 1.93v8.09h-6.55z" />

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
      <path d="M 5.43 12.11 H 3.93 C 3.23 12.11 2.67 11.54 2.67 10.84 V 4.53 C 2.67 3.83 3.23 3.26 3.93 3.26 H 4.88 M 11.51 5.16 V 4.53 C 11.51 3.83 10.94 3.26 10.25 3.26 H 9.3" />
      <rect x="5.19" y="2" width="3.79" height="1.89" rx="0.63" />
      <path d="M 12.14 5.16 H 6.46 C 5.76 5.16 5.19 5.72 5.19 6.42 V 12.74 C 5.19 13.43 5.76 14 6.46 14 H 10.32 L 13.4 10.68 V 6.42 C 13.4 5.72 12.84 5.16 12.14 5.16 Z" />
      <path d="M 7.09 7.37 H 11.51 M 7.09 9.58 H 9.61" />
      <path d="M 10.25 14 V 11.16 C 10.25 10.81 10.53 10.53 10.88 10.53 H 13.4" />
    </Svg>
  )
}

export function IconCut(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.27 12.85 L 5.94 11.71 L 11.32 2.4 M 4.68 2.33 L 10.05 11.65 L 10.73 12.85" />
      <circle cx="3.89" cy="12.08" r="1.58" />
      <circle cx="12.11" cy="12.08" r="1.58" />
    </Svg>
  )
}

export function IconCopy(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.67" y="4.67" width="9.33" height="9.33" rx="2" />
      <path d="M 9.67 2.67 H 4.67 C 3.56 2.67 2.67 3.56 2.67 4.67 V 9.67" />
    </Svg>
  )
}

export function IconFormatPainter(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 11 2.67 H 2.65 V 7.14 H 11 Z" />
      <path d="M 11.01 4.75 H 13.68 V 9.56 L 6.53 10.48 V 14" />
    </Svg>
  )
}

/* ---------- Insert ---------- */

export function IconTable(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.44" width="9.96" height="9.13" rx="0.66" />
      <path d="M 3.02 6.51 h 9.96 M 3.02 9.58 h 9.96 M 6.34 3.44 v 9.13 M 9.66 3.44 v 9.13" />
    </Svg>
  )
}

export function IconPicture(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.85" width="9.96" height="8.3" rx="0.66" />
      <circle cx="5.84" cy="6.51" r="0.91" />
      <path d="M 3.44 11.32 6.76 8 l 2.49 2.49 1.66 -1.66 1.66 1.66" />
    </Svg>
  )
}

export function IconRemoveBg(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.85" width="9.96" height="8.3" rx="0.66" strokeDasharray="2.2 1.6" />
      <circle cx="8" cy="6.92" r="1.41" />
      <path d="M 5.43 12.15 c 0.33 -1.91 1.41 -2.9 2.57 -2.9 s 2.24 1 2.57 2.91" />
    </Svg>
  )
}

export function IconCrop(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.41 2.97 v 7.62 h 7.62" />
      <path d="M 2.97 5.41 h 7.62 v 7.62" />
    </Svg>
  )
}

export function IconReplacePicture(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.87" y="6.03" width="7.11" height="6.32" rx="0.63" />
      <circle cx="4.92" cy="8" r="0.71" />
      <path d="M 3.26 11.79 l 2.13 -2.13 1.5 1.5 1.11 -1.11 1.42 1.42" />
      <path d="M 9.19 3.73 h 3.63 m 0 0 -1.34 -1.26 m 1.34 1.26 -1.34 1.26" />
    </Svg>
  )
}

export function IconChart(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 2.99 2.99 v 10.01 h 10.01" />
      <rect x="4.81" y="8" width="1.82" height="5.01" />
      <rect x="7.55" y="5.27" width="1.82" height="7.74" />
      <rect x="10.27" y="6.64" width="1.82" height="6.37" />
    </Svg>
  )
}

export function IconShapes(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="6.28" cy="6.28" r="3.1" />
      <rect x="7.57" y="7.57" width="5.59" height="5.59" rx="0.69" fill="var(--surface, #fff)" />
    </Svg>
  )
}

export function IconLink(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 6.91 9.09 9.09 6.91" />
      <path d="M 7.55 5.09 8.91 3.72 a 2.37 2.37 0 0 1 3.37 3.37 L 10.91 8.46" />
      <path d="M 8.46 10.91 7.09 12.28 a 2.37 2.37 0 0 1 -3.37 -3.37 l 1.37 -1.36" />
    </Svg>
  )
}

export function IconComment(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 2.99 3.91 h 10.01 v 6.83 h -5.46 L 4.81 13.46 v -2.73 h -1.82 z" />
      <path d="M 5.27 6.18 h 5.46 M 5.27 8.46 h 3.64" />
    </Svg>
  )
}

export function IconPageBreak(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.92 3 h 6.16 v 3.47 M 4.92 3 v 3.47 M 4.92 13.01 h 6.16 v -3.46 M 4.92 13.01 v -3.46" />
      <path
        d="M 3 8 h 1.54 M 5.69 8 h 1.54 M 8.39 8 h 1.54 M 11.08 8 h 1.93"
        strokeDasharray="none"
      />
    </Svg>
  )
}

export function IconHeader(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.15" y="3" width="7.7" height="10.01" rx="0.62" />
      <path d="M 5.31 4.92 h 5.39 M 5.31 6.31 h 5.39" strokeWidth="1" opacity="0.9" />
    </Svg>
  )
}

export function IconFooter(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.15" y="3" width="7.7" height="10.01" rx="0.62" />
      <path d="M 5.31 9.69 h 5.39 M 5.31 11.08 h 5.39" strokeWidth="1" opacity="0.9" />
    </Svg>
  )
}

export function IconPageNumber(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.15" y="3" width="7.7" height="10.01" rx="0.62" />
      <TextGlyph x={6.15} y={10.31} s={5.39}>
        #
      </TextGlyph>
    </Svg>
  )
}

export function IconSymbol(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={3.2} y={13} s={13}>
        Ω
      </TextGlyph>
    </Svg>
  )
}

export function IconEquation(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={4} y={12.5} s={12}>
        π
      </TextGlyph>
    </Svg>
  )
}

/* ---------- Table Design / Layout ---------- */

export function IconTableDelete(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.26" y="3.62" width="8.03" height="7.3" rx="0.58" />
      <path
        d="M 3.26 6.03 h 8.03 M 3.26 8.51 h 8.03 M 5.96 3.62 v 7.3 M 8.58 3.62 v 7.3"
        strokeWidth="1"
      />
      <path d="M 9.17 9.17 h 4.09 v 4.09 H 9.17 z" fill="var(--surface, #fff)" stroke="none" />
      <path d="m 9.97 9.97 2.63 2.63 M 12.6 9.97 l -2.63 2.63" strokeWidth="1" />
    </Svg>
  )
}

export function IconRowInsertAbove(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 8 6.18 V 2.98 M 6.63 4.35 8 2.98 l 1.37 1.37" />
      <rect x="3.44" y="7.62" width="9.12" height="5.32" rx="0.61" />
      <path d="M 3.44 10.28 h 9.12 M 8 7.62 v 5.32" strokeWidth="1" />
    </Svg>
  )
}

export function IconRowInsertBelow(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.44" y="3.06" width="9.12" height="5.32" rx="0.61" />
      <path d="M 3.44 5.72 h 9.12 M 8 3.06 v 5.32" strokeWidth="1" />
      <path d="M 8 9.82 v 3.19 M 6.63 11.65 8 13.02 l 1.37 -1.37" />
    </Svg>
  )
}

export function IconColInsertLeft(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 6.18 8 H 2.98 M 4.35 6.63 2.98 8 l 1.37 1.37" />
      <rect x="7.62" y="3.44" width="5.32" height="9.12" rx="0.61" />
      <path d="M 10.28 3.44 v 9.12 M 7.62 8 h 5.32" strokeWidth="1" />
    </Svg>
  )
}

export function IconColInsertRight(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.06" y="3.44" width="5.32" height="9.12" rx="0.61" />
      <path d="M 5.72 3.44 v 9.12 M 3.06 8 h 5.32" strokeWidth="1" />
      <path d="M 9.82 8 h 3.19 M 11.65 6.63 13.02 8 l -1.37 1.37" />
    </Svg>
  )
}

export function IconMergeCells(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="4.15" width="10.01" height="7.7" rx="0.62" />
      <path d="M 8 4.15 v 1.54 M 8 10.31 v 1.54" strokeWidth="1" />
      <path d="M 4.46 8 h 2.31 M 5.77 7 6.77 8 5.77 9" />
      <path d="M 11.54 8 h -2.31 M 10.23 7 9.23 8 l 1 1" />
    </Svg>
  )
}

export function IconSplitCells(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="4.15" width="10.01" height="7.7" rx="0.62" />
      <path d="M 8 4.15 v 7.7" strokeWidth="1" />
      <path d="M 6.92 8 h -2.31 M 5.61 7 4.61 8 l 1 1" />
      <path d="M 9.08 8 h 2.31 M 10.39 7 11.39 8 l -1 1" />
    </Svg>
  )
}

export function IconRowDelete(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.44" width="9.96" height="9.13" rx="0.66" />
      <path d="M 3.02 6.51 h 9.96 M 3.02 9.49 h 9.96" strokeWidth="1" />
      <path d="m 6.01 6.92 3.98 2.16 M 9.99 6.92 6.01 9.08" />
    </Svg>
  )
}

export function IconColDelete(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.44" y="3.02" width="9.13" height="9.96" rx="0.66" />
      <path d="M 6.51 3.02 v 9.96 M 9.49 3.02 v 9.96" strokeWidth="1" />
      <path d="m 6.92 6.01 2.16 3.98 M 9.08 6.01 6.92 9.99" />
    </Svg>
  )
}

export function IconCellAlignTop(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.44" width="9.96" height="9.13" rx="0.66" />
      <path d="M 5.1 5.68 h 5.81 M 5.1 7.5 h 3.74" />
    </Svg>
  )
}

export function IconCellAlignMiddle(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.44" width="9.96" height="9.13" rx="0.66" />
      <path d="M 5.1 7.09 h 5.81 M 5.1 8.91 h 3.74" />
    </Svg>
  )
}

export function IconCellAlignBottom(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.44" width="9.96" height="9.13" rx="0.66" />
      <path d="M 5.1 8.5 h 5.81 M 5.1 10.32 h 3.74" />
    </Svg>
  )
}

export function IconBorderAll(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.02" width="9.96" height="9.96" rx="0.42" />
      <path d="M 3.02 8 h 9.96 M 8 3.02 v 9.96" />
    </Svg>
  )
}

export function IconBorderOuter(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.02" width="9.96" height="9.96" rx="0.42" />
      <path
        d="M 3.02 8 h 9.96 M 8 3.02 v 9.96"
        strokeWidth="1"
        strokeDasharray="1.5 1.7"
        opacity="0.55"
      />
    </Svg>
  )
}

export function IconBorderInner(props: IconProps) {
  return (
    <Svg {...props}>
      <rect
        x="3.02"
        y="3.02"
        width="9.96"
        height="9.96"
        rx="0.42"
        strokeWidth="1"
        strokeDasharray="1.5 1.7"
        opacity="0.55"
      />
      <path d="M 3.02 8 h 9.96 M 8 3.02 v 9.96" />
    </Svg>
  )
}

export function IconBorderNone(props: IconProps) {
  return (
    <Svg {...props}>
      <rect
        x="3.02"
        y="3.02"
        width="9.96"
        height="9.96"
        rx="0.42"
        strokeWidth="1"
        strokeDasharray="1.5 1.7"
        opacity="0.55"
      />
      <path
        d="M 3.02 8 h 9.96 M 8 3.02 v 9.96"
        strokeWidth="1"
        strokeDasharray="1.5 1.7"
        opacity="0.55"
      />
    </Svg>
  )
}

/* ---------- Design ---------- */

export function IconTheme(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={1.5} y={11.5} s={11}>
        A
      </TextGlyph>
      <TextGlyph x={8.5} y={11.5} s={8}>
        a
      </TextGlyph>
      <path d="M2.5 13.8h11" strokeWidth="1" />
    </Svg>
  )
}

export function IconThemeFonts(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={2} y={12} s={11}>
        F
      </TextGlyph>
      <path d="M9.5 12 12 4.5 14.5 12M10.3 9.6h3.4" strokeWidth="1" />
    </Svg>
  )
}

export function IconThemeColors(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="4.89" cy="5.33" r="1.87" />
      <circle cx="11.11" cy="5.33" r="1.87" />
      <circle cx="4.89" cy="11.11" r="1.87" />
      <circle cx="11.11" cy="11.11" r="1.87" fill="currentColor" />
    </Svg>
  )
}

export function IconPageColor(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 8.56 3.36 4.4 7.52 a 1.04 1.04 0 0 0 0 1.44 l 2.64 2.64 a 1.04 1.04 0 0 0 1.44 0 l 4.16 -4.16 z" />
      <path d="M 8.56 3.36 7.2 4.8" />
      <path
        d="M 12.48 10.08 s 1.12 1.36 1.12 2.16 a 1.12 1.12 0 0 1 -2.24 0 c 0 -0.8 1.12 -2.16 1.12 -2.16 z"
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
      <path d="M 5.84 10.7 10.16 5.69" strokeWidth="1" opacity="0.45" />
    </Svg>
  )
}

export function IconPageBorders(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.02" width="9.96" height="9.96" rx="0.66" />
      <rect x="5.01" y="5.01" width="5.98" height="5.98" />
    </Svg>
  )
}

/* ---------- Layout ---------- */

export function IconMargins(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.15" y="3" width="7.7" height="10.01" rx="0.62" />
      <rect x="5.84" y="4.69" width="4.31" height="6.62" strokeDasharray="1.6 1.4" />
    </Svg>
  )
}

export function IconOrientation(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.2" y="4.8" width="6" height="8" rx="0.64" />
      <rect x="6" y="7.6" width="7.2" height="5.2" rx="0.64" fill="var(--surface, #fff)" />
      <path d="M 10.4 3.36 a 4 4 0 0 1 2.4 2.08 M 12.8 3.6 v 2 h -2" strokeWidth="1" />
    </Svg>
  )
}

export function IconPageSize(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.15" y="3" width="7.7" height="10.01" rx="0.62" />
      <path
        d="M 6.08 8 h 3.85 M 8 6.08 v 3.85 M 7 7 6.08 6.08 m 3.85 0 -0.92 0.92 m 0 2.01 0.92 0.92 m -3.85 0 0.92 -0.92"
        strokeWidth="1"
      />
    </Svg>
  )
}

export function IconColumns(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 2.99 3.45 h 4.1 M 2.99 5.73 h 4.1 M 2.99 8 h 4.1 M 2.99 10.27 h 4.1 M 2.99 12.55 h 4.1" />
      <path d="M 8.91 3.45 h 4.1 M 8.91 5.73 h 4.1 M 8.91 8 h 4.1 M 8.91 10.27 h 4.1 M 8.91 12.55 h 4.1" />
    </Svg>
  )
}

/* ---------- References ---------- */

export function IconToc(props: IconProps) {
  return (
    <Svg {...props}>
      {PAGE}
      <path d="M 6.08 5.69 h 3.85 M 7 7.46 h 2.93 M 7 9.23 h 2.93 M 6.08 11 h 3.85" />
    </Svg>
  )
}

export function IconRefresh(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 12.68 6.65 a 4.86 4.86 0 0 0 -9 -1.08 M 3.32 9.35 a 4.86 4.86 0 0 0 9 1.08" />
      <path d="M 12.95 3.05 v 2.7 h -2.7 M 3.05 12.95 v -2.7 h 2.7" />
    </Svg>
  )
}

export function IconFootnote(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={1.5} y={12} s={9}>
        AB
      </TextGlyph>
      <TextGlyph x={11.5} y={8} s={7} bold>
        1
      </TextGlyph>
    </Svg>
  )
}

export function IconEndnote(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={1.5} y={12} s={9}>
        AB
      </TextGlyph>
      <TextGlyph x={11.3} y={8} s={7} bold>
        n
      </TextGlyph>
    </Svg>
  )
}

export function IconCitation(props: IconProps) {
  // drawn quote marks, not a font glyph: text quotes hug the ascender line, so
  // the old TextGlyph version floated small at the top of the canvas
  return (
    <Svg {...props}>
      <path
        d="M6.9 4.9c-2 .7-3.3 2.3-3.3 4.3 0 1.3.9 2.3 2.1 2.3s2.1-1 2.1-2.2c0-1.2-.8-2.1-1.9-2.1.3-.8 1-1.4 1.9-1.8z"
        fill="currentColor"
        stroke="none"
      />
      <path
        d="M12.9 4.9c-2 .7-3.3 2.3-3.3 4.3 0 1.3.9 2.3 2.1 2.3s2.1-1 2.1-2.2c0-1.2-.8-2.1-1.9-2.1.3-.8 1-1.4 1.9-1.8z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconBook(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 8 4.09 C 6.96 3.3 5.22 2.95 3.22 3.13 v 9.05 c 2 -0.17 3.74 0.17 4.79 0.96 1.04 -0.78 2.78 -1.13 4.79 -0.96 V 3.13 c -2 -0.17 -3.74 0.17 -4.78 0.96 z" />
      <path d="M 8 4.09 v 9.05" />
    </Svg>
  )
}

export function IconCaption(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.02 5.1 h 7.06 L 12.98 8 l -2.9 2.91 H 3.02 z" />
      <circle cx="5.51" cy="8" r="0.75" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconIndex(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={1.8} y={6.5} s={6.5}>
        A
      </TextGlyph>
      <TextGlyph x={1.8} y={13.5} s={6.5}>
        B
      </TextGlyph>
      <path d="M8 4.5h6M8 8h6M8 11.5h6" />
    </Svg>
  )
}

/* ---------- Review ---------- */

export function IconWordCount(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={1.6} y={8} s={8}>
        123
      </TextGlyph>
      <path d="M2 11h12M2 13.5h8" />
    </Svg>
  )
}

export function IconSpellcheck(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={1.4} y={8.5} s={7.5}>
        abc
      </TextGlyph>
      <path d="M6 11.5 8.5 13.5 13 7.5" strokeWidth="1" />
    </Svg>
  )
}

export function IconSparkle(props: IconProps) {
  return (
    <Svg {...props}>
      <path
        d="M 8 3.03 C 8 5.78 10.22 8 12.97 8 C 10.22 8 8 10.22 8 12.97 C 8 10.22 5.78 8 3.03 8 C 5.78 8 8 5.78 8 3.03 Z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconWand(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.6 12.4 9.6 6.4" strokeWidth="1" />
      <path
        d="M 11.04 2.8 l 0.56 1.52 1.52 0.56 -1.52 0.56 -0.56 1.52 -0.56 -1.52 -1.52 -0.56 1.52 -0.56 z"
        fill="currentColor"
        stroke="none"
      />
      <path
        d="M 12.4 8.4 l 0.32 0.88 0.88 0.32 -0.88 0.32 -0.32 0.88 -0.32 -0.88 -0.88 -0.32 0.88 -0.32 z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconTranslate(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={1.2} y={9} s={8.5}>
        文
      </TextGlyph>
      <path d="M8.8 13.5 11.5 6.5 14.2 13.5M9.7 11.3h3.6" strokeWidth="1" />
    </Svg>
  )
}

export function IconTrackChanges(props: IconProps) {
  return (
    <Svg {...props}>
      {PAGE}
      <path d="M 6.08 6.08 h 3.85 M 6.08 8 h 2.31" />
      <path
        d="M 8.38 12 12.31 8.08 l 0.92 0.92 -3.93 3.93 -1.39 0.46 z"
        fill="var(--surface, #fff)"
      />
    </Svg>
  )
}

export function IconAccept(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.91 9.36 7.09 12.55 l 6.83 -7.28" />
    </Svg>
  )
}

export function IconReject(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m4.5 4.5 9 9M13.5 4.5l-9 9" />
    </Svg>
  )
}

export function IconCompare(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="4.15" width="4.24" height="7.7" rx="0.62" />
      <rect x="8.77" y="4.15" width="4.24" height="7.7" rx="0.62" />
      <path d="M 6.46 8 h 3.08 M 8.46 6.92 9.54 8 l -1.08 1.08" strokeWidth="1" />
    </Svg>
  )
}

export function IconLock(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.27" y="7.17" width="7.47" height="6.23" rx="0.83" />
      <path d="M 5.93 7.17 V 5.51 a 2.07 2.07 0 0 1 4.15 0 v 1.66" />
      <circle cx="8" cy="10.07" r="0.83" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/* ---------- View ---------- */

function Magnifier({ children }: { children?: ReactNode }) {
  return (
    <>
      <circle cx="7.15" cy="7.15" r="4.08" />
      <path d="M 10.21 10.21 13.1 13.1" strokeWidth="1" />
      {children}
    </>
  )
}

export function IconZoomOut(props: IconProps) {
  return (
    <Svg {...props}>
      <Magnifier>
        <path d="M 5.28 7.15 h 3.74" />
      </Magnifier>
    </Svg>
  )
}

export function IconZoomIn(props: IconProps) {
  return (
    <Svg {...props}>
      <Magnifier>
        <path d="M 5.28 7.15 h 3.74 M 7.15 5.28 v 3.74" />
      </Magnifier>
    </Svg>
  )
}

export function IconZoom100(props: IconProps) {
  return (
    <Svg {...props}>
      <Magnifier>
        <TextGlyph x={4.43} y={8.85} s={4.25} bold>
          1:1
        </TextGlyph>
      </Magnifier>
    </Svg>
  )
}

export function IconPageWidth(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.02" width="9.96" height="9.96" rx="0.66" />
      <path
        d="M 4.68 8 h 6.64 M 6.17 6.51 4.68 8 l 1.49 1.49 M 9.83 6.51 11.32 8 l -1.49 1.49"
        strokeWidth="1"
      />
    </Svg>
  )
}

export function IconWholePage(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.15" y="3" width="7.7" height="10.01" rx="0.62" />
      <path d="M 6.08 8 h 3.85 M 8 6.08 v 3.85" strokeWidth="1" />
    </Svg>
  )
}

export function IconAiPanel(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="3.76" width="10.01" height="8.47" rx="0.62" />
      <path d="M 9.39 3.76 v 8.47" />
      <path
        d="M 10.31 6.61 l 0.39 1 1 0.39 -1 0.39 -0.38 1 -0.38 -1 -1 -0.38 1 -0.38 z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconMoon(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 12.52 9.57 A 5.05 5.05 0 0 1 6.43 3.48 a 5.05 5.05 0 1 0 6.09 6.09 z" />
    </Svg>
  )
}

export function IconReadMode(props: IconProps) {
  return <IconBook {...props} />
}

export function IconOutlineView(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="3.65" cy="4.09" r="0.87" fill="currentColor" stroke="none" />
      <path d="M 5.83 4.09 h 6.96" />
      <circle cx="5.83" cy="8" r="0.87" fill="currentColor" stroke="none" />
      <path d="M 8 8 h 4.79" />
      <circle cx="5.83" cy="11.92" r="0.87" fill="currentColor" stroke="none" />
      <path d="M 8 11.92 h 4.79" />
    </Svg>
  )
}

export function IconRuler(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="6.08" width="10.01" height="3.85" rx="0.62" />
      <path
        d="M 5.31 6.08 v 1.54 M 7.23 6.08 v 2.31 M 9.16 6.08 v 1.54 M 11.08 6.08 v 2.31"
        strokeWidth="1"
      />
    </Svg>
  )
}

export function IconNavPane(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="3.76" width="10.01" height="8.47" rx="0.62" />
      <path d="M 6.46 3.76 v 8.47" />
      <path d="M 4 5.69 h 1.54 M 4 7.62 h 1.54 M 4 9.54 h 1.54" strokeWidth="1" />
    </Svg>
  )
}

export function IconSplit(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.02" width="9.96" height="9.96" rx="0.66" />
      <path d="M 3.02 8 h 9.96" strokeWidth="1" />
    </Svg>
  )
}

export function IconPrintLayout(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.54" y="3" width="6.93" height="10.01" rx="0.62" />
      <path
        d="M 6.08 5.31 h 3.85 M 6.08 7.23 h 3.85 M 6.08 9.16 h 3.85 M 6.08 11.08 h 2.31"
        strokeWidth="1"
      />
    </Svg>
  )
}

export function IconWebLayout(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="3.76" width="10.01" height="8.47" rx="0.62" />
      <path d="M 3 5.69 h 10.01" />
      <path d="M 4.54 7.62 h 6.93 M 4.54 9.16 h 6.93 M 4.54 10.7 h 4.62" strokeWidth="1" />
    </Svg>
  )
}

export function IconGridlines(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.02" width="9.96" height="9.96" rx="0.66" />
      <path
        d="M 3.02 6.34 h 9.96 M 3.02 9.66 h 9.96 M 6.34 3.02 v 9.96 M 9.66 3.02 v 9.96"
        strokeWidth="1"
      />
    </Svg>
  )
}

export function IconNewWindow(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="5.31" width="7.7" height="7.7" rx="0.62" />
      <path d="M 5.31 5.31 v -1.54 a 0.77 0.77 0 0 1 0.77 -0.77 h 6.16 a 0.77 0.77 0 0 1 0.77 0.77 v 6.16 a 0.77 0.77 0 0 1 -0.77 0.77 h -1.54" />
      <path d="M 6.85 9.16 h 3.08 M 8.39 7.62 v 3.08" />
    </Svg>
  )
}

export function IconArrangeAll(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="3.38" width="10.01" height="4" rx="0.62" />
      <rect x="3" y="8.62" width="10.01" height="4" rx="0.62" />
    </Svg>
  )
}

export function IconSwitchWindows(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="6.08" width="6.93" height="6.16" rx="0.62" />
      <path d="M 5.69 6.08 v -1.54 a 0.77 0.77 0 0 1 0.77 -0.77 h 5.78 a 0.77 0.77 0 0 1 0.77 0.77 v 5.39 a 0.77 0.77 0 0 1 -0.77 0.77 h -2.31" />
    </Svg>
  )
}

export function IconPosition(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="3" width="10.01" height="10.01" rx="0.77" />
      <rect x="5.69" y="5.69" width="4.62" height="4.62" />
    </Svg>
  )
}

export function IconWrapText(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="5.69" width="4.62" height="4.62" />
      <path d="M 9.16 3.38 h 3.85 M 9.16 5.69 h 3.85 M 9.16 8 h 3.85 M 9.16 10.31 h 3.85 M 3 12.62 h 10.01 M 3 3.38 h 4.62" />
    </Svg>
  )
}

export function IconDoc(props: IconProps) {
  return (
    <Svg {...props}>
      {PAGE}
      <path d="M 9.54 3 V 4.92 h 1.93" />
      <path d="M 6.08 6.84 h 3.85 M 6.08 8.77 h 3.85 M 6.08 10.7 h 2.7" />
    </Svg>
  )
}

/* ---------- AI panel ---------- */

export function IconSend(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.01 8 12.99 3.36 10.58 12.64 7.66 9.38 z" strokeLinejoin="round" />
      <path d="M 7.66 9.38 12.99 3.36" />
    </Svg>
  )
}

export function IconStop(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="3" width="10" height="10" rx="1.88" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconGear(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="8" cy="8" r="1.78" />
      <path d="M 8 2.98 v 1.62 M 8 11.4 v 1.62 M 13.02 8 h -1.62 M 4.6 8 h -1.62 M 11.56 4.44 l -1.13 1.13 M 5.57 10.43 l -1.13 1.13 M 11.56 11.56 10.43 10.43 M 5.57 5.57 4.44 4.44" />
    </Svg>
  )
}

/** collapse the left-docked AI panel: sheets-parity glyph (16-canvas, 1.2/1.3 stroke),
 *  self-contained so the shared Svg wrapper's pinned stroke doesn't alter its weight */
export function IconSidebarCollapse({ size = 20 }: IconProps) {
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
      <circle cx="8" cy="8" r="4.98" />
      <path d="M 8 5.34 V 8 l 1.91 1.33" />
    </Svg>
  )
}

export function IconPaperclip(props: IconProps) {
  return (
    <Svg {...props}>
      <path
        d="M 12.5 7.28 8.18 11.6 a 3.06 3.06 0 0 1 -4.32 -4.32 l 4.5 -4.5 a 2.07 2.07 0 0 1 2.88 2.88 l -4.5 4.5 a 0.99 0.99 0 0 1 -1.44 -1.44 l 4.14 -4.14"
        strokeLinejoin="round"
      />
    </Svg>
  )
}

export function IconNewChat(props: IconProps) {
  return (
    <Svg {...props}>
      <path
        d="M 12.68 7.32 v -2.55 A 1.44 1.44 0 0 0 11.23 3.33 H 4.77 a 1.44 1.44 0 0 0 -1.44 1.44 v 5.19 a 1.44 1.44 0 0 0 1.44 1.44 h 0.94 v 1.7 l 2.21 -1.7 h 1.11"
        strokeLinejoin="round"
      />
      <path d="M 11.57 9.19 v 3.4 M 9.87 10.89 h 3.4" />
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
      <path d="M4 2.5 12 8l-3.4.8L7 12.4 4 2.5Z" />
    </Svg>
  )
}

export function IconPen(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m3 13 .8-3L10.6 3.2a1.4 1.4 0 0 1 2 0l.2.2a1.4 1.4 0 0 1 0 2L6 12.2 3 13Z" />
      <path d="M9.6 4.2 11.8 6.4" />
    </Svg>
  )
}

export function IconHighlighterPen(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.68 9.4 10.6 4.47 a 1.21 1.21 0 0 1 1.77 0 l -0.84 -0.84 0.84 0.84 a 1.21 1.21 0 0 1 0 1.77 L 7.44 11.16 l -2.42 0.65 0.65 -2.42 Z" />
      <path d="M 3.35 13.58 h 9.3" strokeWidth="1" opacity="0.5" />
    </Svg>
  )
}

export function IconEraser(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m8.3 3.6 4.1 4.1a1.2 1.2 0 0 1 0 1.7L9.6 12.2H6.8L3.6 9a1.2 1.2 0 0 1 0-1.7l3-3a1.2 1.2 0 0 1 1.7 0Z" />
      <path d="M5.5 5.8 10.2 10.5" />
      <path d="M6.8 12.2h6.4" />
    </Svg>
  )
}

export function IconTextBox(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="4.7" width="10" height="6.6" rx="0.85" />
      <path d="M5.5 7.2h5" />
      <path d="M8 7.2v3.3" />
    </Svg>
  )
}

export function IconWordArt(props: IconProps) {
  return (
    <Svg {...props}>
      {/* stylized A with gradient effect hint */}
      <path d="M8 3 3.5 13h2.3l1-2.5h2.4l1 2.5h2.3L8 3Z" />
      <path d="M5.6 9.2h4.8" />
    </Svg>
  )
}

export function IconTrash(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.11 4.89h9.79M6.4 4.89V3.73a.62.62 0 0 1 .62-.62h1.96a.62.62 0 0 1 .62.62v1.16" />
      <path d="M4.44 4.89l.62 7.39a.89.89 0 0 0 .89.8h4.09a.89.89 0 0 0 .89-.8l.62-7.39" />
      <path d="M6.75 7.11v3.56M9.25 7.11v3.56" />
    </Svg>
  )
}

/** Thin dropdown chevron replacing the ▾ text glyph (same path as the slides ribbon's RbCaret);
 *  1.5 stroke on a 24 viewBox keeps the 1 : 16 stroke : canvas ratio. */
export function IconCaret({ size = 10 }: IconProps) {
  return (
    <svg
      className="rb-caret-svg"
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.6"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      <path d="M5.5 9.25 12 15.75l6.5-6.5" />
    </svg>
  )
}

export function IconPalette(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 12.98a4.98 4.98 0 1 1 4.98-4.98c0 2.44-1.74 2.49-2.74 2.49-.8 0-1.25.5-1.25 1.25 0 .7-.45 1.25-1 1.25Z" />
      <circle cx="8.83" cy="4.93" r="0.71" fill="currentColor" stroke="none" />
      <circle cx="11.07" cy="6.71" r="0.71" fill="currentColor" stroke="none" />
      <circle cx="6.09" cy="5.51" r="0.71" fill="currentColor" stroke="none" />
      <circle cx="4.93" cy="8.25" r="0.71" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconSort(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={1.5} y={7.2} s={7}>
        A
      </TextGlyph>
      <TextGlyph x={1.5} y={14.5} s={7}>
        Z
      </TextGlyph>
      <path d="M11.5 2.5V13M11.5 13 9.3 10.8M11.5 13l2.2-2.2" />
    </Svg>
  )
}

export function IconPilcrow(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 12.1 2.99 H 7.45 a 2.73 2.73 0 0 0 0 5.46 h 1.91 M 9.36 2.99 v 10.01 M 12.1 2.99 v 10.01" />
    </Svg>
  )
}

export function IconShading(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.99" y="3.91" width="10.01" height="8.19" rx="0.46" />
      <path
        d="M 2.99 7.55 6.64 3.91 M 2.99 11.18 10.27 3.91 M 5.27 12.1 13.01 4.36 M 8.91 12.1 13.01 8"
        opacity="0.55"
      />
    </Svg>
  )
}

export function IconCheck(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.8 8.6 6.2 12l7-7.5" />
    </Svg>
  )
}

export function IconCheckbox(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.99" y="2.99" width="10.01" height="10.01" rx="1.37" />
    </Svg>
  )
}

export function IconCheckboxChecked(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.99" y="2.99" width="10.01" height="10.01" rx="1.37" />
      <path d="M 5.45 8.27 l 1.91 2 3.37 -4.19" />
    </Svg>
  )
}

export function IconClose(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.5 3.5l9 9M12.5 3.5l-9 9" />
    </Svg>
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
