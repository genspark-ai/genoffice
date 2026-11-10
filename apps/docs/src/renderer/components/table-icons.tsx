import type { ReactNode } from 'react'
import type { CellHAlign, CellVAlign } from '../editor/table-ops'

interface IconProps {
  size?: number
}

function Svg({ size = 18, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth={(1.25 * 16) / size}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

/** a cell outline with three text strokes parked in one of the nine positions */
export function IconCellAlign({ v, h, size }: { v: CellVAlign; h: CellHAlign; size?: number }) {
  const y = v === 'top' ? 5.2 : v === 'center' ? 7.2 : 9.2
  const long = 6
  const short = 3.8
  const x1 = h === 'left' ? 4.4 : h === 'center' ? 8 - long / 2 : 11.6 - long
  const x2 = h === 'left' ? 4.4 : h === 'center' ? 8 - short / 2 : 11.6 - short
  return (
    <Svg size={size}>
      <rect x="2.5" y="2.9" width="11" height="10.2" rx="0.7" />
      <path d={`M ${x1} ${y} h ${long} M ${x2} ${y + 1.8} h ${short}`} />
    </Svg>
  )
}

export function IconTextDirection(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.5" y="2.9" width="11" height="10.2" rx="0.7" />
      <path d="M 5.2 10.6 V 5.4 M 5.2 5.4 L 6.9 10.6 M 5.2 5.4 L 3.5 10.6 M 4.1 8.9 h 2.2" />
      <path d="M 9.6 5 v 6 M 9.6 11 l -1.4 -1.4 M 9.6 11 l 1.4 -1.4" />
    </Svg>
  )
}

export function IconSplitTable(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 2.5 2.9 h 11 v 4 h -11 z M 2.5 9.4 h 11 v 4 h -11 z M 8 2.9 v 4 M 8 9.4 v 4" />
      <path d="M 1.6 8.15 h 12.8" strokeDasharray="1.6 1.4" />
    </Svg>
  )
}

export function IconDistributeRows(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.5" y="2.9" width="11" height="10.2" rx="0.7" />
      <path d="M 2.5 6.3 h 11 M 2.5 9.7 h 11" />
      <path d="M 14.8 4.2 v 7.6 M 14.1 4.9 l 0.7 -0.7 l 0.7 0.7 M 14.1 11.1 l 0.7 0.7 l 0.7 -0.7" />
    </Svg>
  )
}

export function IconDistributeColumns(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.5" y="2.9" width="11" height="10.2" rx="0.7" />
      <path d="M 6.2 2.9 v 10.2 M 9.8 2.9 v 10.2" />
      <path d="M 4.2 14.8 h 7.6 M 4.9 14.1 l -0.7 0.7 l 0.7 0.7 M 11.1 14.1 l 0.7 0.7 l -0.7 0.7" />
    </Svg>
  )
}

export function IconGridlines(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.5" y="2.9" width="11" height="10.2" rx="0.7" strokeDasharray="1.5 1.3" />
      <path d="M 8 2.9 v 10.2 M 2.5 8 h 11" strokeDasharray="1.5 1.3" />
    </Svg>
  )
}

export function IconInsertCells(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 2.5 2.9 h 7 v 10.2 h -7 z M 2.5 8 h 7 M 6 2.9 v 10.2" />
      <path d="M 12.6 6 v 4 M 10.6 8 h 4" />
    </Svg>
  )
}

export function IconDeleteCells(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 2.5 2.9 h 7 v 10.2 h -7 z M 2.5 8 h 7 M 6 2.9 v 10.2" />
      <path d="M 11.2 6.6 l 2.8 2.8 M 14 6.6 l -2.8 2.8" />
    </Svg>
  )
}

/** a grid with the caret's row filled: Table Layout ▸ Select */
export function IconSelectCells(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2" y="2.5" width="12" height="11" rx="1" />
      <path d="M2 6.2h12M2 9.8h12M6 2.5v11" />
      <rect
        x="2"
        y="6.2"
        width="12"
        height="3.6"
        fill="currentColor"
        opacity="0.35"
        stroke="none"
      />
    </Svg>
  )
}

/** a cell with an inset dashed box: Table Layout ▸ Cell Margins */
export function IconCellMargins(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2" y="2.5" width="12" height="11" rx="1" />
      <rect x="4.6" y="5" width="6.8" height="6" strokeDasharray="1.6 1.2" />
    </Svg>
  )
}
