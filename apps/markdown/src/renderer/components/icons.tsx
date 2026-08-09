/**
 * Icon set for the markdown app. Everything that exists in the docs ribbon
 * library is re-exported from there so glyph style stays uniform across the
 * suite; the handful of markdown-only glyphs below are drawn on the same
 * 16-grid / pinned-stroke contract. (Long term this belongs in packages/ui.)
 */

import type { ReactNode } from 'react'

export {
  IconBullets,
  IconNumbered,
  IconIndentDec,
  IconIndentInc,
  IconTable,
  IconPicture,
  IconLink,
  IconSave,
  IconUndo,
  IconRedo,
  IconCopy,
} from '../../../../docs/src/renderer/components/icons'

interface IconProps {
  size?: number
}

/** same pinned-stroke rule as the docs icon library */
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

export function IconTaskList(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="1.8" y="2.2" width="4.4" height="4.4" rx="1" />
      <path d="M3.2 4.4l1 1 1.7-1.9" />
      <path d="M8.8 4.4h5.4" />
      <rect x="1.8" y="9.4" width="4.4" height="4.4" rx="1" />
      <path d="M8.8 11.6h5.4" />
    </Svg>
  )
}

export function IconHr(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2 8h12" />
      <path d="M4.5 4.2h7M4.5 11.8h7" opacity="0.45" />
    </Svg>
  )
}

export function IconProperties(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.5 4.4h11M2.5 8h11M2.5 11.6h6.2" />
    </Svg>
  )
}

/* ── table-menu glyphs: insert = explicit "+", delete = explicit "×" ──
 * (redrawn locally — the docs arrow variants read as "move" at 15px) */

export function IconRowInsertAbove(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 1.7v4M6 3.7h4" />
      <rect x="2.2" y="7.8" width="11.6" height="6" rx="0.8" />
      <path d="M2.2 10.8h11.6" strokeWidth="1" />
    </Svg>
  )
}

export function IconRowInsertBelow(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.2" y="2.2" width="11.6" height="6" rx="0.8" />
      <path d="M2.2 5.2h11.6" strokeWidth="1" />
      <path d="M8 10.3v4M6 12.3h4" />
    </Svg>
  )
}

export function IconColInsertLeft(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M1.7 8h4M3.7 6v4" />
      <rect x="7.8" y="2.2" width="6" height="11.6" rx="0.8" />
      <path d="M10.8 2.2v11.6" strokeWidth="1" />
    </Svg>
  )
}

export function IconColInsertRight(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.2" y="2.2" width="6" height="11.6" rx="0.8" />
      <path d="M5.2 2.2v11.6" strokeWidth="1" />
      <path d="M10.3 8h4M12.3 6v4" />
    </Svg>
  )
}

export function IconRowDelete(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.4 2.6h9.2M3.4 13.4h9.2" opacity="0.4" strokeWidth="1" />
      <rect x="2.2" y="5" width="11.6" height="6" rx="0.8" />
      <path d="M6.8 6.8l2.4 2.4M9.2 6.8l-2.4 2.4" />
    </Svg>
  )
}

export function IconColDelete(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.6 3.4v9.2M13.4 3.4v9.2" opacity="0.4" strokeWidth="1" />
      <rect x="5" y="2.2" width="6" height="11.6" rx="0.8" />
      <path d="M6.8 6.8l2.4 2.4M9.2 6.8l-2.4 2.4" />
    </Svg>
  )
}

export function IconTableDelete(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2" y="2.6" width="9" height="8.4" rx="0.8" />
      <path d="M2 5.4h9M6.5 2.6v8.4" strokeWidth="1" />
      <path d="M10.7 10.4l3.2 3.2M13.9 10.4l-3.2 3.2" />
    </Svg>
  )
}

export function IconHeaderRow(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2" y="3" width="12" height="10" rx="1" />
      <path d="M2 6.4h12M6.7 6.4V13M11.3 6.4V13" />
      <path d="M2.6 3.6h10.8v2.3H2.6z" fill="currentColor" stroke="none" opacity="0.35" />
    </Svg>
  )
}

export function IconInlineCode(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M5.4 4.6L2.4 8l3 3.4M10.6 4.6l3 3.4-3 3.4" />
    </Svg>
  )
}

export function IconQuoteMark(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3 4.5v7" />
      <path d="M6.4 5h6.8M6.4 8h6.8M6.4 11h4.4" />
    </Svg>
  )
}
