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
  IconHighlight,
  IconUndo,
  IconRedo,
  IconCopy,
  IconRowInsertAbove,
  IconRowInsertBelow,
  IconColInsertLeft,
  IconColInsertRight,
  IconRowDelete,
  IconColDelete,
  IconTableDelete,
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

export function IconCallout(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3 2.8v10.4" />
      <path d="M5.8 5.2h7.2M5.8 8h5" />
    </Svg>
  )
}

export function IconToggle(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.2 4.2l3.2 2.4-3.2 2.4z" fill="currentColor" stroke="none" />
      <path d="M8.6 5.4h5M8.6 8h5M4 12h9.6" />
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
