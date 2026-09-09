/**
 * Icon set for the html app: everything shared re-exports from the docs
 * ribbon library so glyph style stays uniform across the suite; the html-only
 * glyphs below follow the same 16-grid / pinned-stroke contract.
 */

import type { ReactNode } from 'react'

export {
  IconAlignCenter,
  IconAlignLeft,
  IconAlignRight,
  IconCopy,
  IconCrop,
  IconFlipH,
  IconFlipV,
  IconLink,
  IconPalette,
  IconRedo,
  IconRemoveBg,
  IconReplacePicture,
  IconRotateLeft,
  IconRotateRight,
  IconSave,
  IconSearch,
  IconSparkle,
  IconTrash,
  IconUndo,
  IconWand,
} from '../../../../docs/src/renderer/components/icons'

interface IconProps {
  size?: number
}

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

export function IconMoveUp(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 13V4.2M4.4 7.6 8 4l3.6 3.6" />
    </Svg>
  )
}

export function IconMoveDown(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 3v8.8M11.6 8.4 8 12 4.4 8.4" />
    </Svg>
  )
}

export function IconCode(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M5.2 4.6 2 8l3.2 3.4M10.8 4.6 14 8l-3.2 3.4M9.4 3.2 6.6 12.8" />
    </Svg>
  )
}

export function IconPreview(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.2" y="3" width="11.6" height="10" rx="1.4" />
      <path d="M2.2 6h11.6" />
    </Svg>
  )
}

export function IconSplitView(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.2" y="3" width="11.6" height="10" rx="1.4" />
      <path d="M8 3v10" />
    </Svg>
  )
}

export function IconSliders(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.5 4.5h11M2.5 8h11M2.5 11.5h11" />
      <circle cx="6" cy="4.5" r="1.4" fill="var(--surface)" />
      <circle cx="10.5" cy="8" r="1.4" fill="var(--surface)" />
      <circle cx="5" cy="11.5" r="1.4" fill="var(--surface)" />
    </Svg>
  )
}

export function IconChevronDown(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m4.5 6.5 3.5 3.5 3.5-3.5" />
    </Svg>
  )
}

export function IconExpand(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M9.5 2.5h4v4M13.5 2.5 9 7M6.5 13.5h-4v-4M2.5 13.5 7 9" />
    </Svg>
  )
}

export function IconPlay(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M4.5 3.2v9.6L12.5 8z" />
    </Svg>
  )
}

export function IconGlobe(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="8" cy="8" r="5.5" />
      <path d="M2.5 8h11M8 2.5c1.8 1.6 2.6 3.4 2.6 5.5S9.8 12 8 13.5C6.2 11.9 5.4 10.1 5.4 8S6.2 4.1 8 2.5Z" />
    </Svg>
  )
}

/** docs "AI Summarize" glyph (24-grid, drawn at the ribbon's big-button size) */
export function IconSummarize({ size = 24 }: IconProps) {
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
      <path d="M13.875 21H12H6.5C5.39543 21 4.5 20.1046 4.5 19V5C4.5 3.89543 5.39543 3 6.5 3H17.5C18.6046 3 19.5 3.89543 19.5 5V9V12V13" />
      <path d="M8.00001 7H16" />
      <path d="M8.00007 10.2032H14.0001" />
      <path d="M8.00007 13.4062H12.0001" />
      <path d="M17 14L17.2579 14.697C17.5961 15.611 17.7652 16.068 18.0986 16.4014C18.432 16.7348 18.889 16.9039 19.803 17.2421L20.5 17.5L19.803 17.7579C18.889 18.0961 18.432 18.2652 18.0986 18.5986C17.7652 18.932 17.5961 19.389 17.2579 20.303L17 21L16.7421 20.303C16.4039 19.389 16.2348 18.932 15.9014 18.5986C15.568 18.2652 15.111 18.0961 14.197 17.7579L13.5 17.5L14.197 17.2421C15.111 16.9039 15.568 16.7348 15.9014 16.4014C16.2348 16.068 16.4039 15.611 16.7421 14.697L17 14Z" />
    </svg>
  )
}
