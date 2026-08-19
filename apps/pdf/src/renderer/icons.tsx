import type { ReactElement, ReactNode } from 'react'

// ── ribbon icons (aligned with slides' rb-big visual language) ──

/** Constant painted stroke instead of proportional scaling — same rule as the
 *  slides icons: ~1.5px lines at 20px+, ~1.25px on 13-19px glyphs, ~1.1px below.
 *  stroke-width is in 24-canvas units: units = painted-px × 24 / rendered-px. */
export function pinnedStroke(size: number): number {
  const painted = size >= 20 ? 1.5 : size >= 13 ? 1.25 : 1.1
  return (painted * 24) / size
}

export function Icon({
  size = 28,
  children,
}: {
  size?: number
  children: ReactNode
}): ReactElement {
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

export const IconThumbs = () => (
  <Icon>
    <rect x="4.5" y="5" width="6" height="6.5" rx="1" />
    <rect x="4.5" y="14" width="6" height="5" rx="1" />
    <path d="M14 6 L19.5 6 M14 10 L19.5 10 M14 15 L19.5 15 M14 18 L19.5 18" />
  </Icon>
)
export const IconHighlight = () => (
  <Icon>
    <path d="M6.04 15.09 L13.54 7.59 L16.63 10.68 L9.13 18.18 L6.04 18.18 L6.04 15.09 Z" />
    <path d="M13.54 6.71 L15.75 4.5 L18.84 7.59 L16.63 9.79" />
    <path d="M5.16 19.5 L17.51 19.5" strokeWidth={2.2} />
  </Icon>
)
export const IconUnderline = () => (
  <Icon>
    <path d="M7.56 4.5 L7.56 11.17 A 4.44 4.44 0 0 0 16.44 11.17 L16.44 4.5" />
    <path d="M5.89 19.5 L18.11 19.5" strokeWidth={1.85} />
  </Icon>
)
export const IconStrike = () => (
  <Icon>
    <path d="M16.29 7.21 C15.72 5.52 14.03 4.5 12 4.5 C9.52 4.5 7.71 5.85 7.71 7.88 C7.71 9.46 8.73 10.36 10.65 10.93" />
    <path d="M8.62 16.91 C9.18 18.48 10.87 19.5 13.02 19.5 C15.5 19.5 17.41 18.26 17.41 16.12 C17.41 15.44 17.3 14.88 16.96 14.42" />
    <path d="M4.67 12.28 L19.33 12.28" strokeWidth={1.65} />
  </Icon>
)
export const IconEditText = () => (
  <Icon>
    <path d="M5 5.5 L14.5 5.5 M9.75 5.5 L9.75 15.5 M7.5 15.5 L12 15.5" />
    <path d="M16.7 10.3 L19.7 13.3 L13.4 19.6 L10.4 20.1 L10.9 17.1 Z" />
  </Icon>
)
export const IconInk = () => (
  <Icon>
    <path d="M16.15 4.85 L19.15 7.85 L8.9 18.1 L4.9 19.1 L5.9 15.1 Z" />
    <path d="M14.15 6.85 L17.15 9.85" />
  </Icon>
)
export const IconRect = () => (
  <Icon>
    <rect x="4.5" y="6.64" width="15" height="10.71" rx="1.07" />
  </Icon>
)
export const IconEllipse = () => (
  <Icon>
    <ellipse cx="12" cy="12" rx="7.5" ry="5.57" />
  </Icon>
)
export const IconArrow = () => (
  <Icon>
    <path d="M4.5 19.2 L19.5 4.8" />
    <path d="M12.9 4.8 L19.5 4.8 L19.5 11.4" />
  </Icon>
)
export const IconNote = () => (
  <Icon>
    <path d="M4.5 5.57 L19.5 5.57 L19.5 15.21 L10.93 15.21 L6.64 18.43 L6.64 15.21 L4.5 15.21 Z" />
    <path d="M8.25 9.32 L15.75 9.32 M8.25 12 L13.07 12" />
  </Icon>
)
export const IconSign = () => (
  <Icon>
    <path d="M5.5 15.1 C7.8 12.3 9.5 9 9.2 7 C9 5.7 7.9 5.9 7.6 7.4 C7.2 9.6 8.6 13.4 10.5 14.9 C12 16.1 13.9 15.3 14.7 13.8 C15.1 13 15.9 13 16.3 13.8 C16.7 14.7 17.7 15 18.5 14.4" />
    <path d="M4.75 18.6 L19.25 18.6" />
  </Icon>
)
export const IconPreviousField = () => (
  <Icon>
    <rect x="6" y="4.5" width="12" height="15" rx="1.5" />
    <path d="M14.5 8 L10.5 12 L14.5 16" />
  </Icon>
)
export const IconNextField = () => (
  <Icon>
    <rect x="6" y="4.5" width="12" height="15" rx="1.5" />
    <path d="M9.5 8 L13.5 12 L9.5 16" />
  </Icon>
)
export const IconCompleteForm = () => (
  <Icon>
    <rect x="4.5" y="5" width="15" height="14" rx="1.5" />
    <path d="M8 12 L10.8 14.8 L16.5 9" />
  </Icon>
)
export const IconFormText = () => (
  <Icon>
    <path d="M5 6 H19 M12 6 V19 M8.5 19 H15.5" />
  </Icon>
)
export const IconFormCheck = () => (
  <Icon>
    <path d="M4.5 12.5 L9.5 17.5 L19.5 6.5" />
  </Icon>
)
export const IconFormCross = () => (
  <Icon>
    <path d="M6 6 L18 18 M18 6 L6 18" />
  </Icon>
)
export const IconExportImg = () => (
  <Icon>
    <rect x="4.5" y="6.75" width="15" height="10.5" rx="1" />
    <circle cx="9" cy="10.55" r="1.2" />
    <path d="M4.8 14.95 L9 11.75 L12.4 14.35 L15 12.35 L19.2 15.75" />
  </Icon>
)
export const IconInsertImage = () => (
  <Icon>
    <rect x="4.5" y="6" width="11.5" height="9.5" rx="1" />
    <circle cx="8" cy="9.2" r="1.1" />
    <path d="M4.8 13.6 L8 11.2 L11 13.4 L13.2 11.8 L15.8 13.9" />
    <path d="M18.6 13.4 V19 M15.8 16.2 H21.4" />
  </Icon>
)
export const IconEditImage = () => (
  <Icon>
    <rect x="4.5" y="6" width="12.5" height="10" rx="1" />
    <circle cx="8.3" cy="9.4" r="1.1" />
    <path d="M4.8 14 L8.3 11.4 L11.5 13.7 L13.8 12" />
    <path d="M14.2 18.9 L19.7 13.4 A1.06 1.06 0 0 0 18.2 11.9 L12.7 17.4 L12.2 19.4 Z" />
  </Icon>
)
export const IconNight = () => (
  <Icon>
    <path d="M19.5 13.48 A 7.58 7.58 0 0 1 10.52 4.5 A 7.58 7.58 0 1 0 19.5 13.48 Z" />
  </Icon>
)
export const IconSpread = () => (
  <Icon>
    <rect x="4.5" y="6" width="6.5" height="12" rx="1" />
    <rect x="13" y="6" width="6.5" height="12" rx="1" />
  </Icon>
)
export const IconSinglePage = () => (
  <Icon>
    <rect x="6.81" y="4.5" width="10.38" height="15" rx="1.15" />
  </Icon>
)
export const IconWatermark = () => (
  <Icon>
    <rect x="4.5" y="5.04" width="15" height="13.93" rx="1.07" />
    <path d="M7.71 15.75 L15.75 7.71" />
    <path d="M7.71 11.46 L11.46 7.71 M12.54 15.75 L16.29 12" />
  </Icon>
)
export const IconProps = () => (
  <Icon>
    <circle cx="12" cy="12" r="7.5" />
    <path d="M12 10.96 L12 15.65" />
    <circle cx="12" cy="8.46" r="0.94" fill="currentColor" stroke="none" />
  </Icon>
)
export const IconRotateL = () => (
  <Icon>
    <path d="M8.28 10.3 L4.53 10.3 L4.53 6.55" />
    <path d="M4.75 9.98 A 7.5 7.5 0 1 1 4.53 12.98" />
  </Icon>
)
export const IconRotateR = () => (
  <Icon>
    <path d="M15.72 10.3 L19.47 10.3 L19.47 6.55" />
    <path d="M19.25 9.98 A 7.5 7.5 0 1 0 19.47 12.98" />
  </Icon>
)
export const IconDeletePage = () => (
  <Icon>
    <path d="M7.7 4.5 H13.7 L17.2 8 V18.5 A1 1 0 0 1 16.2 19.5 H7.7 A1 1 0 0 1 6.7 18.5 V5.5 A1 1 0 0 1 7.7 4.5 Z" />
    <path d="M13.7 4.5 V8 H17.2" />
    <path d="M9.7 11.75 L14.2 16.25 M14.2 11.75 L9.7 16.25" />
  </Icon>
)
export const IconExtract = () => (
  <Icon>
    <path d="M7.2 4.5 H13.2 L16.7 8 V11.5" />
    <path d="M6.2 5.5 V18.5 A1 1 0 0 0 7.2 19.5 H11.2" />
    <path d="M14.95 13.5 V19 M12.2 16.5 L14.95 19.25 L17.7 16.5" />
  </Icon>
)
export const IconInsertPdf = () => (
  <Icon>
    <path d="M7.7 4.5 H13.7 L17.2 8 V18.5 A1 1 0 0 1 16.2 19.5 H7.7 A1 1 0 0 1 6.7 18.5 V5.5 A1 1 0 0 1 7.7 4.5 Z" />
    <path d="M13.7 4.5 V8 H17.2" />
    <path d="M11.95 11 V17 M8.95 14 H14.95" />
  </Icon>
)
export const IconInsertBlank = () => (
  <Icon>
    <rect x="6.7" y="4.5" width="10.5" height="15" rx="1" strokeDasharray="2.2 1.8" />
    <path d="M11.95 9.5 V14.5 M9.45 12 H14.45" />
  </Icon>
)
export const IconRotateAll = () => (
  <Icon>
    <rect x="8.5" y="8.5" width="7" height="9" rx="0.8" />
    <path d="M6.2 6.2 A 8.2 8.2 0 0 1 17.8 6.2" />
    <path d="M17.8 3.4 L17.8 6.4 L14.8 6.4" />
    <path d="M17.8 17.8 A 8.2 8.2 0 0 1 6.2 17.8" />
    <path d="M6.2 20.6 L6.2 17.6 L9.2 17.6" />
  </Icon>
)
export const IconReverse = () => (
  <Icon>
    <path d="M9 5.5 L9 18 M6.5 8 L9 5.5 L11.5 8" />
    <path d="M15 6 L15 18.5 M12.5 16 L15 18.5 L17.5 16" />
  </Icon>
)
export const IconSplitPdf = () => (
  <Icon>
    <rect x="4.5" y="6.5" width="6.2" height="11" rx="0.8" />
    <rect x="13.3" y="6.5" width="6.2" height="11" rx="0.8" />
    <path d="M12 4.6 V7 M12 9.4 V11.8 M12 14.2 V16.6 M12 19 V19.4" />
  </Icon>
)
export const IconMergePdf = () => (
  <Icon>
    <path d="M4.5 6.5 H8.5 V17.5 H4.5 Z" />
    <path d="M15.5 6.5 H19.5 V17.5 H15.5 Z" />
    <path d="M10 9 L12 12 L10 15 M14 9 L12 12 L14 15" />
  </Icon>
)
export const IconMergePages = () => (
  <Icon>
    <rect x="5" y="4.5" width="14" height="15" rx="1" />
    <path d="M12 4.5 V19.5 M5 12 H19" />
  </Icon>
)

export const IconReplacePages = () => (
  <Icon>
    <rect x="4.5" y="7.5" width="9" height="12" rx="1" />
    <rect x="10.5" y="4.5" width="9" height="12" rx="1" strokeDasharray="2.2 1.8" />
  </Icon>
)
export const IconSplitPages = () => (
  <Icon>
    <rect x="5" y="4.5" width="14" height="15" rx="1" />
    <path d="M12 4.5 V19.5" strokeDasharray="2.2 1.8" />
    <path d="M8.2 12 L5.8 12 M6.6 10.8 L5.4 12 L6.6 13.2 M15.8 12 L18.2 12 M17.4 10.8 L18.6 12 L17.4 13.2" />
  </Icon>
)
export const IconCropPages = () => (
  <Icon>
    <path d="M7.5 4.5 V16.5 H19.5" />
    <path d="M4.5 7.5 H16.5 V19.5" />
  </Icon>
)
export const IconPageSize = () => (
  <Icon>
    <rect x="4.5" y="4.5" width="15" height="15" rx="1" />
    <rect x="8.5" y="8.5" width="7" height="9" rx="0.8" strokeDasharray="2 1.6" />
  </Icon>
)
export const IconFitWidth = () => (
  <Icon>
    <path d="M4.5 5.57 L4.5 18.43 M19.5 5.57 L19.5 18.43" />
    <path d="M7.71 12 L16.29 12 M10.07 9.64 L7.71 12 L10.07 14.36 M13.93 9.64 L16.29 12 L13.93 14.36" />
  </Icon>
)
export const IconFitPage = () => (
  <Icon>
    <rect x="7" y="4.5" width="10" height="15" rx="1" />
    <path d="M12 8 L12 16 M9.8 10.2 L12 8 L14.2 10.2 M9.8 13.8 L12 16 L14.2 13.8" />
  </Icon>
)
export const IconOutline = () => (
  <Icon>
    <path d="M4.84 4.78 L19.5 4.78 M8.22 9.29 L19.5 9.29 M8.22 13.8 L19.5 13.8 M11.61 18.32 L19.5 18.32" />
    <circle cx="5.4" cy="9.29" r="0.9" fill="currentColor" stroke="none" />
    <circle cx="5.4" cy="13.8" r="0.9" fill="currentColor" stroke="none" />
    <circle cx="8.79" cy="18.32" r="0.9" fill="currentColor" stroke="none" />
  </Icon>
)
export const IconDrawColor = () => (
  <Icon>
    <path d="M12 4.5 C14.2 7.3 17.25 9.2 17.25 12.4 C17.25 15.4 14.9 17.5 12 17.5 C9.1 17.5 6.75 15.4 6.75 12.4 C6.75 9.2 9.8 7.3 12 4.5 Z" />
  </Icon>
)
/* dropdown chevron, same glyph as slides' RbCaret */
export const RbCaret = () => (
  <svg className="rb-caret" width="10" height="10" viewBox="0 0 24 24" fill="none" aria-hidden>
    <path
      d="M5.5 9.25 12 15.75l6.5-6.5"
      stroke="currentColor"
      strokeWidth="2.6"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
)
export const IconSearch = () => (
  <Icon>
    <circle cx="10.61" cy="10.61" r="6.11" />
    <path d="M15.28 15.28 L19.5 19.5" />
  </Icon>
)
export const IconPrint = () => (
  <Icon>
    <path d="M7.71 8.79 L7.71 4.5 L16.29 4.5 L16.29 8.79" />
    <rect x="5.04" y="8.79" width="13.93" height="6.96" rx="1.07" />
    <path d="M7.71 13.07 L16.29 13.07 L16.29 19.5 L7.71 19.5 Z" />
  </Icon>
)
/** Design-supplied glyphs on the 1:16 stroke:canvas ratio (24-canvas / 1.5 stroke),
 *  geometry shared with docs/slides/sheets. */
export const IconRatio = ({ size = 16, children }: { size?: number; children: ReactNode }) => (
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
export const IconUndo = () => (
  <IconRatio>
    <path d="M5.91026 4L2.5 7.14791L5.91026 10.8205" />
    <path d="M3.96154 7.41028H15.1636C18.5169 7.41028 21.3646 10.1484 21.4953 13.5C21.6334 17.0416 18.707 20.0769 15.1636 20.0769H6.88384" />
  </IconRatio>
)
export const IconRedo = () => (
  <IconRatio>
    <path d="M18.0897 4L21.5 7.14791L18.0897 10.8205" />
    <path d="M20.0385 7.41028H8.83636C5.4831 7.41028 2.63537 10.1484 2.5047 13.5C2.36657 17.0416 5.29296 20.0769 8.83636 20.0769H17.1162" />
  </IconRatio>
)
export const IconSave = () => (
  <IconRatio>
    <path d="M3 4.5C3 3.67158 3.67158 3 4.5 3H17.1407L21 6.60325V19.5C21 20.3285 20.3285 21 19.5 21H4.5C3.67158 21 3 20.3285 3 19.5V4.5Z" />
    <path d="M12.0042 3L12 6.6923C12 6.86225 11.7761 7 11.5 7H7.5C7.22385 7 7 6.86225 7 6.6923V3H12.0042Z" />
    <path d="M7 13H17" />
    <path d="M7 17H12.0042" />
  </IconRatio>
)

// ── selection-popup icons (14px; bring-forward / send-backward / trash) ──

export const IconLayerUp = () => (
  <Icon size={18}>
    <rect x="9.5" y="4.5" width="10" height="10" rx="1" />
    <path d="M14.5 19.5 H5.5 A1 1 0 0 1 4.5 18.5 V9.5" />
  </Icon>
)
export const IconLayerDown = () => (
  <Icon size={18}>
    <rect x="4.5" y="9.5" width="10" height="10" rx="1" />
    <path d="M9.5 4.5 H18.5 A1 1 0 0 1 19.5 5.5 V14.5" />
  </Icon>
)
export const IconTrash = () => (
  <Icon size={18}>
    <path d="M4.5 6.5 H19.5" />
    <path d="M9 6.5 V5 A1 1 0 0 1 10 4 H14 A1 1 0 0 1 15 5 V6.5" />
    <path d="M6.5 6.5 L7.3 18.6 A1.4 1.4 0 0 0 8.7 19.9 H15.3 A1.4 1.4 0 0 0 16.7 18.6 L17.5 6.5" />
    <path d="M10.2 10 V16 M13.8 10 V16" />
  </Icon>
)
export const IconRotateCw = () => (
  <Icon size={18}>
    <path d="M18.5 8.5 A7.5 7.5 0 1 0 19.5 12" />
    <path d="M19 4 V8.5 H14.5" />
  </Icon>
)
export const IconRotateCcw = () => (
  <Icon size={18}>
    <path d="M5.5 8.5 A7.5 7.5 0 1 1 4.5 12" />
    <path d="M5 4 V8.5 H9.5" />
  </Icon>
)
export const IconSwapImage = () => (
  <Icon size={18}>
    <rect x="4" y="6.5" width="11" height="11" rx="1" />
    <circle cx="7.5" cy="10" r="1.2" />
    <path d="M4.5 16 L8.5 12.5 L11 15 L12.5 13.5 L15 16" />
    <path d="M17 4.5 H19 A1 1 0 0 1 20 5.5 V13" />
    <path d="M18.2 11.2 L20 13 L21.8 11.2" />
  </Icon>
)
export const IconFlipH = () => (
  <Icon size={18}>
    <path d="M12 3.5 V20.5" strokeDasharray="2.6 2.2" />
    <path d="M8.5 7 V17 L3.5 17 Z" />
    <path d="M15.5 7 V17 L20.5 17 Z" />
  </Icon>
)
export const IconFlipV = () => (
  <Icon size={18}>
    <path d="M3.5 12 H20.5" strokeDasharray="2.6 2.2" />
    <path d="M7 8.5 H17 L17 3.5 Z" />
    <path d="M7 15.5 H17 L17 20.5 Z" />
  </Icon>
)
export const IconCrop = () => (
  <Icon size={18}>
    <path d="M7 3.5 V17 H20.5" />
    <path d="M3.5 7 H17 V20.5" />
  </Icon>
)
export const IconCutout = () => (
  <Icon size={18}>
    <path d="M13.5 6.5 L17.5 10.5 L8 20 H4 V16 Z" />
    <path d="M16 4 L20 8" />
    <path d="M18.5 12.5 L19.4 14.6 L21.5 15.5 L19.4 16.4 L18.5 18.5 L17.6 16.4 L15.5 15.5 L17.6 14.6 Z" />
  </Icon>
)
export const IconOpacity = () => (
  <Icon size={18}>
    <path d="M12 3.5 C12 3.5 5.5 10 5.5 14.5 A6.5 6.5 0 0 0 18.5 14.5 C18.5 10 12 3.5 12 3.5 Z" />
    <path d="M12 18.2 A3.7 3.7 0 0 1 8.3 14.5" />
  </Icon>
)

/** One-click AI feature glyphs (same 24-canvas/1.5-stroke artwork as the docs ribbon) */
export const IconAiSummarize = () => (
  <svg
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
    <path
      d="M17 14L17.2579 14.697C17.5961 15.611 17.7652 16.068 18.0986 16.4014C18.432 16.7348 18.889 16.9039 19.803 17.2421L20.5 17.5L19.803 17.7579C18.889 18.0961 18.432 18.2652 18.0986 18.5986C17.7652 18.932 17.5961 19.389 17.2579 20.303L17 21L16.7421 20.303C16.4039 19.389 16.2348 18.932 15.9014 18.5986C15.568 18.2652 15.111 18.0961 14.197 17.7579L13.5 17.5L14.197 17.2421C15.111 16.9039 15.568 16.7348 15.9014 16.4014C16.2348 16.068 16.4039 15.611 16.7421 14.697L17 14Z"
      strokeLinejoin="round"
    />
  </svg>
)
export const IconAiKeyPoints = () => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden
  >
    <path d="M4 5H20" />
    <path d="M4 9H16" />
    <path d="M4 13H11" />
    <path d="M4 17H10" />
    <path
      d="M17 14L17.2579 14.697C17.5961 15.611 17.7652 16.068 18.0986 16.4014C18.432 16.7348 18.889 16.9039 19.803 17.2421L20.5 17.5L19.803 17.7579C18.889 18.0961 18.432 18.2652 18.0986 18.5986C17.7652 18.932 17.5961 19.389 17.2579 20.303L17 21L16.7421 20.303C16.4039 19.389 16.2348 18.932 15.9014 18.5986C15.568 18.2652 15.111 18.0961 14.197 17.7579L13.5 17.5L14.197 17.2421C15.111 16.9039 15.568 16.7348 15.9014 16.4014C16.2348 16.068 16.4039 15.611 16.7421 14.697L17 14Z"
      strokeLinejoin="round"
    />
  </svg>
)
