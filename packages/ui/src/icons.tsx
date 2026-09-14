import type { ReactNode } from 'react'

export interface IconProps {
  size?: number
}

function Svg({ size = 16, children }: IconProps & { children: ReactNode }) {
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
      {children}
    </svg>
  )
}

export function IconSend(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.2 8 13.8 2.6 11 13.4 7.6 9.6z" strokeLinejoin="round" />
      <path d="M7.6 9.6 13.8 2.6" />
    </Svg>
  )
}

export function IconStop(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4" y="4" width="8" height="8" rx="1.5" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/** return/enter arrow (↵) for the icon-only send button */
export function IconEnter(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M13 3.5v4a2.5 2.5 0 0 1-2.5 2.5H3.5" />
      <path d="M6.5 7 3.5 10l3 3" />
    </Svg>
  )
}


export function IconClose(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.5 3.5 12.5 12.5" />
      <path d="M12.5 3.5 3.5 12.5" />
    </Svg>
  )
}

export function IconRetry(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M13.5 8a5.5 5.5 0 1 1-1.6-3.9" />
      <path d="M13.5 2.5v3h-3" />
    </Svg>
  )
}

export function IconEdit(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M11 2.5 13.5 5 5 13.5 2 14l.5-3z" strokeLinejoin="round" />
    </Svg>
  )
}

export function IconCheck(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3 8.5 6.5 12 13 4.5" strokeLinecap="round" strokeLinejoin="round" />
    </Svg>
  )
}

export function IconStopFilled(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.5" y="3.5" width="9" height="9" rx="1.5" fill="currentColor" stroke="currentColor" />
    </Svg>
  )
}

export function IconTool(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M10 2.5 13.5 6 6 13.5l-3.5.5.5-3.5z" strokeLinejoin="round" />
      <path d="M9 3.5 12.5 7" />
    </Svg>
  )
}

export function IconAttachment(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M10.5 5.5 6 10a2.5 2.5 0 1 0 3.5 3.5l6-6a4 4 0 1 0-5.5-5.5l-5.5 5.5" strokeLinejoin="round" />
    </Svg>
  )
}

export function IconWarning(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 2 14.5 13.5h-13z" strokeLinejoin="round" />
      <path d="M8 7v3" />
      <circle cx="8" cy="11.8" r="0.6" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconSparkle(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 2v4M8 10v4M2 8h4M10 8h4" strokeLinecap="round" />
    </Svg>
  )
}
