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

/** gear icon for AI settings */
export function IconSettings(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="8" cy="8" r="2" />
      <path d="M7.05.5h1.9l.4 1.2a5.5 5.5 0 0 1 1.1.6l1.1-.6.9 1.6-1 .8a5.5 5.5 0 0 1 0 1.2l1 .8-.9 1.6-1.1-.6a5.5 5.5 0 0 1-1.1.6l-.4 1.2h-1.9l-.4-1.2a5.5 5.5 0 0 1-1.1-.6l-1.1.6-.9-1.6 1-.8a5.5 5.5 0 0 1 0-1.2l-1-.8.9-1.6 1.1.6a5.5 5.5 0 0 1 1.1-.6l.4-1.2Z" />
    </Svg>
  )
}
