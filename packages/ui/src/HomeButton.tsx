import type { SVGProps } from 'react'

function HomeIcon(props: SVGProps<SVGSVGElement>): React.JSX.Element {
  return (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" {...props}>
      <path d="m3 10 9-7 9 7" />
      <path d="M5 9.5V21h14V9.5" />
      <path d="M9 21v-6h6v6" />
    </svg>
  )
}

/** Returns to the web shell home page. Native Electron windows keep their shell navigation. */
export function HomeButton(): React.JSX.Element | null {
  if (typeof window === 'undefined' || !/^https?:$/.test(window.location.protocol)) return null

  const goHome = (): void => {
    window.location.assign(new URL('/', window.location.href).href)
  }

  return (
    <button
      type="button"
      className="home-button qa-btn"
      data-home-button="true"
      data-tip="返回主页"
      aria-label="返回主页"
      onClick={goHome}
    >
      <HomeIcon width="16" height="16" />
    </button>
  )
}
