/** Transient feedback for user-triggered actions (save etc.): a fixed
 * top-center pill that auto-dismisses. The status bar stays the durable log.
 * Trigger via showToast from './toast-bus' (kept component-only here so React
 * Fast Refresh works in dev). */
import { useEffect, useState } from 'react'
import { setToastEmitter, type ToastData } from './toast-bus'

export function ToastHost() {
  const [toast, setToast] = useState<ToastData | null>(null)
  const [visible, setVisible] = useState(false)
  useEffect(() => {
    let hideTimer: number | undefined
    let clearTimer: number | undefined
    let raf = 0
    setToastEmitter((next) => {
      window.clearTimeout(hideTimer)
      window.clearTimeout(clearTimer)
      window.cancelAnimationFrame(raf)
      setToast(next)
      // Mount hidden first: CSS transitions don't run on initial mount, so
      // the show class lands a frame later for the fade/slide-in to play.
      setVisible(false)
      raf = window.requestAnimationFrame(() => {
        raf = window.requestAnimationFrame(() => setVisible(true))
      })
      const shownMs = next.kind === 'error' ? 4000 : 2000
      hideTimer = window.setTimeout(() => setVisible(false), shownMs)
      // keep the node mounted through the fade-out transition
      clearTimer = window.setTimeout(() => setToast(null), shownMs + 200)
    })
    return () => {
      setToastEmitter(null)
      window.clearTimeout(hideTimer)
      window.clearTimeout(clearTimer)
      window.cancelAnimationFrame(raf)
    }
  }, [])
  if (!toast) return null
  return (
    <div className={`app-toast ${toast.kind}${visible ? ' show' : ''}`} role="status">
      <svg
        className="app-toast-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth={1.5}
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden
      >
        {toast.kind === 'success' ? (
          <>
            <circle cx="12" cy="12" r="9" />
            <path d="m8.2 12.3 2.6 2.6 5-5" />
          </>
        ) : (
          <>
            <circle cx="12" cy="12" r="9" />
            <path d="M12 7.5v5.3" />
            <path d="M12 16.4v.1" />
          </>
        )}
      </svg>
      {toast.text}
    </div>
  )
}
