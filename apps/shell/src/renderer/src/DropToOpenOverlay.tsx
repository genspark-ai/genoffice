import { useEffect, useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n } from './locale'
import { OPEN_LOCAL_EXTENSIONS } from './home-shared'

// ── Drop-to-open overlay ────────────────────────────────

/**
 * Full-window affordance while OS files hover over Home. Purely visual — the
 * actual open is owned by the preload drop bridge (installDropOpenBridge), so
 * this overlay stays pointer-events:none and never handles events itself.
 * Visibility tracks a dragenter/dragleave depth counter: `dragover` stops
 * being delivered while the cursor is stationary (macOS), so a debounce would
 * hide the overlay mid-drag. Enter fires before the matching leave when
 * moving between elements, so the depth never dips to zero inside the window.
 */
export function DropToOpenOverlay(): ReactElement | null {
  const [visible, setVisible] = useState(false)
  useEffect(() => {
    let depth = 0
    const hasFiles = (ev: DragEvent): boolean => ev.dataTransfer?.types.includes('Files') ?? false
    // NB: the preload drop bridge also listens here and cancels file drags, so
    // defaultPrevented can't discriminate anything at this layer — only zones
    // that stopPropagation (none on Home) would keep us out entirely.
    const onDragEnter = (ev: DragEvent) => {
      if (!hasFiles(ev)) return
      depth += 1
      setVisible(true)
    }
    const onDragLeave = (ev: DragEvent) => {
      if (!hasFiles(ev)) return
      depth = Math.max(0, depth - 1)
      if (depth === 0) setVisible(false)
    }
    // drop/blur reset the depth outright: leaving the window mid-drag can eat
    // a dragleave, and a stuck overlay would be worse than a re-shown one
    const onHide = () => {
      depth = 0
      setVisible(false)
    }
    window.addEventListener('dragenter', onDragEnter)
    window.addEventListener('dragleave', onDragLeave)
    window.addEventListener('drop', onHide)
    window.addEventListener('blur', onHide)
    return () => {
      window.removeEventListener('dragenter', onDragEnter)
      window.removeEventListener('dragleave', onDragLeave)
      window.removeEventListener('drop', onHide)
      window.removeEventListener('blur', onHide)
    }
  }, [])
  const { t } = useI18n()
  if (!visible) return null
  return (
    <div className="home-drop-overlay" aria-hidden="true">
      <div className="home-drop-card">
        <svg width="40" height="40" viewBox="0 0 24 24" fill="none" aria-hidden="true">
          <path
            d="M12 3.5v11M7.5 10.5l4.5 4.5 4.5-4.5"
            stroke="currentColor"
            strokeWidth="1.6"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
          <path
            d="M4 16.5v2A1.5 1.5 0 0 0 5.5 20h13a1.5 1.5 0 0 0 1.5-1.5v-2"
            stroke="currentColor"
            strokeWidth="1.6"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
        <h2>{t('dropToOpenTitle')}</h2>
        <p>{OPEN_LOCAL_EXTENSIONS}</p>
      </div>
    </div>
  )
}
