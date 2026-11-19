import React, { useState } from 'react'
import type { Lang } from '@genoffice/i18n'
import { AI_PANEL_SIDE_LABELS } from './strings-ai-panel-side'
import type { AiPanelPrefs, AiPanelSide } from './ai-panel-prefs'
import { applyAiPanelPrefs, useAiPanelPrefs } from './ai-panel-prefs-store'

export function AiPanelSideButton({
  lang,
  onMove,
}: {
  lang: Lang
  onMove: (side: AiPanelSide) => Promise<AiPanelPrefs>
}) {
  const { side } = useAiPanelPrefs()
  const [pending, setPending] = useState(false)
  const destination = side === 'left' ? 'right' : 'left'
  const label = AI_PANEL_SIDE_LABELS[lang][destination]
  const move = async () => {
    setPending(true)
    try {
      applyAiPanelPrefs(await onMove(destination))
    } catch (error) {
      console.error('[ai-panel] Failed to save panel position:', error)
    } finally {
      setPending(false)
    }
  }
  return (
    <button
      type="button"
      className="ai-header-btn ai-panel-side-button"
      data-tip={label}
      aria-label={label}
      disabled={pending}
      onClick={() => void move()}
    >
      <svg
        width="16"
        height="16"
        viewBox="0 0 16 16"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.25"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden="true"
        style={destination === 'left' ? { transform: 'scaleX(-1)' } : undefined}
      >
        <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
        <path d="M11.5 2.5v11M4 8h5M7 5.8 9.2 8 7 10.2" />
      </svg>
    </button>
  )
}
