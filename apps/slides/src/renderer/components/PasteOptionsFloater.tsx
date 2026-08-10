/**
 * PowerPoint-style paste-options floater: a small clipboard button sitting on
 * the just-pasted slide's thumbnail. Its menu re-runs the paste with another
 * mode (keep source formatting / use destination theme / paste as picture).
 */
import React, { useEffect, useRef, useState } from 'react'
import type { PasteSlideMode } from '../../shared/ipc'
import { ContextMenu, type CtxItem } from './ContextMenu'
import { IconPaste } from './icons'
import { t } from '../i18n/locale'

interface Props {
  mode: PasteSlideMode
  onSelect: (mode: PasteSlideMode) => void
  onDismiss: () => void
}

export function PasteOptionsFloater({ mode, onSelect, onDismiss }: Props) {
  const btnRef = useRef<HTMLButtonElement>(null)
  const [menuPos, setMenuPos] = useState<{ x: number; y: number } | null>(null)

  // Escape with the menu closed dismisses the whole floater (the open menu's own
  // Escape handler only closes the menu)
  useEffect(() => {
    if (menuPos) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onDismiss()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [menuPos, onDismiss])

  const mark = (m: PasteSlideMode) => (m === mode ? '✓ ' : '\u2003 ')
  const items: CtxItem[] = [
    {
      label: `${mark('source')}${t('appPasteOptSource')}`,
      onClick: () => onSelect('source'),
    },
    {
      label: `${mark('theme')}${t('appPasteOptTheme')}`,
      onClick: () => onSelect('theme'),
    },
    {
      label: `${mark('picture')}${t('appCtxPasteSlideAsPicture')}`,
      onClick: () => onSelect('picture'),
    },
  ]

  return (
    <>
      <button
        ref={btnRef}
        className="paste-floater-btn"
        data-tip={t('appPasteOptions')}
        aria-label={t('appPasteOptions')}
        // The thumbnail behind is clickable and draggable: don't select the page or start a drag
        draggable={false}
        onMouseDown={(e) => e.stopPropagation()}
        onClick={(e) => {
          e.stopPropagation()
          const r = btnRef.current?.getBoundingClientRect()
          if (r) setMenuPos({ x: r.left, y: r.bottom + 4 })
        }}
        onContextMenu={(e) => {
          e.preventDefault()
          e.stopPropagation()
        }}
      >
        <IconPaste size={14} />
        <span className="paste-floater-caret">▾</span>
      </button>
      {menuPos && (
        <ContextMenu x={menuPos.x} y={menuPos.y} items={items} onClose={() => setMenuPos(null)} />
      )}
    </>
  )
}
