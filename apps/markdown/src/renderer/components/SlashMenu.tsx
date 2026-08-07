import { forwardRef, useEffect, useImperativeHandle, useState } from 'react'
import type { SlashMenuState } from '../editor/slashCommand'
import { useI18n } from '../i18n/locale'

export interface SlashMenuHandle {
  handleKey(event: KeyboardEvent): boolean
}

interface Props {
  state: SlashMenuState | null
  onDismiss: () => void
}

export const SlashMenu = forwardRef<SlashMenuHandle, Props>(function SlashMenu(
  { state, onDismiss },
  ref,
) {
  const { t } = useI18n()
  const [index, setIndex] = useState(0)

  const count = state?.items.length ?? 0

  useEffect(() => {
    setIndex(0)
  }, [state?.items])

  useImperativeHandle(
    ref,
    () => ({
      handleKey(event: KeyboardEvent): boolean {
        if (!state) return false
        if (event.key === 'ArrowDown') {
          setIndex((i) => (count === 0 ? 0 : (i + 1) % count))
          return true
        }
        if (event.key === 'ArrowUp') {
          setIndex((i) => (count === 0 ? 0 : (i - 1 + count) % count))
          return true
        }
        if (event.key === 'Enter') {
          const item = state.items[index]
          if (item) state.command(item)
          return true
        }
        if (event.key === 'Escape') {
          onDismiss()
          return true
        }
        return false
      },
    }),
    [state, index, count, onDismiss],
  )

  if (!state || !state.clientRect) return null
  const rect = state.clientRect

  return (
    <div
      className="slash-menu"
      style={{ position: 'fixed', top: rect.bottom + 4, left: rect.left }}
      onMouseDown={(e) => e.preventDefault()}
    >
      {state.items.length === 0 ? (
        <div className="slash-empty">{t('slashNoResults')}</div>
      ) : (
        state.items.map((item, i) => (
          <button
            key={item.id}
            type="button"
            className={`slash-item${i === index ? ' active' : ''}`}
            onMouseEnter={() => setIndex(i)}
            onClick={() => state.command(item)}
          >
            {t(item.labelKey)}
          </button>
        ))
      )}
    </div>
  )
})
