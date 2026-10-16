import { useEffect, useMemo, useRef, useState } from 'react'
import { matchesCommand, type FeatureRecord } from '@genoffice/agent-core'
import { canRunAction, runSlotAction, type SlotActions } from '@genoffice/ui'
import { useI18n } from './locale'
import './command-palette.css'

/**
 * The command palette: the commands the user's features registered (K6).
 *
 * A command runs the same declared actions a drawn control may run, so the palette
 * adds an entry point and nothing else. It is deliberately plain: type, arrow,
 * Enter. Every command that cannot be served by this host is listed as unavailable
 * rather than left to fail silently.
 */

export function CommandPalette({
  commands,
  actions,
  onClose,
}: {
  commands: FeatureRecord[]
  actions: SlotActions
  onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [query, setQuery] = useState('')
  const [active, setActive] = useState(0)
  const listRef = useRef<HTMLUListElement>(null)

  const shown = useMemo(
    () => commands.filter((feature) => feature.command && matchesCommand(feature.command, query)),
    [commands, query],
  )

  useEffect(() => {
    setActive(0)
  }, [query])

  const invoke = (feature: FeatureRecord): void => {
    if (!feature.command) return
    if (!canRunAction(feature.command.action, actions)) return
    onClose()
    runSlotAction(feature.command.action, actions)
  }

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        event.preventDefault()
        onClose()
        return
      }
      if (event.key === 'ArrowDown') {
        event.preventDefault()
        setActive((index) => Math.min(index + 1, Math.max(shown.length - 1, 0)))
        return
      }
      if (event.key === 'ArrowUp') {
        event.preventDefault()
        setActive((index) => Math.max(index - 1, 0))
        return
      }
      if (event.key === 'Enter') {
        const entry = shown[active]
        if (entry) {
          event.preventDefault()
          invoke(entry)
        }
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  })

  return (
    <div className="palette-backdrop" onMouseDown={onClose}>
      <div
        className="palette"
        role="dialog"
        aria-label={t('paletteTitle')}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <input
          className="palette-input"
          autoFocus
          value={query}
          placeholder={t('palettePlaceholder')}
          aria-label={t('palettePlaceholder')}
          onChange={(event) => setQuery(event.target.value)}
        />
        {shown.length === 0 ? (
          <p className="palette-empty">{t('paletteEmpty')}</p>
        ) : (
          <ul className="palette-list" ref={listRef}>
            {shown.map((feature, index) => {
              const command = feature.command!
              const available = canRunAction(command.action, actions)
              return (
                <li key={feature.id}>
                  <button
                    className={'palette-item' + (index === active ? ' palette-item--active' : '')}
                    disabled={!available}
                    onMouseEnter={() => setActive(index)}
                    onClick={() => invoke(feature)}
                  >
                    <span className="palette-label">{command.label}</span>
                    {command.description ? (
                      <span className="palette-desc">{command.description}</span>
                    ) : null}
                  </button>
                </li>
              )
            })}
          </ul>
        )}
      </div>
    </div>
  )
}
