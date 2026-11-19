/**
 * Word's keyboard-shortcut reference (File ▸ Options ▸ Customize Keyboard in
 * Word; Help ▸ Keyboard Shortcuts here). Read-only: it renders the registry in
 * shortcuts.ts, so a binding added there shows up without touching this file,
 * with the chords the user assigned in Customize Keyboard laid over it.
 */
import { useEffect, useState } from 'react'
import { bindingId } from '../../shared/keybindings'
import { useI18n } from '../i18n/locale'
import { useKeyOverrides } from '../key-layer'
import {
  formatChord,
  listKeyCommands,
  valueBindingLabel,
  type MenuCatalogEntry,
} from '../keybindings'
import { SHORTCUT_GROUPS, SHORTCUTS, shortcutKeys } from '../shortcuts'
import { useModalKeys } from './modal-keys'

export function ShortcutsDialog({ onClose }: { onClose: () => void }) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [query, setQuery] = useState('')

  const overrides = useKeyOverrides()
  const [catalog, setCatalog] = useState<MenuCatalogEntry[]>([])
  const customized = Object.keys(overrides).length > 0
  useEffect(() => {
    if (!customized) return
    let live = true
    void window.desktop
      ?.keyCatalog?.()
      .then((entries) => live && setCatalog(entries))
      .catch(() => undefined)
    return () => {
      live = false
    }
  }, [customized])

  const needle = query.trim().toLowerCase()
  const assigned = (id: string) => overrides[id].map((chord) => formatChord(chord)).join(' / ')
  const listed = new Set<string>()
  const builtIn = SHORTCUTS.map((def) => {
    const binding = def.cmd ? bindingId(def.cmd, def.payload) : null
    const custom = binding !== null && Object.hasOwn(overrides, binding)
    if (binding) listed.add(binding)
    return {
      id: def.id,
      group: def.group as string,
      label: t(def.labelKey) + (def.labelSuffix ?? ''),
      keys: custom ? assigned(binding) : shortcutKeys(def),
    }
  })
  const extra = customized
    ? [
        ...listKeyCommands(catalog, t).filter(
          (command) => Object.hasOwn(overrides, command.id) && !listed.has(command.id),
        ),
        // a font, a style, a size or a colour given a key
        ...Object.keys(overrides).flatMap((id) => {
          const label = listed.has(id) ? null : valueBindingLabel(id, t)
          return label ? [{ id, label }] : []
        }),
      ].map((command) => ({
        id: `custom:${command.id}`,
        group: 'custom',
        label: command.label,
        keys: assigned(command.id),
      }))
    : []
  const rows = [...builtIn, ...extra].filter(
    (row) =>
      row.keys &&
      (!needle ||
        row.label.toLowerCase().includes(needle) ||
        row.keys.toLowerCase().includes(needle)),
  )

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal modal-shortcuts" role="dialog" aria-label={t('appScTitle')}>
        <h2>{t('appScTitle')}</h2>
        <input
          type="search"
          className="sc-filter"
          placeholder={t('appScFilter')}
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
        <div className="sc-list">
          {[...SHORTCUT_GROUPS, { id: 'custom', labelKey: 'appScGroupCustom' } as const].map(
            (group) => {
              const groupRows = rows.filter((row) => row.group === group.id)
              if (groupRows.length === 0) return null
              return (
                <section key={group.id} className="sc-group">
                  <h3>{t(group.labelKey)}</h3>
                  {groupRows.map((row) => (
                    <div key={row.id} className="sc-row">
                      <span className="sc-label">{row.label}</span>
                      <kbd className="sc-keys">{row.keys}</kbd>
                    </div>
                  ))}
                </section>
              )
            },
          )}
          {rows.length === 0 && <p className="sc-empty">{t('appScNone')}</p>}
        </div>
        <div className="modal-actions">
          <button className="btn-primary" onClick={onClose}>
            {t('appClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
