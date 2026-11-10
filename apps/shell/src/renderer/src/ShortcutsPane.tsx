import { useEffect, useMemo, useState, useSyncExternalStore } from 'react'
import { useI18n } from './locale'
import type { StringKey } from './locale'
import {
  IS_MAC,
  SHORTCUTS_CHANGED_EVENT,
  SHORTCUT_DEFS,
  chordFromEvent,
  findOwner,
  formatChord,
  formatChordParts,
  getShortcutMap,
  isReservedChord,
  isUsableChord,
  resetShortcutMap,
  setShortcutMap,
  setShortcutRecording,
} from './shortcuts'
import type { ShortcutGroup, ShortcutId } from './shortcuts'

// ── Settings › Shortcuts ────────────────────────────────────
// Each action lists its bindings as keycaps with a remove button; "+" records a
// new chord. Changes persist and take effect immediately (the global
// dispatcher re-reads the map on SHORTCUTS_CHANGED_EVENT).

export const SHORTCUT_LABEL_KEYS: Record<ShortcutId, StringKey> = {
  openConversationPicker: 'scaOpenConversationPicker',
  openFileSearch: 'scaOpenFileSearch',
  newConversation: 'scaNewConversation',
  goBack: 'scaGoBack',
  goForward: 'scaGoForward',
  newTab: 'scaNewTab',
  reopenClosedTab: 'scaReopenClosedTab',
  previousTab: 'scaPreviousTab',
  nextTab: 'scaNextTab',
  openSettings: 'scaOpenSettings',
  openRecycleBin: 'scaOpenRecycleBin',
}

const GROUP_KEYS: Record<ShortcutGroup, StringKey> = {
  navigation: 'scGroupNavigation',
  general: 'scGroupGeneral',
}

function subscribeMap(cb: () => void): () => void {
  window.addEventListener(SHORTCUTS_CHANGED_EVENT, cb)
  return () => window.removeEventListener(SHORTCUTS_CHANGED_EVENT, cb)
}

function Keycap({ chord }: { chord: string }) {
  const { mods, key } = formatChordParts(chord)
  return (
    <kbd className="sc-keycap">
      {mods}
      {mods && !IS_MAC ? ' ' : ''}
      {key}
    </kbd>
  )
}

export function ShortcutsPane() {
  const { t } = useI18n()
  const map = useSyncExternalStore(subscribeMap, getShortcutMap, getShortcutMap)
  const [filter, setFilter] = useState('')
  const [recording, setRecording] = useState<ShortcutId | null>(null)
  const [message, setMessage] = useState<{ id: ShortcutId; text: string } | null>(null)

  const labelOf = (id: ShortcutId): string =>
    id === 'openRecycleBin' && IS_MAC ? t('scaOpenTrash') : t(SHORTCUT_LABEL_KEYS[id])

  // record the next chord while a row is armed; the global dispatcher stays quiet meanwhile
  useEffect(() => {
    if (!recording) return
    setShortcutRecording(true)
    const onKey = (e: KeyboardEvent): void => {
      e.preventDefault()
      e.stopPropagation()
      if (e.key === 'Escape') {
        setRecording(null)
        setMessage(null)
        return
      }
      const chord = chordFromEvent(e)
      if (!chord) return // bare modifier: keep waiting
      if (!isUsableChord(chord)) {
        setMessage({ id: recording, text: t('scErrModifier') })
        return
      }
      if (isReservedChord(chord)) {
        setMessage({ id: recording, text: t('scErrReserved') })
        return
      }
      const owner = findOwner(map, chord)
      if (owner) {
        setMessage({ id: recording, text: t('scErrUsed', { name: labelOf(owner) }) })
        return
      }
      setShortcutMap({ ...map, [recording]: [...map[recording], chord] })
      setRecording(null)
      setMessage(null)
    }
    window.addEventListener('keydown', onKey, true)
    return () => {
      window.removeEventListener('keydown', onKey, true)
      setShortcutRecording(false)
    }
    // labelOf closes over t; re-arming on language change is harmless
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [recording, map, t])

  const q = filter.trim().toLowerCase()
  const groups = useMemo(() => {
    const out: { group: ShortcutGroup; ids: ShortcutId[] }[] = []
    for (const def of SHORTCUT_DEFS) {
      const hay = `${labelOf(def.id)} ${map[def.id].map((c) => formatChord(c)).join(' ')}`.toLowerCase()
      if (q && !hay.includes(q)) continue
      let bucket = out.find((g) => g.group === def.group)
      if (!bucket) {
        bucket = { group: def.group, ids: [] }
        out.push(bucket)
      }
      bucket.ids.push(def.id)
    }
    return out
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [q, map, t])

  const fileRows: readonly (readonly [StringKey, readonly string[]])[] = [
    ['scOpen', ['Enter']],
    ['scNavigate', ['↑', '↓']],
    ['scRename', ['F2']],
    ['scDelete', ['Delete']],
    ['scSelectAll', [IS_MAC ? '⌘' : 'Ctrl', 'A']],
    ['searchFiles', [IS_MAC ? '⌘' : 'Ctrl', 'F']],
    ['scRange', ['Shift', 'Click']],
    ['scMenu', ['Right-click']],
    ['scEscape', ['Esc']],
  ]
  const fileRowsShown = fileRows.filter(
    ([k, keys]) => !q || `${t(k)} ${keys.join(' ')}`.toLowerCase().includes(q),
  )

  const addButton = (id: ShortcutId) => (
    <button
      type="button"
      className={`sc-add${map[id].length === 0 ? ' empty' : ''}`}
      title={t('scAdd')}
      aria-label={`${t('scAdd')}: ${labelOf(id)}`}
      onClick={() => {
        setMessage(null)
        setRecording(id)
      }}
    >
      {map[id].length === 0 ? (
        t('scAdd')
      ) : (
        <svg width="12" height="12" viewBox="0 0 14 14" aria-hidden="true">
          <path d="M7 2v10M2 7h10" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
        </svg>
      )}
    </button>
  )

  return (
    <>
      <h3 className="set-pane-title ap-title">{t('setSecShortcuts')}</h3>
      <p className="ap-sub">{t('setShortcutsSub')}</p>

      <label className="sc-search">
        <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
          <circle cx="7" cy="7" r="4.6" stroke="currentColor" strokeWidth="1.3" />
          <path d="M10.5 10.5L14 14" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
        </svg>
        <input
          value={filter}
          placeholder={t('scSearch')}
          aria-label={t('scSearch')}
          spellCheck={false}
          onChange={(e) => setFilter(e.target.value)}
        />
      </label>

      {groups.map(({ group, ids }) => (
        <section className="ap-group" key={group}>
          <h4 className="ap-group-title">{t(GROUP_KEYS[group])}</h4>
          <div className="ap-card">
            {ids.map((id) => (
              <div className="sc-row" key={id}>
                <div className="ap-row-label">{labelOf(id)}</div>
                <div className="sc-bindings">
                  {map[id].map((chord, i) => (
                    <div className="sc-binding" key={chord}>
                      {i === map[id].length - 1 && recording !== id && addButton(id)}
                      <Keycap chord={chord} />
                      <button
                        type="button"
                        className="sc-icon-btn"
                        title={t('scRemove')}
                        aria-label={`${t('scRemove')}: ${formatChord(chord)}`}
                        onClick={() =>
                          setShortcutMap({ ...map, [id]: map[id].filter((c) => c !== chord) })
                        }
                      >
                        <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                          <path
                            d="M3 4.5h10M6.5 4.5V3h3v1.5M4.5 4.5l.6 8h5.8l.6-8"
                            stroke="currentColor"
                            strokeWidth="1.3"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          />
                        </svg>
                      </button>
                    </div>
                  ))}
                  {recording === id && (
                    <div className="sc-binding">
                      <kbd className="sc-keycap recording">{t('scPress')}</kbd>
                      <button
                        type="button"
                        className="sc-icon-btn"
                        title={t('scCancelHint')}
                        aria-label={t('scCancelHint')}
                        onClick={() => {
                          setRecording(null)
                          setMessage(null)
                        }}
                      >
                        <svg width="12" height="12" viewBox="0 0 14 14" aria-hidden="true">
                          <path
                            d="M2 2l10 10M12 2L2 12"
                            stroke="currentColor"
                            strokeWidth="1.5"
                            strokeLinecap="round"
                          />
                        </svg>
                      </button>
                    </div>
                  )}
                  {map[id].length === 0 && recording !== id && addButton(id)}
                  {message?.id === id && (
                    <div className="sc-msg" role="alert">
                      {message.text}
                    </div>
                  )}
                </div>
              </div>
            ))}
          </div>
        </section>
      ))}

      {fileRowsShown.length > 0 && (
        <section className="ap-group">
          <h4 className="ap-group-title">{t('scGroupFiles')}</h4>
          <div className="ap-card">
            {fileRowsShown.map(([labelKey, keys]) => (
              <div className="sc-row" key={labelKey}>
                <div className="ap-row-label">{t(labelKey)}</div>
                <div className="sc-bindings">
                  <div className="sc-binding">
                    <kbd className="sc-keycap fixed">{keys.join(!IS_MAC && keys.length > 1 && keys[0] !== '↑' ? '+ ' : ' ')}</kbd>
                  </div>
                </div>
              </div>
            ))}
          </div>
          <p className="ap-footnote sc-note">{t('shortcutsNote')}</p>
        </section>
      )}

      {groups.length === 0 && fileRowsShown.length === 0 && (
        <div className="sc-none">{t('scNone')}</div>
      )}

      <div className="ap-footer">
        <span />
        <button type="button" className="set-btn" onClick={resetShortcutMap}>
          {t('scReset')}
        </button>
      </div>
    </>
  )
}
