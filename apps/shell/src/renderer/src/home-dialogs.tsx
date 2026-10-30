import { useEffect, useState } from 'react'
import type { MoveConflictPolicy } from '../../shared/home-api'
import { useI18n } from './locale'
import { FolderIcon } from './home-shared'

// ── Same-name conflict prompt ───────────────────────────

export interface ConflictPromptProps {
  names: string[]
  onChoose: (policy: MoveConflictPolicy) => void
}

export function ConflictPrompt({ names, onChoose }: ConflictPromptProps) {
  const { t } = useI18n()
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onChoose('skip')
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onChoose])
  return (
    <div className="modal-overlay" onClick={() => onChoose('skip')}>
      <div
        className="modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('conflictTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <h3>{t('conflictTitle')}</h3>
        <p>
          {names.length === 1
            ? t('conflictBodyOne', { name: names[0] })
            : t('conflictBodyMany', { n: names.length })}
        </p>
        {names.length > 1 && (
          <ul className="modal-file-list">
            {names.slice(0, 6).map((n) => (
              <li key={n}>{n}</li>
            ))}
            {names.length > 6 && <li>{t('deleteMoreCount', { n: names.length })}</li>}
          </ul>
        )}
        <div className="modal-buttons">
          <button className="btn btn-secondary" autoFocus onClick={() => onChoose('skip')}>
            {t('cancel')}
          </button>
          <button className="btn btn-secondary" onClick={() => onChoose('keepBoth')}>
            {t('conflictKeepBoth')}
          </button>
          <button className="btn btn-danger" onClick={() => onChoose('replace')}>
            {t('conflictReplace')}
          </button>
        </div>
      </div>
    </div>
  )
}

// ── Starred-group picker ─────────────────────────────────

export interface GroupPickerProps {
  /** existing group names (a group exists while it has at least one file) */
  groups: readonly string[]
  count: number
  onCancel: () => void
  /** put the files into an existing group */
  onPick: (group: string) => void
  /** create a group named `name` and put the files into it */
  onCreate: (name: string) => void
  /** take the files out of their group (back to the unfiltered list) */
  onRemove: () => void
}

/** Modal moving starred files into a named group. Deliberately minimal: no
 *  group rename/delete — a group is just a label its entries carry, and it
 *  vanishes when its last member leaves (via this dialog's remove action,
 *  an unstar, or a deleted file). */
export function GroupPicker({
  groups,
  count,
  onCancel,
  onPick,
  onCreate,
  onRemove,
}: GroupPickerProps) {
  const { t } = useI18n()
  const [name, setName] = useState('')

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onCancel()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onCancel])

  const create = () => {
    const trimmed = name.trim()
    if (!trimmed) return
    setName('')
    onCreate(trimmed)
  }

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div
        className="modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('moveToGroupTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <h3>{t('moveToGroupTitle')}</h3>
        <p>{t('selectedCount', { n: count })}</p>
        {groups.length > 0 && (
          <ul className="group-picker-list">
            {groups.map((group) => (
              <li key={group}>
                <button
                  className="group-picker-row"
                  onClick={() => {
                    setName('')
                    onPick(group)
                  }}
                >
                  <FolderIcon size={16} />
                  <span className="group-picker-name">{group}</span>
                </button>
              </li>
            ))}
          </ul>
        )}
        <div className="group-picker-new">
          <input
            className="folder-rename-input"
            placeholder={t('groupNamePlaceholder')}
            value={name}
            onChange={(e) => setName(e.target.value)}
            onKeyDown={(e) => {
              e.stopPropagation()
              if (e.nativeEvent.isComposing) return
              if (e.key === 'Enter') create()
            }}
          />
          <button className="btn btn-secondary" disabled={!name.trim()} onClick={create}>
            {t('newGroup')}
          </button>
        </div>
        <div className="modal-buttons">
          <button className="btn btn-secondary" onClick={onRemove}>
            {t('removeFromGroup')}
          </button>
          <button className="btn btn-secondary" autoFocus onClick={onCancel}>
            {t('cancel')}
          </button>
        </div>
      </div>
    </div>
  )
}
