import type { TFunc } from './i18n/locale'

/**
 * View → Attachments: the embedded files the document carries, plus what this
 * session adds or lifts — those flow into the next save like every other edit,
 * so rows show a pending marker until the save lands (the post-save reload
 * refreshes the list from the new file).
 */
export interface AttachmentRow {
  name: string
  bytes: number
  /** in the file on disk; rows without it are pending adds */
  saved: boolean
  /** pending removal from the file on disk */
  removing?: boolean
}

export function humanBytes(n: number): string {
  if (n < 1024) return `${n} B`
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`
  return `${(n / (1024 * 1024)).toFixed(1)} MB`
}

export function AttachmentsPanel({
  t,
  rows,
  onAdd,
  onExport,
  onToggleRemove,
}: {
  t: TFunc
  rows: AttachmentRow[]
  onAdd: () => void
  onExport: (row: AttachmentRow) => void
  onToggleRemove: (row: AttachmentRow) => void
}): React.JSX.Element {
  return (
    <div className="pdf-attachments">
      <button type="button" className="pdf-modal-btn primary pdf-attachments-add" onClick={onAdd}>
        {t('attachmentsAdd')}
      </button>
      {rows.length === 0 && <div className="pdf-modal-hint">{t('attachmentsEmpty')}</div>}
      <ul className="pdf-attachments-list">
        {rows.map((row) => (
          <li
            key={row.name}
            className={row.removing ? 'pdf-attachment removing' : 'pdf-attachment'}
          >
            <div className="pdf-attachment-name" title={row.name}>
              {row.name}
              {!row.saved && (
                <span className="pdf-attachment-tag">{t('attachmentsPendingAdd')}</span>
              )}
              {row.removing && (
                <span className="pdf-attachment-tag">{t('attachmentsPendingRemove')}</span>
              )}
            </div>
            <div className="pdf-attachment-size">{humanBytes(row.bytes)}</div>
            {!row.removing && (
              <button
                type="button"
                className="rb-icon"
                aria-label={t('attachmentsExport')}
                data-tip={t('attachmentsExport')}
                disabled={!row.saved}
                onClick={() => onExport(row)}
              >
                ↓
              </button>
            )}
            <button
              type="button"
              className="rb-icon"
              aria-label={row.removing ? t('attachmentsKeep') : t('attachmentsRemove')}
              data-tip={row.removing ? t('attachmentsKeep') : t('attachmentsRemove')}
              onClick={() => onToggleRemove(row)}
            >
              {row.removing ? '↺' : '×'}
            </button>
          </li>
        ))}
      </ul>
    </div>
  )
}
