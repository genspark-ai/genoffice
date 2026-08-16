import type { NoteInfo } from '@genoffice/docx-engine'
import { toRoman } from '../note-format'
import { useI18n } from '../i18n/locale'

function NoteRow({
  marker,
  note,
  editTip,
  deleteTip,
  onEdit,
  onDelete,
}: {
  marker: string
  note: NoteInfo
  editTip: string
  deleteTip: string
  onEdit: (id: string) => void
  onDelete: (id: string) => void
}) {
  return (
    <div className="page-note">
      <sup>{marker}</sup>
      <span className="page-note-text">{note.text}</span>
      <button
        className="page-note-btn"
        data-tip={editTip}
        aria-label={editTip}
        onClick={() => onEdit(note.id)}
      >
        ✎
      </button>
      <button
        className="page-note-btn"
        data-tip={deleteTip}
        aria-label={deleteTip}
        onClick={() => onDelete(note.id)}
      >
        ×
      </button>
    </div>
  )
}

/** End-of-document footnote list; entries already shown in canvas page gaps are skipped (numbering keeps full-list order) */
export function PageFootnotes({
  notes,
  skipIds,
  onEdit,
  onDelete,
}: {
  notes: NoteInfo[]
  skipIds: ReadonlySet<string>
  onEdit: (id: string) => void
  onDelete: (id: string) => void
}) {
  const { t } = useI18n()
  const shown = notes.filter((n) => !skipIds.has(n.id))
  if (shown.length === 0) return null
  return (
    <div className="page-notes">
      {shown.map((n) => (
        <NoteRow
          key={`f${n.id}`}
          marker={String(notes.indexOf(n) + 1)}
          note={n}
          editTip={t('appEditFootnote')}
          deleteTip={t('appDeleteFootnote')}
          onEdit={onEdit}
          onDelete={onDelete}
        />
      ))}
    </div>
  )
}

/**
 * Endnote area, roman-numbered, never mixed into the footnote block. Word shows
 * no heading and places it right after the last body line, so when the measured
 * flow-end anchor is available the area is absolutely positioned there instead
 * of stacking after the (page-tall) editor.
 */
export function PageEndnotes({
  notes,
  top,
  onEdit,
  onDelete,
}: {
  notes: NoteInfo[]
  /** flow-end anchor (layout px from the page-wrap top); null = not measured yet */
  top: number | null
  onEdit: (id: string) => void
  onDelete: (id: string) => void
}) {
  const { t } = useI18n()
  if (notes.length === 0) return null
  return (
    <div
      className={`page-notes page-endnotes${top != null ? ' page-endnotes-anchored' : ''}`}
      style={top != null ? { top } : undefined}
    >
      {notes.map((n, i) => (
        <NoteRow
          key={`e${n.id}`}
          marker={toRoman(i + 1)}
          note={n}
          editTip={t('appEditEndnote')}
          deleteTip={t('appDeleteEndnote')}
          onEdit={onEdit}
          onDelete={onDelete}
        />
      ))}
    </div>
  )
}
