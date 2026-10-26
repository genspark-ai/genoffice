import { useState } from 'react'
import type { ReactElement } from 'react'
import type { TFunc } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

/** Pending link box drawn on a page (PDF user space), waiting for its target */
export interface LinkDraft {
  pageIndex: number
  rect: [number, number, number, number]
}

/**
 * Target picker for a drawn link box: exactly one of a web address or a page in
 * this document (Acrobat's Edit PDF > Link offers the same two). An empty field
 * means "not this one"; OK stays disabled until exactly one field is valid.
 */
export function LinkDialog({
  t,
  pageCount,
  onCancel,
  onApply,
}: {
  t: TFunc
  pageCount: number
  onCancel: () => void
  onApply: (target: { url: string } | { destPage: number }) => void
}): ReactElement {
  const [url, setUrl] = useState('')
  const [page, setPage] = useState('')
  const dialogRef = useModalDialog(onCancel)

  const pageNum = Number(page)
  const urlValid = /^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(url.trim()) || url.trim().startsWith('www.')
  const pageValid =
    page.trim() !== '' && Number.isInteger(pageNum) && pageNum >= 1 && pageNum <= pageCount
  // exactly one target: the field the user cleared last loses
  const target: { url: string } | { destPage: number } | null =
    urlValid && !pageValid
      ? { url: url.trim().startsWith('www.') ? `https://${url.trim()}` : url.trim() }
      : pageValid && !urlValid
        ? { destPage: pageNum }
        : null

  return (
    <div className="pdf-modal-mask" onClick={onCancel}>
      <div
        ref={dialogRef}
        className="pdf-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('linkDialogTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="pdf-modal-title">{t('linkDialogTitle')}</div>
        <label className="pdf-field">
          <span>{t('linkUrlLabel')}</span>
          <input
            className="pdf-modal-input"
            autoFocus
            placeholder="https://…"
            value={url}
            onChange={(e) => {
              setUrl(e.target.value)
              if (e.target.value.trim() !== '') setPage('')
            }}
          />
        </label>
        <label className="pdf-field">
          <span>{t('linkPageLabel', { n: pageCount })}</span>
          <input
            className="pdf-modal-input"
            inputMode="numeric"
            placeholder="1–"
            value={page}
            onChange={(e) => {
              setPage(e.target.value.replace(/[^0-9]/g, ''))
              if (e.target.value !== '') setUrl('')
            }}
          />
        </label>
        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn" onClick={onCancel}>
            {t('cancel')}
          </button>
          <button
            className="pdf-modal-btn primary"
            disabled={!target}
            onClick={() => target && onApply(target)}
          >
            {t('ok')}
          </button>
        </div>
      </div>
    </div>
  )
}
