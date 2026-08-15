import { useEffect, useState } from 'react'
import type { ReactElement } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { MetadataInput } from '../shared/ipc'
import type { TFunc } from './i18n/locale'

interface RawInfo {
  Title?: string
  Author?: string
  Subject?: string
  Keywords?: string
  Creator?: string
  Producer?: string
  CreationDate?: string
  ModDate?: string
  PDFFormatVersion?: string
}

/** PDF date string D:YYYYMMDDHHmmSS… → locally readable */
function fmtPdfDate(raw?: string): string {
  const m = /^D:(\d{4})(\d{2})(\d{2})(\d{2})?(\d{2})?(\d{2})?/.exec(raw ?? '')
  if (!m) return raw ?? '—'
  const [, y, mo, d, h = '00', mi = '00'] = m
  return `${y}-${mo}-${d} ${h}:${mi}`
}

const fmtSize = (bytes: number): string =>
  bytes >= 1 << 20
    ? `${(bytes / (1 << 20)).toFixed(1)} MB`
    : `${Math.max(1, Math.round(bytes / 1024))} KB`

/** Document properties: first four fields are editable and become unsaved changes; the rest are read-only */
export function PropertiesDialog({
  doc,
  fileName,
  fileSize,
  pageCount,
  pending,
  readOnly,
  t,
  onCancel,
  onApply,
}: {
  doc: PDFDocumentProxy
  fileName: string
  fileSize: number
  pageCount: number
  /** Properties edited but not yet saved */
  pending: MetadataInput | null
  readOnly: boolean
  t: TFunc
  onCancel: () => void
  onApply: (meta: MetadataInput) => void
}): ReactElement {
  const [info, setInfo] = useState<RawInfo | null>(null)
  const [form, setForm] = useState<MetadataInput>({})

  useEffect(() => {
    let cancelled = false
    void doc.getMetadata().then(({ info: raw }) => {
      if (cancelled) return
      const i = (raw ?? {}) as RawInfo
      setInfo(i)
      setForm({
        title: pending?.title ?? i.Title ?? '',
        author: pending?.author ?? i.Author ?? '',
        subject: pending?.subject ?? i.Subject ?? '',
        keywords: pending?.keywords ?? i.Keywords ?? '',
      })
    })
    return () => {
      cancelled = true
    }
  }, [doc, pending])

  const edit = (key: keyof MetadataInput, label: string): ReactElement => (
    <label className="pdf-field">
      <span>{label}</span>
      <input
        className="pdf-modal-input"
        value={form[key] ?? ''}
        disabled={readOnly}
        onChange={(e) => setForm({ ...form, [key]: e.target.value })}
      />
    </label>
  )

  const row = (label: string, value: string): ReactElement => (
    <div className="pdf-prop-row">
      <span>{label}</span>
      <em data-tip={value}>{value || '—'}</em>
    </div>
  )

  return (
    <div className="pdf-modal-mask" onClick={onCancel}>
      <div className="pdf-modal pdf-modal-wide" onClick={(e) => e.stopPropagation()}>
        <div className="pdf-modal-title">{t('propsTitle')}</div>
        {edit('title', t('propTitle'))}
        {edit('author', t('propAuthor'))}
        {edit('subject', t('propSubject'))}
        {edit('keywords', t('propKeywords'))}
        <div className="pdf-prop-table">
          {row(t('propFileName'), fileName)}
          {row(t('propPages'), String(pageCount))}
          {row(t('propSize'), fmtSize(fileSize))}
          {row(t('propVersion'), info?.PDFFormatVersion ?? '—')}
          {row(t('propProducer'), info?.Producer ?? '')}
          {row(t('propCreated'), fmtPdfDate(info?.CreationDate))}
          {row(t('propModified'), fmtPdfDate(info?.ModDate))}
        </div>
        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn" onClick={onCancel}>
            {readOnly ? t('ok') : t('cancel')}
          </button>
          {!readOnly && (
            <button className="pdf-modal-btn primary" onClick={() => onApply(form)}>
              {t('ok')}
            </button>
          )}
        </div>
      </div>
    </div>
  )
}
