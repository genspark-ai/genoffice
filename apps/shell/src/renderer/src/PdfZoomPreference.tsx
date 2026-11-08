import { useEffect, useState } from 'react'
import { Dropdown } from '@genoffice/ui'
import { DEFAULT_PDF_ZOOM } from '../../../../pdf/src/shared/default-zoom'
import type { PdfDefaultZoom } from '../../../../pdf/src/shared/default-zoom'
import { useI18n } from './locale'

export function PdfZoomPreference() {
  const { t } = useI18n()
  const [value, setValue] = useState<PdfDefaultZoom>(DEFAULT_PDF_ZOOM)
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [failed, setFailed] = useState(false)

  useEffect(() => {
    let alive = true
    void window.aiOffice
      .getPdfDefaultZoom?.()
      .then((saved) => {
        if (alive) setValue(saved)
      })
      .catch(() => {
        if (alive) setFailed(true)
      })
      .finally(() => {
        if (alive) setLoading(false)
      })
    return () => {
      alive = false
    }
  }, [])

  const pick = (next: PdfDefaultZoom) => {
    setSaving(true)
    setFailed(false)
    void window.aiOffice
      .setPdfDefaultZoom(next)
      .then(setValue)
      .catch(() => setFailed(true))
      .finally(() => setSaving(false))
  }

  return (
    <div className="set-field">
      <div className="set-field-text">
        <div className="set-field-stack">
          <div className="set-field-label">{t('setPdfZoom')}</div>
          <div className="set-field-desc">{t('setPdfZoomDesc')}</div>
          {failed && (
            <div className="set-field-desc" role="alert">
              {t('setPdfZoomFailed')}
            </div>
          )}
        </div>
      </div>
      <Dropdown
        className="set-dd"
        ariaLabel={t('setPdfZoom')}
        value={value}
        disabled={loading || saving}
        options={[
          { value: 'width', label: t('setPdfZoomWidth') },
          { value: 'page', label: t('setPdfZoomPage') },
          { value: 'last-used', label: t('setPdfZoomLast') },
        ]}
        onPick={pick}
      />
    </div>
  )
}
