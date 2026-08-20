import { useEffect, useState } from 'react'
import appIcon from './assets/app-icon.png'
import { useI18n } from './locale'

export interface AboutDialogProps {
  onClose: () => void
}

/**
 * About dialog: product identity (KĀRYA — Intelligent Workspace), the
 * installed version, and the upstream attribution (this project is a fork of
 * GenOffice, so the derivation is stated explicitly rather than hidden).
 */
export function AboutDialog({ onClose }: AboutDialogProps) {
  const { t } = useI18n()
  const [appVersion, setAppVersion] = useState('')

  useEffect(() => {
    void window.aiOffice.getAppVersion?.().then((v) => {
      if (v) setAppVersion(v)
    })
  }, [])

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        className="modal about-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('aboutTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <img className="about-icon" src={appIcon} alt="" width={72} height={72} />
        <h3>{t('aboutTitle')}</h3>
        <p className="about-subtitle">{t('aboutSubtitle')}</p>
        <p className="about-version">
          {t('versionLabel')} {appVersion ? `v${appVersion}` : '—'}
        </p>
        <p className="about-attribution">{t('aboutAttribution')}</p>
        <p className="about-license">{t('aboutLicense')}</p>
        <div className="modal-buttons">
          <button className="btn btn-secondary" autoFocus onClick={onClose}>
            {t('aboutClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
