import React, { useEffect, useState } from 'react'
import type { ImageHostConfig, ImageHostKind } from '../../shared/ipc'
import { t, type StringKey } from '../i18n/locale'
import { setImageHostConfigCache } from '../imageHostCache'

/**
 * Configure the bring-your-own image host for pasted/dropped pictures
 * (genoffice#388): pick a kind, fill its credentials, and pastes upload there
 * first with the local assets/ copy as the fallback. Keys stay in the user's
 * own settings file, like every other credential in the app.
 */
export function ImageHostDialog({
  onClose,
  onSaved,
}: {
  onClose: () => void
  onSaved: (config: ImageHostConfig | null) => void
}): React.JSX.Element {
  const [loaded, setLoaded] = useState(false)
  const [kind, setKind] = useState<ImageHostKind>('s3')
  const [fields, setFields] = useState<Record<string, string>>({})
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState('')

  useEffect(() => {
    let alive = true
    void window.markdownApi
      .getImageHost()
      .then((config) => {
        if (!alive) return
        if (config) {
          setKind(config.kind)
          setFields(
            Object.fromEntries(
              Object.entries(config).filter(([, v]) => typeof v === 'string'),
            ) as Record<string, string>,
          )
        }
        setLoaded(true)
      })
      .catch(() => {
        if (alive) setLoaded(true)
      })
    return () => {
      alive = false
    }
  }, [])

  const field = (key: string): string => fields[key] ?? ''
  const setField = (key: string) => (e: React.ChangeEvent<HTMLInputElement>) =>
    setFields((prev) => ({ ...prev, [key]: e.target.value }))

  /** build the config from the form; '' empties stay undefined so validation sees them */
  const buildConfig = (): ImageHostConfig | null => {
    const value = (key: string) => field(key).trim()
    if (kind === 's3')
      return {
        kind,
        endpoint: value('endpoint') || undefined,
        region: value('region') || undefined,
        bucket: value('bucket') || undefined,
        accessKeyId: value('accessKeyId') || undefined,
        secretAccessKey: value('secretAccessKey') || undefined,
        prefix: value('prefix') || undefined,
        publicBase: value('publicBase') || undefined,
      }
    if (kind === 'smms') return { kind, token: value('token') || undefined }
    return {
      kind,
      token: value('token') || undefined,
      owner: value('owner') || undefined,
      repo: value('repo') || undefined,
      branch: value('branch') || undefined,
      dir: value('dir') || undefined,
      urlPrefix: value('urlPrefix') || undefined,
    }
  }

  const save = async (): Promise<void> => {
    setSaving(true)
    setError('')
    try {
      const saved = await window.markdownApi.setImageHost(buildConfig())
      if (!saved) {
        setError(t('imageHostInvalid'))
        return
      }
      setImageHostConfigCache(saved)
      onSaved(saved)
      onClose()
    } finally {
      setSaving(false)
    }
  }

  const textRow = (key: string, labelKey: StringKey, opts?: { placeholder?: string }) => (
    <label className="image-host-row" key={key}>
      <span className="image-host-label">{t(labelKey)}</span>
      <input
        type="text"
        className="image-host-input"
        value={field(key)}
        onChange={setField(key)}
        placeholder={opts?.placeholder}
        autoComplete="off"
        spellCheck={false}
      />
    </label>
  )

  return (
    <div className="image-host-backdrop" onClick={onClose}>
      <div
        className="image-host-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('imageHostTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <h3 className="image-host-title">{t('imageHostTitle')}</h3>
        <p className="image-host-sub">{t('imageHostSub')}</p>
        {!loaded ? null : (
          <>
            <label className="image-host-row">
              <span className="image-host-label">{t('imageHostKind')}</span>
              <select
                className="image-host-input"
                value={kind}
                onChange={(e) => setKind(e.target.value as ImageHostKind)}
              >
                <option value="s3">{t('imageHostKindS3')}</option>
                <option value="smms">SM.MS</option>
                <option value="github">GitHub</option>
              </select>
            </label>
            {kind === 's3' && (
              <>
                {textRow('endpoint', 'imageHostEndpoint')}
                {textRow('region', 'imageHostRegion', { placeholder: 'auto' })}
                {textRow('bucket', 'imageHostBucket')}
                {textRow('accessKeyId', 'imageHostAccessKeyId')}
                {textRow('secretAccessKey', 'imageHostSecretKey')}
                {textRow('prefix', 'imageHostPrefix', { placeholder: 'notes/' })}
                {textRow('publicBase', 'imageHostPublicBase')}
              </>
            )}
            {kind === 'smms' && textRow('token', 'imageHostToken')}
            {kind === 'github' && (
              <>
                {textRow('token', 'imageHostToken')}
                {textRow('owner', 'imageHostOwner')}
                {textRow('repo', 'imageHostRepo')}
                {textRow('branch', 'imageHostBranch', { placeholder: 'main' })}
                {textRow('dir', 'imageHostDir', { placeholder: 'images/' })}
                {textRow('urlPrefix', 'imageHostUrlPrefix')}
              </>
            )}
            {error && <div className="image-host-error">{error}</div>}
            <div className="image-host-actions">
              <button type="button" className="image-host-btn" onClick={onClose}>
                {t('aiCancel')}
              </button>
              <button
                type="button"
                className="image-host-btn image-host-btn-primary"
                disabled={saving}
                onClick={() => void save()}
              >
                {t('imageHostSave')}
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  )
}
