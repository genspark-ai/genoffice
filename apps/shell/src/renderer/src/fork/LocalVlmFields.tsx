import { SettingsRow } from './settings-ui/SettingsCard'
import { useCallback, useEffect, useState } from 'react'
import type { LocalVlmPickKind, LocalVlmStatus } from '../../../shared/fork/local-vlm'
import { useI18n } from '../locale'
import { localVlmString } from '../i18n/strings-local-vlm'
import './local-vlm.css'

export function LocalVlmFields() {
  const { lang } = useI18n()
  const [status, setStatus] = useState<LocalVlmStatus | null>(null)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string>('')

  const t = useCallback(
    (key: Parameters<typeof localVlmString>[1], params?: Record<string, string | number>) =>
      localVlmString(lang, key, params),
    [lang],
  )

  const refresh = useCallback(async () => {
    const next = await window.aiOffice.getLocalVlmStatus?.().catch(() => null)
    if (next) setStatus(next)
  }, [])

  useEffect(() => {
    void refresh()
    const off = window.aiOffice.onLocalVlmStatusChanged?.((next) => {
      setStatus(next)
    })
    return () => off?.()
  }, [refresh])

  const patchSettings = async (patch: Partial<LocalVlmStatus['settings']>) => {
    setError('')
    const next = await window.aiOffice.setLocalVlmSettings?.(patch).catch(() => null)
    if (next) setStatus(next)
  }

  const startDownload = async () => {
    if (busy) return
    setBusy(true)
    setError('')
    try {
      const res = await window.aiOffice.downloadLocalVlm?.({
        includeRuntime: true,
        includeModel: true,
      })
      if (res?.status) setStatus(res.status)
      if (res && !res.ok && res.error) setError(res.error)
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err))
    } finally {
      setBusy(false)
    }
  }

  const cancelDownload = async () => {
    await window.aiOffice.cancelLocalVlmDownload?.().catch(() => false)
  }

  const toggleServer = async () => {
    if (busy || !status) return
    setBusy(true)
    setError('')
    try {
      if (status.running) {
        const next = await window.aiOffice.stopLocalVlm?.()
        if (next) setStatus(next)
      } else {
        const res = await window.aiOffice.startLocalVlm?.()
        if (res?.status) setStatus(res.status)
        if (res && !res.ok && res.error) setError(res.error)
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err))
    } finally {
      setBusy(false)
    }
  }

  const pickFile = async (
    kind: LocalVlmPickKind,
    field: 'modelPath' | 'mmprojPath' | 'runtimePath',
  ) => {
    const picked = await window.aiOffice.pickLocalVlmFile?.(kind).catch(() => null)
    if (picked) {
      await patchSettings({ [field]: picked })
    }
  }

  if (!status) return null

  const missingParts: string[] = []
  if (!status.runtimeInstalled) missingParts.push('llama-server')
  if (!status.modelInstalled) missingParts.push('LFM3-VL-2B-Q4_K_M.gguf')
  if (!status.mmprojInstalled) missingParts.push('mmproj')

  const isDownloading = Boolean(status.download?.active)
  const ready = status.installed && status.runtimeInstalled

  return (
    <div className="set-local-vlm">
      <div className="set-local-vlm-status" aria-live="polite">
        <div className="set-field-stack">
          <div className="set-field-label">
            {isDownloading
              ? t('downloading', {
                  file: status.download!.file,
                  percent: status.download!.percent ?? 0,
                })
              : status.running && status.port
                ? t('statusRunning', { port: status.port })
                : ready
                  ? t('statusReady')
                  : t('statusMissing', { detail: missingParts.join(', ') })}
          </div>
          {!status.ramSufficient && (
            <div className="set-field-desc set-local-vlm-warn">
              {t('ramWarn', { mb: status.availableRamMB })}
            </div>
          )}
          {(error || status.lastError) && (
            <div className="set-field-desc set-local-vlm-warn">{error || status.lastError}</div>
          )}
        </div>
        <div className="set-local-vlm-actions">
          {isDownloading ? (
            <button type="button" className="set-btn" onClick={() => void cancelDownload()}>
              {t('cancelDownload')}
            </button>
          ) : !ready ? (
            <button
              type="button"
              className="set-btn primary"
              disabled={busy}
              onClick={() => void startDownload()}
            >
              {t('downloadBtn')}
            </button>
          ) : (
            <button
              type="button"
              className="set-btn"
              disabled={busy || status.starting}
              onClick={() => void toggleServer()}
            >
              {status.running ? t('stopBtn') : t('startBtn')}
            </button>
          )}
        </div>
      </div>

      {isDownloading && (
        <progress
          className="set-local-vlm-progress"
          max={100}
          value={status.download?.percent ?? 0}
        />
      )}

      <SettingsRow className="set-field">
        <div className="set-field-text">
          <div className="set-field-stack">
            <div className="set-field-label">{t('autoFallback')}</div>
            <div className="set-field-desc">{t('autoFallbackHint')}</div>
          </div>
        </div>
        <button
          type="button"
          className="set-switch"
          role="switch"
          aria-checked={status.settings.autoFallbackFromAgy}
          aria-label={t('autoFallback')}
          onClick={() =>
            void patchSettings({ autoFallbackFromAgy: !status.settings.autoFallbackFromAgy })
          }
        />
      </SettingsRow>

      <SettingsRow className="set-field">
        <div className="set-field-text">
          <label className="set-field-label" htmlFor="set-local-vlm-model-path">
            {t('modelPathLabel')}
          </label>
        </div>
        <div className="set-local-vlm-path-row">
          <input
            id="set-local-vlm-model-path"
            className="set-input"
            type="text"
            value={status.settings.modelPath}
            placeholder={status.modelPath}
            spellCheck={false}
            onChange={(e) => void patchSettings({ modelPath: e.target.value })}
          />
          <button
            type="button"
            className="set-btn"
            onClick={() => void pickFile('model', 'modelPath')}
          >
            {t('browse')}
          </button>
        </div>
      </SettingsRow>

      <SettingsRow className="set-field">
        <div className="set-field-text">
          <label className="set-field-label" htmlFor="set-local-vlm-mmproj-path">
            {t('mmprojPathLabel')}
          </label>
        </div>
        <div className="set-local-vlm-path-row">
          <input
            id="set-local-vlm-mmproj-path"
            className="set-input"
            type="text"
            value={status.settings.mmprojPath}
            placeholder={status.mmprojPath}
            spellCheck={false}
            onChange={(e) => void patchSettings({ mmprojPath: e.target.value })}
          />
          <button
            type="button"
            className="set-btn"
            onClick={() => void pickFile('mmproj', 'mmprojPath')}
          >
            {t('browse')}
          </button>
        </div>
      </SettingsRow>

      <SettingsRow className="set-field">
        <div className="set-field-text">
          <label className="set-field-label" htmlFor="set-local-vlm-runtime-path">
            {t('runtimePathLabel')}
          </label>
        </div>
        <div className="set-local-vlm-path-row">
          <input
            id="set-local-vlm-runtime-path"
            className="set-input"
            type="text"
            value={status.settings.runtimePath}
            placeholder={status.runtimePath ?? 'llama-server'}
            spellCheck={false}
            onChange={(e) => void patchSettings({ runtimePath: e.target.value })}
          />
          <button
            type="button"
            className="set-btn"
            onClick={() => void pickFile('runtime', 'runtimePath')}
          >
            {t('browse')}
          </button>
        </div>
      </SettingsRow>
    </div>
  )
}
