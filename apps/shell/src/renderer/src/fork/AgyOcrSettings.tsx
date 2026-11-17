import { SettingsCard, SettingsRow } from './settings-ui/SettingsCard'
import { useCallback, useEffect, useRef, useState } from 'react'
import type { Lang } from '@genoffice/i18n'
import { useI18n } from '../locale'
import { appConfirm } from '../ui-feedback'
import type {
  AgyOcrModelList,
  AgyOcrSettings as AgyOcrSettingsValue,
  AgyOcrStatus,
} from '../../../shared/fork/agy-ocr'
import { DEFAULT_LOCAL_OCR_SETTINGS, LOCAL_OCR_MAX_PAGES } from '../../../shared/fork/agy-ocr'
import { activityLine, agyOcrString, formatCount, type AgyOcrStringKey } from './agy-ocr-strings'
import { readIndexRequest } from './index-request'
import { localOcrCopy } from '../i18n/strings-local-ocr'
import { localVlmString } from '../i18n/strings-local-vlm'
import type { OcrEscalationBackend } from '../../../shared/fork/agy-ocr'
import './agy-ocr.css'

const POLL_MS = 3000

interface StatusApi {
  getAgyOcrStatus?: () => Promise<AgyOcrStatus | null>
}

/** Live reader status, polled only while the host is mounted and the window is visible. */
export function useAgyOcrStatus(
  api: StatusApi,
  intervalMs = POLL_MS,
): [AgyOcrStatus | null, () => void, boolean] {
  const [status, setStatus] = useState<AgyOcrStatus | null>(null)
  const [failed, setFailed] = useState(false)
  const kick = useRef<() => void>(() => undefined)
  const getStatus = api.getAgyOcrStatus
  useEffect(() => {
    let alive = true
    let inFlight = false
    let timer: ReturnType<typeof setTimeout> | undefined
    const load = async () => {
      if (!alive || inFlight) return
      if (!getStatus) {
        setFailed(true)
        return
      }
      if (timer) clearTimeout(timer)
      inFlight = true
      try {
        if (document.visibilityState === 'visible') {
          const next = await readIndexRequest(() => getStatus(), isAgyOcrStatus)
          if (alive) {
            setStatus(next)
            setFailed(false)
          }
        }
      } catch {
        if (alive) setFailed(true)
      } finally {
        inFlight = false
        if (alive) timer = setTimeout(() => void load(), intervalMs)
      }
    }
    kick.current = () => void load()
    const visible = () => {
      if (document.visibilityState === 'visible') void load()
    }
    document.addEventListener('visibilitychange', visible)
    void load()
    return () => {
      alive = false
      if (timer) clearTimeout(timer)
      document.removeEventListener('visibilitychange', visible)
    }
  }, [getStatus, intervalMs])
  const refresh = useCallback(() => kick.current(), [])
  return [status, refresh, failed]
}

function isAgyOcrStatus(value: unknown): value is AgyOcrStatus {
  if (!value || typeof value !== 'object') return false
  const status = value as Partial<AgyOcrStatus>
  return (
    !!status.settings &&
    typeof status.settings.enabled === 'boolean' &&
    !!status.activity &&
    typeof status.activity.kind === 'string' &&
    typeof status.running === 'boolean' &&
    (status.queuedDocumentIds === undefined ||
      (Array.isArray(status.queuedDocumentIds) &&
        status.queuedDocumentIds.every((id) => Number.isSafeInteger(id) && id > 0))) &&
    (status.progress === undefined ||
      (Number.isSafeInteger(status.progress.done) &&
        status.progress.done >= 0 &&
        Number.isSafeInteger(status.progress.total) &&
        status.progress.total > 0 &&
        status.progress.done <= status.progress.total)) &&
    !!status.tokensToday &&
    [
      status.tokensToday.input,
      status.tokensToday.output,
      status.tokensToday.thinking,
      status.pdfsToday,
      status.pagesToday,
      status.filesWaiting,
    ].every((value) => typeof value === 'number' && Number.isFinite(value) && value >= 0)
  )
}

function NumberField({
  value,
  min,
  max,
  label,
  disabled,
  onCommit,
}: {
  value: number
  min: number
  max: number
  label: string
  disabled: boolean
  onCommit: (next: number) => void
}) {
  const [draft, setDraft] = useState<string | null>(null)
  const commit = () => {
    if (draft === null) return
    // an empty box is not a 0: the value stays as it was
    const empty = draft.trim() === ''
    const parsed = Math.round(Number(draft))
    setDraft(null)
    if (empty || !Number.isFinite(parsed)) return
    const next = Math.min(max, Math.max(min, parsed))
    if (next !== value) onCommit(next)
  }
  return (
    <input
      type="number"
      className="set-input set-num-input"
      aria-label={label}
      min={min}
      max={max}
      step={1}
      disabled={disabled}
      value={draft ?? String(value)}
      onFocus={() => setDraft(String(value))}
      onChange={(event) => setDraft(event.target.value)}
      onBlur={commit}
      onKeyDown={(event) => {
        if (event.key === 'Enter') event.currentTarget.blur()
        if (event.key === 'Escape') setDraft(null)
      }}
    />
  )
}

function Switch({
  checked,
  label,
  disabled,
  onToggle,
}: {
  checked: boolean
  label: string
  disabled: boolean
  onToggle: () => void
}) {
  return (
    <button
      className="set-switch"
      type="button"
      role="switch"
      aria-checked={checked}
      aria-label={label}
      disabled={disabled}
      onClick={onToggle}
    />
  )
}

function statusLines(
  lang: Lang,
  status: AgyOcrStatus,
): Array<{ key: string; text: string; tone?: string }> {
  const lines: Array<{ key: string; text: string; tone?: string }> = []
  const tone =
    status.activity.kind === 'halted' || status.activity.kind === 'quota-unreadable'
      ? 'warn'
      : undefined
  lines.push({
    key: 'activity',
    text: activityLine(lang, status.activity),
    ...(tone ? { tone } : {}),
  })
  if (!status.settings.enabled) return lines
  lines.push({
    key: 'today',
    text: agyOcrString(lang, 'statusToday', {
      pdfs: formatCount(lang, status.pdfsToday),
      pages: formatCount(lang, status.pagesToday),
      tokens: formatCount(lang, status.tokensToday.input + status.tokensToday.output),
    }),
  })
  lines.push({
    key: 'waiting',
    text: agyOcrString(lang, 'statusWaiting', { count: formatCount(lang, status.filesWaiting) }),
  })
  if (status.lastError && status.activity.kind !== 'halted')
    lines.push({
      key: 'error',
      tone: 'warn',
      text: agyOcrString(lang, 'statusError', { message: status.lastError.message }),
    })
  if (status.lastResult)
    lines.push({
      key: 'last',
      text: agyOcrString(lang, 'statusLast', {
        pages: status.lastResult.pages,
        file: status.lastResult.file,
      }),
    })
  return lines
}

/** Document memory -> read scanned PDFs through Antigravity: consent, model, quota reserves, live status. */
export function AgyOcrSettings() {
  const { lang } = useI18n()
  const localCopy = localOcrCopy(lang)
  const isMac = navigator.platform.toLowerCase().includes('mac')
  const api = window.aiOffice
  const available = typeof api?.getAgyOcrStatus === 'function'
  const [status, refresh, statusFailed] = useAgyOcrStatus(api ?? {})
  const [message, setMessage] = useState('')
  const [checkingQuota, setCheckingQuota] = useState(false)
  const [saving, setSaving] = useState(false)
  const [models, setModels] = useState<AgyOcrModelList | null>(null)
  const [loadingModels, setLoadingModels] = useState(false)
  const loadedModels = useRef(false)

  const loadModels = useCallback(() => {
    if (loadedModels.current || typeof api?.listAgyOcrModels !== 'function') return
    loadedModels.current = true
    setLoadingModels(true)
    void readIndexRequest(
      () => api.listAgyOcrModels(),
      (value): value is AgyOcrModelList =>
        !!value && typeof value === 'object' && Array.isArray((value as AgyOcrModelList).models),
      20_000,
    )
      .then((list) => {
        setModels(list)
        // a failed listing may be retried the next time the picker is opened
        if (list.error || list.models.length === 0) loadedModels.current = false
      })
      .catch(() => {
        loadedModels.current = false
        setModels({ models: [], error: 'Model list unavailable. Open the picker to retry.' })
      })
      .finally(() => setLoadingModels(false))
  }, [api])

  const enabled = status?.settings.enabled === true
  useEffect(() => {
    if (enabled) loadModels()
  }, [enabled, loadModels])

  if (!available) return null

  const settings = status?.settings
  const save = (patch: Partial<AgyOcrSettingsValue>) => {
    setSaving(true)
    setMessage('')
    void readIndexRequest(
      () => api.setAgyOcrSettings(patch),
      (value): value is AgyOcrSettingsValue =>
        !!value &&
        typeof value === 'object' &&
        typeof (value as AgyOcrSettingsValue).enabled === 'boolean',
    )
      .then(() => refresh())
      .catch(() => {
        setMessage(
          lang === 'vi'
            ? 'Chưa xác nhận được cài đặt. Kiểm tra lại trước khi thử tiếp.'
            : 'Could not confirm the settings. Refresh before trying again.',
        )
        refresh()
      })
      .finally(() => setSaving(false))
  }
  const checkQuota = async () => {
    if (checkingQuota) return
    setCheckingQuota(true)
    setMessage('')
    try {
      if (api.refreshAgyOcrQuota)
        await readIndexRequest(() => api.refreshAgyOcrQuota!(), isAgyOcrStatus, 25_000)
      refresh()
    } catch {
      setMessage(
        lang === 'vi'
          ? 'Không đọc được hạn mức. Kiểm tra đăng nhập Antigravity rồi thử lại.'
          : 'Quota could not be read. Check your Antigravity sign-in and try again.',
      )
    } finally {
      setCheckingQuota(false)
    }
  }
  const disabled = !settings || saving
  const t = (key: AgyOcrStringKey, params?: Record<string, string | number>) =>
    agyOcrString(lang, key, params)

  const options = (models?.models ?? []).map((m) => ({
    id: m.id,
    label: m.cheapest ? `${m.id} · ${t('cheapest')}` : m.id,
  }))
  if (settings && !options.some((o) => o.id === settings.model))
    options.unshift({ id: settings.model, label: settings.model })
  const shownModel = options.find((o) => o.id === settings?.model)?.label ?? settings?.model ?? ''

  const lines = status ? statusLines(lang, status) : []
  return (
    <SettingsCard className="set-agyocr" title={t('title')} subtitle={t('desc')} icon="media">
      <SettingsRow className="set-field set-field-top">
        <div className="set-field-text">
          <div className="set-field-stack">
            <div className="set-field-label">{t('enable')}</div>
            <div className="set-field-desc set-agyocr-consent">{t('consent')}</div>
          </div>
        </div>
        <Switch
          checked={settings?.enabled ?? false}
          label={t('enable')}
          disabled={disabled}
          onToggle={() => settings && save({ enabled: !settings.enabled })}
        />
      </SettingsRow>

      <SettingsRow className="set-field set-field-top">
        <div className="set-field-text">
          <div className="set-field-stack">
            <div className="set-field-label">{localVlmString(lang, 'ocrBackendLabel')}</div>
            {status?.activeBackend === 'local-vlm' && (
              <div className="set-field-desc">{localVlmString(lang, 'ocrBackendLocalActive')}</div>
            )}
          </div>
        </div>
        <label className="set-select-wrap">
          <span className="set-select-text">
            {(settings?.escalationBackend ?? 'auto') === 'local-vlm'
              ? localVlmString(lang, 'ocrBackendLocal')
              : (settings?.escalationBackend ?? 'auto') === 'agy'
                ? localVlmString(lang, 'ocrBackendAgy')
                : localVlmString(lang, 'ocrBackendAuto')}
          </span>
          <select
            className="set-select"
            aria-label={localVlmString(lang, 'ocrBackendLabel')}
            value={settings?.escalationBackend ?? 'auto'}
            disabled={disabled}
            onChange={(event) =>
              save({ escalationBackend: event.target.value as OcrEscalationBackend })
            }
          >
            <option value="auto">{localVlmString(lang, 'ocrBackendAuto')}</option>
            <option value="agy">{localVlmString(lang, 'ocrBackendAgy')}</option>
            <option value="local-vlm">{localVlmString(lang, 'ocrBackendLocal')}</option>
          </select>
        </label>
      </SettingsRow>

      <SettingsRow className="set-field set-field-top">
        <div className="set-field-text">
          <div className="set-field-stack">
            <div className="set-field-label">{t('model')}</div>
            <div className="set-field-desc">
              {models?.error
                ? t('modelsFailed', { error: models.error })
                : loadingModels
                  ? t('modelsLoading')
                  : t('modelDesc')}
            </div>
          </div>
        </div>
        <label className="set-select-wrap" onMouseEnter={loadModels}>
          <span className="set-select-text">{shownModel}</span>
          <select
            className="set-select"
            aria-label={t('model')}
            value={settings?.model ?? ''}
            disabled={disabled}
            onFocus={loadModels}
            onChange={(event) => save({ model: event.target.value })}
          >
            {options.map((o) => (
              <option key={o.id} value={o.id}>
                {o.label}
              </option>
            ))}
          </select>
        </label>
      </SettingsRow>

      <section
        className="set-agyocr-overview"
        aria-label={lang === 'vi' ? 'Hạn mức và ngân sách OCR' : 'Quota and OCR budget'}
      >
        <div className="set-agyocr-overview-head">
          <div>
            <h5 className="set-agyocr-subtitle">
              {lang === 'vi' ? 'Chia sẻ hạn mức Antigravity' : 'Share Antigravity quota'}
            </h5>
            <p className="set-field-desc">
              {lang === 'vi'
                ? 'Hạn mức được chia theo tuổi của chu kỳ. OCR tự động chỉ dùng phần dư sau mức chừa cho bạn; đọc thủ công bỏ qua các giới hạn này.'
                : 'Quota is shared across the provider cycle. Automatic OCR uses only the share above your reserve; manual reading bypasses these limits.'}
            </p>
          </div>
          <button
            type="button"
            className="idx-btn"
            disabled={checkingQuota}
            onClick={() => void checkQuota()}
          >
            {checkingQuota
              ? lang === 'vi'
                ? 'Đang kiểm tra…'
                : 'Checking…'
              : lang === 'vi'
                ? 'Kiểm tra hạn mức'
                : 'Check quota'}
          </button>
        </div>
        <div className="set-agyocr-quota-grid">
          {(['weekly', 'fiveHour'] as const).map((key) => {
            const bucket = status?.quota?.[key]
            return (
              <div className="set-agyocr-quota" key={key}>
                <span>{t(key === 'weekly' ? 'weeklyTitle' : 'fiveHourTitle')}</span>
                <strong>{bucket ? `${Math.round(bucket.percent)}%` : '—'}</strong>
                <span>{lang === 'vi' ? 'còn lại' : 'remaining'}</span>
                {bucket && (
                  <progress
                    max={100}
                    value={bucket.percent}
                    aria-label={t(key === 'weekly' ? 'weeklyTitle' : 'fiveHourTitle')}
                  />
                )}
                <small>
                  {settings?.autoUnlimited
                    ? lang === 'vi'
                      ? 'Không giới hạn OCR tự động'
                      : 'Automatic OCR is unlimited'
                    : bucket
                      ? `${lang === 'vi' ? 'Chừa lại hiện tại' : 'Current reserve'} ${Math.round(bucket.floor)}%`
                      : lang === 'vi'
                        ? 'Chưa đọc được mức dự trữ hiện tại'
                        : 'Current reserve is not available'}
                  {status?.autoBudget && (
                    <>
                      <br />
                      {key === 'weekly'
                        ? `${lang === 'vi' ? 'OCR đã dùng trong ngày chu kỳ' : 'OCR used this cycle day'} ${Number(status.autoBudget.weeklyDaily.spent.toFixed(1))}% / ${status.autoBudget.weeklyDaily.limit}%`
                        : `${lang === 'vi' ? 'OCR đã dùng chu kỳ này' : 'OCR used this window'} ${Number(status.autoBudget.fiveHour.spent.toFixed(1))}%`}
                    </>
                  )}
                </small>
              </div>
            )
          })}
        </div>
        {status?.quota && (
          <p className="set-field-desc">
            {lang === 'vi' ? 'Kiểm tra gần nhất' : 'Last checked'}:{' '}
            {new Date(status.quota.readAt).toLocaleString(lang)}
            {Date.now() - status.quota.readAt > 300_000
              ? lang === 'vi'
                ? ' · Dữ liệu cũ, kiểm tra lại'
                : ' · Outdated, check again'
              : ''}
          </p>
        )}
        {status?.running && (
          <p className="set-field-desc" role="status">
            {lang === 'vi' ? 'Đang xử lý' : 'Processing'}:{' '}
            {status.stage
              ? {
                  quota: lang === 'vi' ? 'kiểm tra hạn mức' : 'checking quota',
                  rendering: lang === 'vi' ? 'chuẩn bị ảnh trang' : 'preparing page images',
                  recognizing: lang === 'vi' ? 'nhận dạng chữ' : 'recognizing text',
                  indexing: lang === 'vi' ? 'lưu vào chỉ mục' : 'adding to index',
                }[status.stage]
              : 'OCR'}
            {status.currentFile ? ` · ${status.currentFile.split(/[\\/]/).pop()}` : ''}
          </p>
        )}
        {!!status?.queuedDocuments && (
          <p className="set-field-desc">
            {status.queuedDocuments} {lang === 'vi' ? 'tệp đã xếp hàng' : 'files queued'}
          </p>
        )}
        {status?.running && api.cancelAgyOcr && (
          <button
            type="button"
            className="idx-btn"
            onClick={() => {
              void api.cancelAgyOcr!()
                .then(refresh)
                .catch(() =>
                  setMessage(lang === 'vi' ? 'Chưa dừng được tác vụ.' : 'Could not stop the task.'),
                )
            }}
          >
            {lang === 'vi' ? 'Dừng OCR' : 'Stop OCR'}
          </button>
        )}
      </section>
      {(statusFailed || message) && (
        <p className="set-agyocr-error" role="status">
          {message ||
            (lang === 'vi'
              ? 'Chưa cập nhật được trạng thái. Dữ liệu đang hiển thị có thể đã cũ.'
              : 'Status could not be refreshed. Shown data may be outdated.')}{' '}
          <button type="button" className="idx-btn" onClick={refresh}>
            {lang === 'vi' ? 'Thử lại' : 'Retry'}
          </button>
        </p>
      )}
      {!status && !statusFailed && (
        <p role="status">{lang === 'vi' ? 'Đang lấy trạng thái…' : 'Loading status…'}</p>
      )}
      <div className="set-agyocr-policy">
        <div className="set-agyocr-overview-head">
          <p className="set-field-desc">
            {settings?.autoUnlimited
              ? lang === 'vi'
                ? 'Đang không giới hạn: OCR tự động có thể dùng hết hạn mức chung.'
                : 'Unlimited mode: automatic OCR may use all shared quota.'
              : lang === 'vi'
                ? `Tự động dùng tối đa ${settings?.autoWeeklyDailyBudgetPercent ?? 12}% hạn mức tuần mỗi ngày chu kỳ; chừa tối thiểu 16% hạn mức tuần và 20% hạn mức 5 giờ.`
                : `Automatic OCR uses at most ${settings?.autoWeeklyDailyBudgetPercent ?? 12}% of weekly quota per cycle day; reserves at least 16% weekly and 20% of the 5-hour quota.`}
          </p>
          <button
            type="button"
            className="idx-btn"
            aria-pressed={!!settings?.autoUnlimited}
            disabled={disabled}
            onClick={async () => {
              if (settings?.autoUnlimited) {
                save({ autoUnlimited: false })
                return
              }
              if (
                await appConfirm(
                  lang === 'vi'
                    ? 'Không giới hạn OCR tự động? Sẽ bỏ ngân sách, mức dự trữ và giới hạn số PDF mỗi ngày của GenOffice. OCR có thể dùng hết hạn mức chung; giới hạn của Antigravity vẫn áp dụng.'
                    : 'Remove automatic OCR limits? GenOffice budgets, reserves and daily PDF limits will be bypassed. OCR may use all shared quota; Antigravity limits still apply.',
                  {
                    confirmLabel: lang === 'vi' ? 'Bật không giới hạn' : 'Enable unlimited',
                    tone: 'info',
                  },
                )
              )
                save({ autoUnlimited: true })
            }}
          >
            {settings?.autoUnlimited
              ? lang === 'vi'
                ? 'Bật lại chia sẻ hạn mức'
                : 'Restore quota sharing'
              : lang === 'vi'
                ? 'Không giới hạn'
                : 'Unlimited'}
          </button>
        </div>
        {!settings?.autoUnlimited && (
          <div className="set-agyocr-cycle">
            <span className="set-field-label">
              {lang === 'vi'
                ? 'Mức chừa theo giờ trong chu kỳ 5 giờ'
                : 'Reserve by hour in the 5-hour window'}
            </span>
            <div className="set-agyocr-cycle-steps">
              {[80, 60, 40, 20, 20].map((floor, i) => (
                <div key={i}>
                  <span>
                    {lang === 'vi' ? 'Giờ' : 'Hour'} {i + 1}
                  </span>
                  <strong>{floor}%</strong>
                </div>
              ))}
            </div>
            <p className="set-field-desc">
              {lang === 'vi'
                ? 'Đây là mức còn lại của cả tài khoản, gồm cả chat và các bên dùng chung. Chu kỳ tính theo thời điểm làm mới của Antigravity.'
                : 'These are remaining shares for the whole account, including chat and other users. Windows follow Antigravity reset times.'}
            </p>
          </div>
        )}
      </div>
      {!settings?.autoUnlimited && (
        <div className="set-agyocr-cycle">
          <span className="set-field-label">
            {lang === 'vi'
              ? 'Mức chừa theo ngày trong chu kỳ tuần'
              : 'Reserve by day in the weekly cycle'}
          </span>
          <div className="set-agyocr-cycle-steps set-agyocr-week-steps">
            {Array.from({ length: 7 }, (_, i) => (
              <div key={i}>
                <span>
                  {lang === 'vi' ? 'Ngày' : 'Day'} {i + 1}
                </span>
                <strong>
                  {Math.max(16, 100 - (settings?.autoWeeklyDailyBudgetPercent ?? 12) * (i + 1))}%
                </strong>
              </div>
            ))}
          </div>
          {status?.autoBudget && status.autoBudget.weeklyDaily.resetAt > 0 && (
            <p className="set-field-desc">
              {lang === 'vi' ? 'Ngân sách ngày tiếp theo mở lúc' : 'Next daily budget opens at'}{' '}
              {new Date(status.autoBudget.weeklyDaily.resetAt).toLocaleString(lang)}
              {status.autoBudget.pending
                ? lang === 'vi'
                  ? ' · Đang chờ Antigravity cập nhật mức dùng.'
                  : ' · Waiting for Antigravity usage to update.'
                : ''}
            </p>
          )}
        </div>
      )}
      {!!settings?.maxPdfsPerDay && !settings.autoUnlimited && (
        <p className="set-field-desc">
          {lang === 'vi'
            ? `Đang bật thêm giới hạn ${settings.maxPdfsPerDay} PDF/ngày. Bạn có thể đổi trong phần điều chỉnh bên dưới.`
            : `An additional ${settings.maxPdfsPerDay} PDFs/day cap is enabled. You can change it below.`}
        </p>
      )}
      <details className="set-agyocr-advanced">
        <summary>
          {lang === 'vi'
            ? 'Điều chỉnh ngân sách và lịch đọc'
            : 'Adjust budget and reading schedule'}
        </summary>
        <p className="set-field-desc">
          {lang === 'vi'
            ? 'Mức dùng được đo từ thay đổi hạn mức do Antigravity trả về; có thể cập nhật trễ hoặc gồm việc dùng tài khoản cùng lúc. Không biết trước chính xác chi phí của một lượt đọc, nên lượt cuối có thể vượt ngưỡng; khi đủ ngân sách sẽ không bắt đầu lượt tự động tiếp theo.'
            : 'Usage is measured from Antigravity quota changes, which may be delayed or include simultaneous account use. A call has no known upfront cost, so the last call may cross the threshold; no further automatic call starts once the budget is spent.'}
        </p>
        <SettingsRow className="set-field">
          <div className="set-field-text">
            <div className="set-field-stack">
              <div className="set-field-label">
                {lang === 'vi'
                  ? 'Hạn mức tuần dùng mỗi ngày chu kỳ (%)'
                  : 'Weekly quota per cycle day (%)'}
              </div>
              <div className="set-field-desc">
                {lang === 'vi'
                  ? 'Mặc định 12%. 0 sẽ dừng OCR tự động; phần chưa dùng không cộng dồn.'
                  : 'Default 12%. Zero pauses automatic OCR; unused budget does not roll over.'}
              </div>
            </div>
          </div>
          <NumberField
            value={settings?.autoWeeklyDailyBudgetPercent ?? 12}
            min={0}
            max={100}
            label={lang === 'vi' ? 'Ngân sách tuần mỗi ngày' : 'Weekly daily budget'}
            disabled={disabled || !!settings?.autoUnlimited}
            onCommit={(autoWeeklyDailyBudgetPercent) => save({ autoWeeklyDailyBudgetPercent })}
          />
        </SettingsRow>
        <button
          type="button"
          className="idx-btn"
          disabled={disabled}
          onClick={() =>
            save({ autoWeeklyDailyBudgetPercent: 12, autoUnlimited: false, maxPdfsPerDay: 0 })
          }
        >
          {lang === 'vi' ? 'Đặt lại mức chia sẻ mặc định' : 'Restore default sharing'}
        </button>
        <SettingsRow className="set-field set-field-top">
          <div className="set-field-text">
            <div className="set-field-label">{localCopy.enable}</div>
          </div>
          <Switch
            checked={(settings?.localOcr ?? DEFAULT_LOCAL_OCR_SETTINGS).enabled}
            label={localCopy.enable}
            disabled={disabled}
            onToggle={() =>
              settings &&
              save({
                localOcr: {
                  ...(settings.localOcr ?? DEFAULT_LOCAL_OCR_SETTINGS),
                  enabled: !(settings.localOcr ?? DEFAULT_LOCAL_OCR_SETTINGS).enabled,
                },
              })
            }
          />
        </SettingsRow>

        <SettingsRow className="set-field set-field-top">
          <div className="set-field-text">
            <div className="set-field-label">{localCopy.engine}</div>
            <div className="set-field-desc">{localCopy.autoHint}</div>
          </div>
          <label className="set-select-wrap">
            <span className="set-select-text">
              {settings?.localOcr?.engine === 'tesseract-vie'
                ? localCopy.tesseract
                : settings?.localOcr?.engine === 'apple-vision'
                  ? localCopy.vision
                  : settings?.localOcr?.engine === 'trueocr-vie'
                    ? 'TrueOCR'
                    : localCopy.auto}
            </span>
            <select
              className="set-select"
              aria-label={localCopy.engine}
              disabled={disabled}
              value={(settings?.localOcr ?? DEFAULT_LOCAL_OCR_SETTINGS).engine}
              onChange={(event) =>
                settings &&
                save({
                  localOcr: {
                    ...(settings.localOcr ?? DEFAULT_LOCAL_OCR_SETTINGS),
                    engine: event.target
                      .value as import('../../../shared/fork/agy-ocr').LocalOcrEngineChoice,
                  },
                })
              }
            >
              <option value="auto">{localCopy.auto}</option>
              <option value="tesseract-vie">{localCopy.tesseract}</option>
              {isMac && <option value="apple-vision">{localCopy.vision}</option>}
              {settings?.localOcr?.engine === 'trueocr-vie' && (
                <option value="trueocr-vie">TrueOCR</option>
              )}
            </select>
          </label>
        </SettingsRow>

        <SettingsRow className="set-field set-field-top">
          <div className="set-field-text">
            <div className="set-field-stack">
              <div className="set-field-label">{t('maxPdfsPerDay')}</div>
              <div className="set-field-desc">{t('maxPdfsPerDayDesc')}</div>
            </div>
          </div>
          <NumberField
            value={settings?.maxPdfsPerDay ?? 0}
            min={0}
            max={2000}
            label={t('maxPdfsPerDay')}
            disabled={disabled}
            onCommit={(maxPdfsPerDay) => save({ maxPdfsPerDay })}
          />
        </SettingsRow>

        <SettingsRow className="set-field set-field-top">
          <div className="set-field-text">
            <div className="set-field-stack">
              <div className="set-field-label">{t('maxPagesPerFile')}</div>
              <div className="set-field-desc">{t('maxPagesPerFileDesc')}</div>
            </div>
          </div>
          <NumberField
            value={settings?.maxPagesPerFile ?? 0}
            min={0}
            max={100000}
            label={t('maxPagesPerFile')}
            disabled={disabled}
            onCommit={(maxPagesPerFile) => save({ maxPagesPerFile })}
          />
        </SettingsRow>

        <SettingsRow className="set-field set-field-top">
          <div className="set-field-text">
            <div className="set-field-stack">
              <div className="set-field-label">{t('localMaxPages')}</div>
              <div className="set-field-desc">{t('localMaxPagesDesc')}</div>
            </div>
          </div>
          <NumberField
            value={(settings?.localOcr ?? DEFAULT_LOCAL_OCR_SETTINGS).maxPages}
            min={1}
            max={LOCAL_OCR_MAX_PAGES}
            label={t('localMaxPages')}
            disabled={disabled}
            onCommit={(maxPages) =>
              settings &&
              save({ localOcr: { ...(settings.localOcr ?? DEFAULT_LOCAL_OCR_SETTINGS), maxPages } })
            }
          />
        </SettingsRow>

        <SettingsRow className="set-field set-field-top">
          <div className="set-field-text">
            <div className="set-field-stack">
              <div className="set-field-label">{t('onlyAC')}</div>
              <div className="set-field-desc">{t('onlyACDesc')}</div>
            </div>
          </div>
          <Switch
            checked={settings?.onlyOnAC ?? true}
            label={t('onlyAC')}
            disabled={disabled}
            onToggle={() => settings && save({ onlyOnAC: !settings.onlyOnAC })}
          />
        </SettingsRow>

        <SettingsRow className="set-field set-field-top">
          <div className="set-field-text">
            <div className="set-field-stack">
              <div className="set-field-label">{t('onlyIdle')}</div>
              <div className="set-field-desc">{t('onlyIdleDesc')}</div>
            </div>
          </div>
          <Switch
            checked={settings?.onlyWhenIdle ?? true}
            label={t('onlyIdle')}
            disabled={disabled}
            onToggle={() => settings && save({ onlyWhenIdle: !settings.onlyWhenIdle })}
          />
        </SettingsRow>
      </details>
      <ul className="set-agyocr-status" role="status" aria-live="polite">
        {lines.map((line) => (
          <li key={line.key} data-tone={line.tone}>
            {line.text}
          </li>
        ))}
      </ul>
    </SettingsCard>
  )
}
