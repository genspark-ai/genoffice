import { SettingsRow } from './settings-ui/SettingsCard'
import { useCallback, useEffect, useRef, useState, type Dispatch, type SetStateAction } from 'react'
import type { AiProviderConfig } from '@genoffice/ai-provider/browser'
import { useI18n, type TFunc } from '../locale'
import { agyString } from './agy-strings'
import { localVlmString } from '../i18n/strings-local-vlm'
import {
  agyConnectionOf,
  agyPlatformOf,
  withAgyModels,
  type AgyCatalog,
  type AgyConnection,
} from './agy-provider-state'

/** Runs `agy models` through the shared provider-models IPC (never starts a model request). */
export async function fetchAgyCatalog(cliPath: string | undefined): Promise<AgyCatalog | null> {
  const get = window.aiOffice.getProviderModels
  if (!get) return null
  const config: AiProviderConfig = { apiKey: '', model: '', cliPath: cliPath?.trim() || undefined }
  return (await get('agy', config)) as AgyCatalog
}

/** Header "Test connection" for agy: lists models instead of spending a real model call. */
export async function testAgyConnection(
  cliPath: string | undefined,
  fallbackError: string,
): Promise<{ ok: boolean; error?: string }> {
  try {
    const connection = agyConnectionOf(await fetchAgyCatalog(cliPath), fallbackError)
    return connection.state === 'missing' ? { ok: false, error: connection.error } : { ok: true }
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : fallbackError }
  }
}

/** Settings "Test connection": agy lists models; every other provider uses the shared IPC. */
export function testAiSettingsFor(
  settings: Parameters<NonNullable<Window['aiOffice']['testAiSettings']>>[0],
  fallbackError: string,
) {
  if (settings.provider === 'agy') {
    return testAgyConnection(settings.providers.agy?.cliPath, fallbackError)
  }
  if (settings.provider === 'local-vlm') {
    return window.aiOffice.testLocalVlm?.()
  }
  return window.aiOffice.testAiSettings?.(settings)
}

/** The one-paragraph description under the provider picker (agy's is honest about its limits). */
export function ProviderNote({ provider, t }: { provider: string; t: TFunc }) {
  const { lang } = useI18n()
  if (provider === 'genspark') return <>{t('setAiGensparkHint')}</>
  if (provider === 'codex') return <>{t('setAiCodexHint')}</>
  if (provider === 'local-vlm') return <>{localVlmString(lang, 'note')}</>
  return <>{provider === 'agy' ? agyString(lang, 'agyNote') : t('setAiByokNote')}</>
}

interface Props<T extends { id: string; models: string[]; defaultModel: string }> {
  config: { cliPath?: string; model: string }
  update(patch: { cliPath: string }): void
  /** the live agy catalog replaces the agy entry's model list */
  setCatalog: Dispatch<SetStateAction<T[]>>
}

/** Settings → AI Model block for the Antigravity CLI: connection state, CLI path, re-check. */
export function AgyProviderFields<
  T extends { id: string; models: string[]; defaultModel: string },
>({ config, update, setCatalog }: Props<T>) {
  const { lang } = useI18n()
  const cliPath = config.cliPath ?? ''
  const onCatalog = (live: AgyCatalog) => setCatalog((c) => withAgyModels(c, live, config.model))
  const [connection, setConnection] = useState<AgyConnection>({ state: 'checking' })
  const [localFallbackReady, setLocalFallbackReady] = useState<boolean | null>(null)
  const onCatalogRef = useRef(onCatalog)
  onCatalogRef.current = onCatalog
  const failure = agyString(lang, 'agyTestFailed')
  const failureRef = useRef(failure)
  failureRef.current = failure
  const seq = useRef(0)

  const check = useCallback(async (path: string) => {
    const mine = ++seq.current
    setConnection({ state: 'checking' })
    try {
      const catalog = await fetchAgyCatalog(path)
      if (mine !== seq.current) return
      setConnection(agyConnectionOf(catalog, failureRef.current))
      if (catalog) onCatalogRef.current(catalog)
    } catch (error) {
      if (mine !== seq.current) return
      setConnection({
        state: 'missing',
        error: error instanceof Error ? error.message : failureRef.current,
      })
    }
  }, [])

  useEffect(() => {
    void check(cliPath)
    void window.aiOffice
      .getLocalVlmStatus?.()
      .then((s) => {
        if (s) {
          setLocalFallbackReady(
            s.installed &&
              s.runtimeInstalled &&
              s.settings.enabled &&
              s.settings.autoFallbackFromAgy,
          )
        }
      })
      .catch(() => undefined)
    // first mount only: later path edits re-check on blur
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const platform = agyPlatformOf(navigator.userAgent)
  const where = agyString(lang, platform === 'win' ? 'agyWhereWin' : 'agyWhereMac')
  return (
    <>
      <div className="set-field-desc set-model-discovery" aria-live="polite">
        <span>
          {connection.state === 'checking'
            ? agyString(lang, 'agyChecking')
            : connection.state === 'connected'
              ? agyString(lang, 'agyConnected', { n: connection.count })
              : agyString(lang, 'agyNotFound', { error: connection.error })}
        </span>
        <button
          className="set-btn"
          type="button"
          disabled={connection.state === 'checking'}
          onClick={() => void check(cliPath)}
        >
          {agyString(lang, 'agyRecheck')}
        </button>
      </div>
      {connection.state === 'missing' && localFallbackReady !== null && (
        <div className="set-local-vlm-banner" role="status">
          {localVlmString(lang, localFallbackReady ? 'agyFallbackReady' : 'agyFallbackMissing')}
        </div>
      )}
      <SettingsRow className="set-field">
        <div className="set-field-text">
          <div className="set-field-stack">
            <label className="set-field-label" htmlFor="set-ai-agy-path">
              {agyString(lang, 'agyPathLabel')}
            </label>
            <div className="set-field-desc">{agyString(lang, 'agyPathHint', { where })}</div>
          </div>
        </div>
        <input
          id="set-ai-agy-path"
          className="set-input"
          type="text"
          value={cliPath}
          placeholder={agyString(lang, 'agyPathPlaceholder')}
          spellCheck={false}
          autoComplete="off"
          onChange={(e) => update({ cliPath: e.target.value.trim() })}
          onBlur={(e) => void check(e.target.value.trim())}
        />
      </SettingsRow>
    </>
  )
}
