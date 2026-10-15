import { useEffect, useMemo, useRef, useState } from 'react'
import {
  AgentLoop,
  createFeaturesSkill,
  createIpcTransport,
  type FeatureRegistryDeps,
} from '@genoffice/agent-core'
import { AiComposer, AiTypingIndicator, Markdown } from '@genoffice/ui'
import type { AiSettings } from '@genoffice/ai-provider'
import { useI18n } from './locale'
import type { ScriptRunOutcome } from '@genoffice/agent-core'
import type {
  HookActivity,
  PackDescription,
  ScriptApplyResult,
} from '../../shared/feature-channels'
import { getFeatures, refreshFeatures, useFeatures } from './slots/features-state'
import { onAssistantRequest, takePendingPrompt } from './slots/assistant-bridge'
import './mods-assistant.css'

/**
 * The home screen's mod assistant: the one surface that talks to the feature
 * registry.
 *
 * It is handed the features skill — list, get, create, update, delete, and a
 * read-only document reader — and nothing else. No tool here writes a document,
 * and no kind executes a feature yet, so the panel shows what was recorded
 * rather than implying that something runs.
 *
 * The layout follows the home screen's own language: a hero line, the
 * conversation, and the registry beside it as cards. Everything is painted from
 * the shared tokens, so light and dark follow the rest of the shell.
 *
 * The conversation lives in memory only; the registry it manages is on disk, so
 * what the assistant records survives a restart even though the transcript does
 * not.
 */

interface AssistantMessage {
  role: 'user' | 'assistant'
  text: string
  /** the tool names this answer called, in order */
  toolNotes: string[]
  streaming?: boolean
  isError?: boolean
}

/** The file name of a path, for a compact activity line. */
function lastSegment(path: string): string {
  const at = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  return at === -1 ? path : path.slice(at + 1)
}

/** The time an activity entry was recorded, in the user's own clock. */
function clockOf(iso: string): string {
  const at = new Date(iso)
  return Number.isNaN(at.getTime()) ? '' : at.toLocaleTimeString()
}

/** One script's run state in the panel: what it reported, and against which file. */
interface ScriptState {
  running?: boolean
  dryRunPath?: string
  outcome?: ScriptRunOutcome
  applied?: ScriptApplyResult
  error?: string
}

export function ModsAssistant({ onOpenAsk }: { onOpenAsk?: () => void }): React.JSX.Element {
  const { t } = useI18n()
  const [messages, setMessages] = useState<AssistantMessage[]>([])
  // a slot's press parks a prompt here before Home switches to this panel
  const [draft, setDraft] = useState(() => takePendingPrompt())
  const [busy, setBusy] = useState(false)
  // the shared registry: Home draws the same copy, so a slot appears as soon as
  // the assistant records one
  const features = useFeatures()
  const scriptFeatures = useMemo(() => features.filter((feature) => feature.script), [features])
  const knowledgeFeatures = useMemo(
    () => features.filter((feature) => feature.knowledge),
    [features],
  )
  // what the knowledge base reports for these sources: documents, and whether a
  // path is missing on this machine (a pack can carry an absolute path that
  // never existed here — flagged, never searched)
  const [sourceStates, setSourceStates] = useState<
    Record<string, { documents: number; missing: boolean; enabled: boolean }>
  >({})
  useEffect(() => {
    if (typeof window.aiOffice?.kbSources !== 'function') return
    let cancelled = false
    void window.aiOffice
      .kbSources()
      .then((sources) => {
        if (cancelled) return
        const states: Record<string, { documents: number; missing: boolean; enabled: boolean }> = {}
        for (const source of sources)
          states[source.id] = {
            documents: source.documents,
            missing: source.missing ?? false,
            enabled: source.enabled,
          }
        setSourceStates(states)
      })
      .catch(() => undefined)
    return () => {
      cancelled = true
    }
  }, [knowledgeFeatures.length])
  const settingsRef = useRef<AiSettings | null>(null)
  const loopRef = useRef<AgentLoop<unknown> | null>(null)
  const scrollRef = useRef<HTMLDivElement>(null)
  // the scripts a user can try: recent documents to point them at, the document
  // chosen per feature, and what each run reported
  const [recents, setRecents] = useState<{ path: string; name: string }[]>([])
  const [scriptPaths, setScriptPaths] = useState<Record<string, string>>({})
  const [scriptRuns, setScriptRuns] = useState<Record<string, ScriptState>>({})
  // what the hooks have seen lately; a record, not a control
  const [hookActivity, setHookActivity] = useState<HookActivity[]>([])
  // a pack that has been read but not installed: what it declares, and where it is
  const [pack, setPack] = useState<{
    path: string
    description?: PackDescription
    errors?: string[]
  } | null>(null)

  useEffect(() => {
    void window.aiOffice
      .recents({ limit: 12 })
      .then((page) =>
        setRecents(page.entries.map((entry) => ({ path: entry.path, name: entry.name }))),
      )
      .catch(() => undefined)
  }, [])

  const exportPack = async (): Promise<void> => {
    try {
      await window.aiOffice.featuresPackExport()
    } catch {
      // the dialog is the feedback; a cancelled export is not an error
    }
  }

  const importPack = async (): Promise<void> => {
    try {
      const picked = await window.aiOffice.featuresPackPreview()
      if (picked.path)
        setPack({ path: picked.path, description: picked.description, errors: picked.errors })
    } catch (err) {
      setPack(null)
      patchScript('pack', { error: err instanceof Error ? err.message : String(err) })
    }
  }

  const installPack = async (): Promise<void> => {
    if (!pack?.path) return
    try {
      await window.aiOffice.featuresPackInstall(pack.path)
      setPack(null)
      refreshFeatures()
    } catch (err) {
      setPack({ path: pack.path, errors: [err instanceof Error ? err.message : String(err)] })
    }
  }

  const refreshHooks = (): void => {
    void window.aiOffice
      .featuresHookActivity()
      .then((entries) => setHookActivity(entries))
      .catch(() => undefined)
  }

  const patchScript = (id: string, patch: ScriptState): void =>
    setScriptRuns((prev) => ({ ...prev, [id]: { ...prev[id], ...patch } }))

  const dryRun = async (id: string, path: string): Promise<void> => {
    patchScript(id, { running: true, error: undefined, outcome: undefined, applied: undefined })
    try {
      const outcome = await window.aiOffice.featuresRun(id, { path })
      patchScript(id, { running: false, outcome, dryRunPath: path })
    } catch (err) {
      patchScript(id, { running: false, error: err instanceof Error ? err.message : String(err) })
    }
  }

  const applyRun = async (id: string, path: string): Promise<void> => {
    patchScript(id, { running: true, error: undefined })
    try {
      // the main process checks both gates again; this is the user's deliberate
      // second press after reading the plan
      const applied = await window.aiOffice.featuresApply(id, { path, confirm: true })
      patchScript(id, { running: false, applied })
    } catch (err) {
      patchScript(id, { running: false, error: err instanceof Error ? err.message : String(err) })
    }
  }

  /** Apply is offered only after a dry run of this feature against this document. */
  const canApply = (id: string): boolean => {
    const state = scriptRuns[id]
    const path = scriptPaths[id]
    return Boolean(path) && Boolean(state?.outcome) && state?.dryRunPath === path && !state?.running
  }

  const refreshSettings = (): void => {
    void window.aiOffice
      .getAiSettings()
      .then((settings) => {
        settingsRef.current = settings
      })
      .catch(() => undefined)
  }

  const patchLast = (
    patch: Partial<AssistantMessage> | ((last: AssistantMessage) => Partial<AssistantMessage>),
  ): void => {
    setMessages((prev) => {
      if (prev.length === 0 || prev[prev.length - 1]!.role !== 'assistant') return prev
      const last = prev[prev.length - 1]!
      const delta = typeof patch === 'function' ? patch(last) : patch
      return [...prev.slice(0, -1), { ...last, ...delta }]
    })
  }

  const finishLast = (): void => {
    setMessages((prev) => {
      if (prev.length === 0) return prev
      const last = prev[prev.length - 1]!
      return [...prev.slice(0, -1), { ...last, streaming: false }]
    })
    setBusy(false)
    // a create/update/delete inside the turn changed the registry: show it now
    refreshFeatures()
  }

  useEffect(() => {
    refreshSettings()
    refreshFeatures()
    refreshHooks()
    const deps: FeatureRegistryDeps = {
      // buildContext reads the cached snapshot; the tools go to the registry
      snapshot: () => getFeatures(),
      list: () => window.aiOffice.featuresList(),
      get: (id) => window.aiOffice.featuresGet(id),
      create: (feature) => window.aiOffice.featuresCreate(feature),
      update: (id, patch) => window.aiOffice.featuresUpdate(id, patch),
      remove: (id) => window.aiOffice.featuresDelete(id),
      readDocument: (path) => window.aiOffice.featuresReadDocument(path),
      runScript: (id, request) => window.aiOffice.featuresRun(id, request),
    }
    loopRef.current = new AgentLoop({
      transport: createIpcTransport<AiSettings>({
        onStream: (listener) => window.aiOffice.onAiStream(listener),
        start: (request) => window.aiOffice.aiStream(request),
        cancel: (requestId) => void window.aiOffice.aiStreamCancel(requestId),
        getSettings: () => settingsRef.current as AiSettings,
        unknownErrorText: () => t('modsUnknownError'),
      }),
      skill: createFeaturesSkill(deps),
      events: {
        onText: (text) => patchLast({ text }),
        onToolStart: (call) => patchLast((last) => ({ toolNotes: [...last.toolNotes, call.name] })),
        onToolExecuted: () => {
          refreshFeatures()
          refreshHooks()
        },
        onDone: () => finishLast(),
        onError: (error) =>
          patchLast((last) => ({
            text: last.text || error,
            streaming: false,
            isError: !last.text,
          })),
      },
    })
    return () => loopRef.current?.cancel()
    // eslint-disable-next-line react-hooks/exhaustive-deps -- mount-only; the loop reads live refs
  }, [])

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight, behavior: 'smooth' })
  }, [messages])

  // a slot press while this panel is already open lands in the composer
  useEffect(() => onAssistantRequest((prompt) => setDraft(prompt)), [])

  const ask = async (): Promise<void> => {
    const instruction = draft.trim()
    if (!instruction || busy) return
    setDraft('')
    setBusy(true)
    refreshSettings()
    // the listing rides along with the turn, so attach the freshest copy
    refreshFeatures()
    refreshHooks()
    setMessages((prev) => [
      ...prev,
      { role: 'user', text: instruction, toolNotes: [] },
      { role: 'assistant', text: '', toolNotes: [], streaming: true },
    ])
    try {
      await loopRef.current!.run(instruction)
    } catch {
      finishLast()
    }
  }

  /**


   * The hooks' recent activity: what each one saw when its event happened.

   */
  const hooksSection = (): React.JSX.Element | null => {
    if (!hookActivity.length) return null
    return (
      <section className="mods-scripts" aria-label={t('modsHooks')}>
        <h3 className="mods-scripts-title">{t('modsHooks')}</h3>
        <ul className="mods-scripts-list">
          {hookActivity
            .slice(-6)
            .reverse()
            .map((entry, index) => (
              <li key={index} className="mods-hook">
                <div className="mods-feature-top">
                  <span className="mods-feature-name">{entry.hookId}</span>
                  <span className="mods-feature-kind">{entry.event}</span>
                </div>
                <p className="mods-hook-when">
                  {lastSegment(entry.path)}
                  {clockOf(entry.at) ? ' · ' + clockOf(entry.at) : ''}
                </p>
                {entry.log.map((line, lineIndex) => (
                  <p key={lineIndex} className="mods-hook-line">
                    {line}
                  </p>
                ))}
                {(entry.vetoes ?? []).map((reason, vetoIndex) => (
                  <p key={vetoIndex} className="mods-script-error">
                    {'⊘ ' + reason}
                  </p>
                ))}
                {entry.error ? <p className="mods-script-error">{entry.error}</p> : null}
              </li>
            ))}
        </ul>
      </section>
    )
  }

  const featureCount = features?.length ?? 0
  const registryTitle = t('modsFeatures', { n: String(featureCount) })

  return (
    <div className="mods-assistant">
      <header className="mods-hero">
        <span className="mods-hero-icon" aria-hidden="true">
          <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
            <path
              d="M10 2.5 11.9 7l4.6.5-3.4 3.1 1 4.5L10 12.8l-4.1 2.3 1-4.5L3.5 7.5 8.1 7z"
              stroke="currentColor"
              strokeWidth="1.4"
              strokeLinejoin="round"
            />
          </svg>
        </span>
        <div className="mods-hero-text">
          <h1 className="mods-hero-title">{t('modsTitle')}</h1>
          <p className="mods-hero-sub">{t('modsSubtitle')}</p>
        </div>
      </header>

      <div className="mods-body">
        <section className="mods-chat" aria-label={t('modsTitle')}>
          <div className="mods-thread" ref={scrollRef}>
            {messages.length === 0 ? (
              <div className="mods-intro">
                <span className="mods-intro-icon" aria-hidden="true">
                  <svg width="22" height="22" viewBox="0 0 22 22" fill="none">
                    <rect
                      x="3"
                      y="4.5"
                      width="16"
                      height="12"
                      rx="2.5"
                      stroke="currentColor"
                      strokeWidth="1.4"
                    />
                    <path
                      d="M6.5 9h9M6.5 12.2h5.5"
                      stroke="currentColor"
                      strokeWidth="1.4"
                      strokeLinecap="round"
                    />
                  </svg>
                </span>
                <p className="mods-intro-text">{t('modsIntro')}</p>
              </div>
            ) : null}
            {messages.map((message, index) => (
              <div
                key={index}
                className={
                  'mods-msg ' +
                  (message.role === 'user' ? 'mods-msg-user' : 'mods-msg-reply') +
                  (message.isError ? ' mods-msg-error' : '')
                }
              >
                {message.role === 'assistant' && message.toolNotes.length > 0 && (
                  <div className="mods-tools">
                    {message.toolNotes.map((note, at) => (
                      <span key={at} className="mods-tool-note">
                        {note}
                        {message.streaming && at === message.toolNotes.length - 1 ? '…' : ''}
                      </span>
                    ))}
                  </div>
                )}
                {message.text ? (
                  message.role === 'assistant' ? (
                    <Markdown text={message.text} />
                  ) : (
                    message.text
                  )
                ) : message.streaming ? (
                  <AiTypingIndicator label={t('modsHintBusy')} />
                ) : null}
              </div>
            ))}
          </div>
          <div className="mods-composer">
            <AiComposer
              value={draft}
              busy={busy}
              placeholder={t('modsPlaceholder')}
              hintIdle={t('modsHintIdle')}
              hintBusy={t('modsHintBusy')}
              sendLabel={t('modsSend')}
              stopLabel={t('modsStop')}
              ariaLabel={t('modsPlaceholder')}
              onChange={setDraft}
              onSend={() => void ask()}
              onStop={() => {
                loopRef.current?.cancel()
                finishLast()
              }}
            />
          </div>
        </section>

        <aside className="mods-registry" aria-label={registryTitle}>
          <div className="mods-registry-head">
            <h2 className="mods-registry-title">{registryTitle}</h2>
            <div className="mods-registry-actions">
              <button className="mods-ghost-btn" onClick={() => void exportPack()}>
                {t('modsExport')}
              </button>
              <button className="mods-ghost-btn" onClick={() => void importPack()}>
                {t('modsImport')}
              </button>
            </div>
          </div>
          {pack ? (
            <section className="mods-pack" aria-label={t('modsImport')}>
              {pack.errors?.length ? (
                pack.errors.map((error, index) => (
                  <p key={index} className="mods-script-error">
                    {error}
                  </p>
                ))
              ) : (
                <>
                  <p className="mods-pack-count">
                    {pack.description?.count ?? 0} · {pack.description?.exportedAt.slice(0, 10)}
                  </p>
                  <ul className="mods-pack-lines">
                    {(pack.description?.lines ?? []).map((line) => (
                      <li key={line.id}>
                        <span className="mods-feature-name">{line.name}</span>
                        <span className="mods-feature-kind">{line.kind}</span>
                        {line.declares.map((text, index) => (
                          <p key={index} className="mods-hook-line">
                            {text}
                          </p>
                        ))}
                      </li>
                    ))}
                  </ul>
                  {(pack.description?.warnings ?? []).map((warning, index) => (
                    <p key={index} className="mods-script-error">
                      {warning}
                    </p>
                  ))}
                </>
              )}
              <div className="mods-script-actions">
                <button
                  className="mods-script-btn mods-script-btn--apply"
                  disabled={Boolean(pack.errors?.length)}
                  onClick={() => void installPack()}
                >
                  {t('modsInstall')}
                </button>
                <button className="mods-script-btn" onClick={() => setPack(null)}>
                  {t('modsCancel')}
                </button>
              </div>
            </section>
          ) : null}
          {features.length > 0 ? (
            <ul className="mods-registry-list">
              {features.map((feature) => (
                <li key={feature.id} className="mods-feature">
                  <div className="mods-feature-top">
                    <span className="mods-feature-name">{feature.name}</span>
                    {feature.slot || feature.kind ? (
                      <span className="mods-feature-kind">{feature.slot?.at ?? feature.kind}</span>
                    ) : null}
                  </div>
                  {feature.description ? (
                    <p className="mods-feature-desc">{feature.description}</p>
                  ) : null}
                  <code className="mods-feature-id">{feature.id}</code>
                </li>
              ))}
            </ul>
          ) : (
            <p className="mods-registry-empty">{t('modsNoFeatures')}</p>
          )}
          {scriptFeatures.length > 0 ? (
            <section className="mods-scripts" aria-label={t('modsScripts')}>
              <h3 className="mods-scripts-title">{t('modsScripts')}</h3>
              <ul className="mods-scripts-list">
                {scriptFeatures.map((feature) => {
                  const state = scriptRuns[feature.id]
                  const chosen = scriptPaths[feature.id]
                  const total = state?.outcome
                    ? state.outcome.applied + state.outcome.failures.length
                    : 0
                  return (
                    <li key={feature.id} className="mods-script">
                      <div className="mods-feature-top">
                        <span className="mods-feature-name">{feature.name}</span>
                        <span className="mods-feature-kind">
                          {feature.kind || 'script'}
                          {feature.script?.format ? ' · ' + feature.script.format : ''}
                        </span>
                      </div>
                      {feature.description ? (
                        <p className="mods-feature-desc">{feature.description}</p>
                      ) : null}
                      <select
                        className="mods-script-pick"
                        aria-label={t('open')}
                        value={chosen ?? ''}
                        onChange={(event) =>
                          setScriptPaths((prev) => ({ ...prev, [feature.id]: event.target.value }))
                        }
                      >
                        <option value="" disabled>
                          {t('open')}
                        </option>
                        {recents.map((entry) => (
                          <option key={entry.path} value={entry.path}>
                            {entry.name}
                          </option>
                        ))}
                      </select>
                      <div className="mods-script-actions">
                        <button
                          className="mods-script-btn"
                          disabled={!chosen || state?.running}
                          onClick={() => chosen && void dryRun(feature.id, chosen)}
                        >
                          {t('modsDryRun')}
                        </button>
                        <button
                          className="mods-script-btn mods-script-btn--apply"
                          disabled={!canApply(feature.id) || state?.running}
                          onClick={() => chosen && void applyRun(feature.id, chosen)}
                        >
                          {t('modsApply')}
                        </button>
                      </div>
                      {state?.error ? <p className="mods-script-error">{state.error}</p> : null}
                      {state?.outcome || state?.applied ? (
                        <div
                          className={
                            'mods-script-plan' + (state?.applied ? ' mods-script-plan--done' : '')
                          }
                        >
                          <p className="mods-script-count">
                            {(state.applied ?? state.outcome)!.applied}
                            {' / '}
                            {state.applied ? state.applied.applied : total}
                          </p>
                          <ul className="mods-script-lines">
                            {(state.applied?.plan ?? state.outcome?.plan ?? [])
                              .slice(0, 6)
                              .map((line, index) => (
                                <li key={index}>{line}</li>
                              ))}
                          </ul>
                          {(state.outcome?.result ?? state.applied?.result) ? (
                            <p className="mods-script-result">
                              {(state.outcome?.result ?? state.applied?.result)!.slice(0, 400)}
                            </p>
                          ) : null}
                          {(state.outcome?.failures ?? []).slice(0, 3).map((line, index) => (
                            <p key={index} className="mods-script-error">
                              {line}
                            </p>
                          ))}
                        </div>
                      ) : null}
                    </li>
                  )
                })}
              </ul>
            </section>
          ) : null}
          {knowledgeFeatures.length > 0 ? (
            <section className="mods-scripts" aria-label={t('modsSources')}>
              <h3 className="mods-scripts-title">{t('modsSources')}</h3>
              <ul className="mods-scripts-list">
                {knowledgeFeatures.map((feature) => {
                  const state = sourceStates[feature.id]
                  return (
                    <li key={feature.id} className="mods-script">
                      <div className="mods-feature-top">
                        <span className="mods-feature-name">
                          {feature.knowledge?.label ?? feature.name}
                        </span>
                        <span className="mods-feature-kind">
                          {feature.knowledge?.source ?? 'knowledge'}
                        </span>
                      </div>
                      {feature.knowledge ? (
                        <p className="mods-feature-desc">{lastSegment(feature.knowledge.path)}</p>
                      ) : null}
                      <p className="mods-script-count">
                        {state?.missing
                          ? t('modsSourceMissing')
                          : state
                            ? t('modsSourceDocs', { n: String(state.documents) })
                            : t('modsSourceRegistered')}
                      </p>
                      {onOpenAsk ? (
                        <div className="mods-script-actions">
                          <button className="mods-script-btn" onClick={() => onOpenAsk()}>
                            {t('modsOpenKnowledge')}
                          </button>
                        </div>
                      ) : null}
                    </li>
                  )
                })}
              </ul>
            </section>
          ) : null}
          {hooksSection()}
        </aside>
      </div>
    </div>
  )
}
