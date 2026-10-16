import { useEffect, useRef, useState } from 'react'
import {
  AgentLoop,
  composeSkills,
  CREATE_DOCUMENT_TOOL,
  createDocumentSkill,
  createIpcTransport,
  createKnowledgeBaseSkill,
  type CreateDocumentFormat,
  formatKbSize,
  type KbFileInfo,
  type KbSourceInfo,
} from '@genoffice/agent-core'
import {
  AiComposer,
  AiTypingIndicator,
  AiUsageChip,
  addTokenUsage,
  createFileNav,
  Markdown,
  type AiTokenUsageView,
} from '@genoffice/ui'
import type { AiSettings } from '@genoffice/ai-provider'
import { useI18n } from './locale'
import sendEnterOn from './assets/send-enter-on.png'
import sendEnterOff from './assets/send-enter-off.png'
import sendStop from './assets/send-stop.png'

/**
 * The home screen's knowledge-base ask panel: chat with the corpus of sources —
 * the reader's starred files plus the folders and files registered by knowledge
 * features — without opening a document. The Sources section shows what the
 * corpus is made of and toggles a registered source on or off (the corpus, the
 * controls and the index live in the shell main; see knowledge-base.ts). The
 * loop, streaming and rendering reuse what every editor's AI panel uses
 * (AgentLoop + createIpcTransport + Markdown with a filenav nav), so answers
 * cite the folder and file they came from and clicking a citation opens that
 * file in the app that owns it.
 *
 * The conversation lives in memory only; it does not survive a restart (the
 * editor panels persist through project-store — wiring that here is a
 * deliberate later step, not an oversight). An empty corpus gets an
 * onboarding hero instead of a bare hint: the only useful move there is
 * going to Recents and starring files, so the panel says so and offers the
 * jump as a button.
 */

/** one tool a turn used, badge-shaped; the icon follows the call name */
interface ToolNote {
  label: string
  kind: 'search' | 'read' | 'create'
}

interface AskMessage {
  role: 'user' | 'assistant'
  text: string
  toolNotes: ToolNote[]
  streaming?: boolean
  isError?: boolean
  /** tokens the provider billed for this answer, summed across its turns (absent when unreported) */
  usage?: AiTokenUsageView
  /** the "save as" buttons' outcome for this answer */
  saved?: { path: string; name: string } | { error: string } | 'saving'
}

export function KbAskPanel({
  onOpenPath,
  onOpenMods,
  onGoToRecents,
}: {
  onOpenPath: (path: string) => void
  /** open the feature assistant: that is where a source is added (a knowledge feature) */
  onOpenMods?: () => void
  onGoToRecents: () => void
}): React.JSX.Element {
  const { t, lang } = useI18n()
  const [messages, setMessages] = useState<AskMessage[]>([])
  const [draft, setDraft] = useState('')
  const [busy, setBusy] = useState(false)
  const [corpusCount, setCorpusCount] = useState<number | null>(null)
  const [sources, setSources] = useState<KbSourceInfo[] | null>(null)
  const [sourcesOpen, setSourcesOpen] = useState(false)
  const [reindexing, setReindexing] = useState(false)
  const [gskEmail, setGskEmail] = useState<string | null>(null)
  const corpusRef = useRef<KbFileInfo[]>([])
  const sourcesRef = useRef<KbSourceInfo[]>([])
  const settingsRef = useRef<AiSettings | null>(null)
  const loopRef = useRef<AgentLoop<unknown> | null>(null)
  const scrollRef = useRef<HTMLDivElement>(null)
  /** bumped by Stop: an ask still refreshing the corpus must not start its run */
  const askSeqRef = useRef(0)
  /** tokens billed so far in the in-flight answer (summed across turns) */
  const runUsageRef = useRef<AiTokenUsageView | undefined>(undefined)
  /** tokens billed across the panel's session (the composer chip) */
  const [sessionUsage, setSessionUsage] = useState<AiTokenUsageView | undefined>(undefined)

  const refreshSettings = (): void => {
    void window.aiOffice
      .aiGetSettings()
      .then((settings) => {
        settingsRef.current = settings
      })
      .catch(() => undefined)
  }

  const refreshCorpus = (): void => {
    void window.aiOffice
      .kbList()
      .then((files) => {
        corpusRef.current = files
        setCorpusCount(files.length)
      })
      .catch(() => setCorpusCount(0))
  }

  const refreshSources = (): void => {
    if (typeof window.aiOffice?.kbSources !== 'function') return
    void window.aiOffice
      .kbSources()
      .then((next) => {
        sourcesRef.current = next
        setSources(next)
      })
      .catch(() => undefined)
  }

  useEffect(() => {
    // one corpus snapshot per mount; sends refresh it again (see ask()).
    // A bridge without the kb channel (defensive: the shell always exposes
    // it) leaves the corpus empty, which keeps the panel on its guidance.
    if (typeof window.aiOffice?.kbList !== 'function') {
      setCorpusCount(0)
      return
    }
    refreshCorpus()
    refreshSources()
    refreshSettings()
    void window.aiOffice
      .aiGskStatus()
      .then((status) => setGskEmail(status.loggedIn ? (status.email ?? '') : null))
      .catch(() => undefined)
    const docSkill = createDocumentSkill({
      create: (request) => window.aiOffice.kbCreateDocument(request),
    })
    loopRef.current = new AgentLoop({
      transport: createIpcTransport<AiSettings>({
        onStream: (listener) => window.aiOffice.onAiStream(listener),
        start: (request) => window.aiOffice.aiStream(request),
        cancel: (requestId) => void window.aiOffice.aiStreamCancel(requestId),
        getSettings: () => settingsRef.current as AiSettings,
        unknownErrorText: () => t('askUnknownError'),
      }),
      skill: composeSkills('home-ask', '', [
        createKnowledgeBaseSkill({
          listFiles: () => corpusRef.current,
          sources: () => sourcesRef.current,
          search: (query, limit) => window.aiOffice.kbSearch({ q: query, limit }),
          read: (path, offset) => window.aiOffice.kbRead(path, offset),
          alsoCitable: () => docSkill.createdPaths(),
        }),
        docSkill,
      ]),
      events: {
        onText: (text) => patchLast({ text }),
        onToolStart: (call) =>
          patchLast((last) => ({
            toolNotes: [
              ...last.toolNotes,
              call.name === 'search_knowledge_base'
                ? { label: t('askToolSearch'), kind: 'search' as const }
                : call.name === CREATE_DOCUMENT_TOOL
                  ? { label: t('askToolCreate'), kind: 'create' as const }
                  : { label: t('askToolRead'), kind: 'read' as const },
            ],
          })),
        onDone: () => finishLast(),
        onUsage: (usage) => {
          runUsageRef.current = addTokenUsage(runUsageRef.current, usage)
          setSessionUsage((prev) => addTokenUsage(prev, usage))
          patchLast({ usage: runUsageRef.current })
        },
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
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight })
  }, [messages])

  const patchLast = (
    patch: Partial<AskMessage> | ((last: AskMessage) => Partial<AskMessage>),
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
  }

  const ask = async (): Promise<void> => {
    const instruction = draft.trim()
    if (!instruction || busy) return
    setDraft('')
    setBusy(true)
    refreshSettings()
    const seq = ++askSeqRef.current
    runUsageRef.current = undefined
    // freshly starred files and freshly registered sources are answerable this turn
    try {
      const files = await window.aiOffice.kbList()
      corpusRef.current = files
      setCorpusCount(files.length)
    } catch {
      // keep the mount snapshot
    }
    refreshSources()
    if (askSeqRef.current !== seq) return
    if (corpusRef.current.length === 0) {
      setMessages((prev) => [
        ...prev,
        { role: 'user', text: instruction, toolNotes: [] },
        { role: 'assistant', text: t('askEmptyCorpus'), toolNotes: [] },
      ])
      setBusy(false)
      return
    }
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

  const setSourceEnabled = (id: string, enabled: boolean): void => {
    void window.aiOffice
      .kbSetSourceEnabled(id, enabled)
      .then((next) => {
        sourcesRef.current = next
        setSources(next)
        refreshCorpus()
      })
      .catch(() => undefined)
  }

  const reindex = (): void => {
    setReindexing(true)
    void window.aiOffice
      .kbReindex()
      .catch(() => undefined)
      .then(() => {
        // the scan is debounced in the main process; a beat later the counts move
        window.setTimeout(() => {
          setReindexing(false)
          refreshSources()
          refreshCorpus()
        }, 1500)
      })
  }

  const fileNav = createFileNav((path) => onOpenPath(path))

  // the deterministic half of "keep this answer": the question becomes the
  // file name, the answer's markdown the body, same write path as the tool
  const saveAnswer = async (index: number, format: CreateDocumentFormat): Promise<void> => {
    const answer = messages[index]
    if (!answer || answer.saved === 'saving') return
    const question = [...messages.slice(0, index)].reverse().find((m) => m.role === 'user')
    const title = (question?.text ?? '').replace(/\s+/g, ' ').trim().slice(0, 60)
    const patchAt = (saved: AskMessage['saved']): void =>
      setMessages((prev) => prev.map((m, i) => (i === index ? { ...m, saved } : m)))
    patchAt('saving')
    try {
      const { path } = await window.aiOffice.kbCreateDocument({
        title,
        format,
        content: answer.text,
      })
      patchAt({ path, name: path.split(/[\\/]/).pop() ?? path })
    } catch (error) {
      patchAt({ error: error instanceof Error ? error.message : String(error) })
    }
  }

  return (
    <div className="kb-ask">
      {corpusCount !== null && corpusCount > 0 && (
        <div className="kb-ask-headbar">
          <b>{t('askTitle')}</b>
          <span>{t('askCorpusCount', { n: String(corpusCount) })}</span>
        </div>
      )}
      {sources !== null && sources.length > 0 ? (
        <div className="kb-ask-sources">
          <button
            type="button"
            className="kb-ask-sources-toggle"
            onClick={() => setSourcesOpen((open) => !open)}
          >
            {t('askSources')} ({sources.length})
            <span className="kb-ask-sources-caret">{sourcesOpen ? '▾' : '▸'}</span>
          </button>
          {sourcesOpen ? (
            <>
              <ul className="kb-ask-sources-list">
                {sources.map((source) => (
                  <li key={source.id} className="kb-ask-source">
                    <span className="kb-ask-source-kind">{source.kind}</span>
                    <span className="kb-ask-source-label">{source.label}</span>
                    <span className="kb-ask-source-count">
                      {source.missing
                        ? t('askSourceMissing')
                        : t('askSourceDocs', {
                            n: String(source.documents),
                            size: formatKbSize(source.bytes ?? 0),
                          })}
                    </span>
                    {source.missing ? null : source.id === 'starred' ? null : (
                      <button
                        type="button"
                        className={'kb-ask-source-toggle' + (source.enabled ? '' : ' is-off')}
                        onClick={() => setSourceEnabled(source.id, !source.enabled)}
                      >
                        {source.enabled ? t('askSourceEnabled') : t('askSourceDisabled')}
                      </button>
                    )}
                  </li>
                ))}
              </ul>
              <div className="kb-ask-sources-actions">
                <button type="button" className="kb-ask-source-toggle" onClick={reindex}>
                  {reindexing ? t('askReindexing') : t('askReindex')}
                </button>
                {onOpenMods ? (
                  <button
                    type="button"
                    className="kb-ask-source-toggle"
                    onClick={() => onOpenMods()}
                  >
                    {t('askAddSource')}
                  </button>
                ) : null}
              </div>
            </>
          ) : null}
        </div>
      ) : null}
      <div className="kb-ask-thread" ref={scrollRef}>
        {messages.length === 0 && corpusCount === 0 && (
          <div className="kb-ask-hero">
            <div className="kb-ask-hero-icon" aria-hidden="true">
              <IconBookOpen />
            </div>
            <h3>{t('askTitle')}</h3>
            <p>{t('askHeroTagline')}</p>
            <div className="kb-ask-steps">
              <div className="kb-ask-step">
                <span className="kb-ask-step-num">1</span>
                <div>
                  <b>{t('askStep1Title')}</b>
                  <span>{t('askStep1Desc')}</span>
                </div>
              </div>
              <div className="kb-ask-step">
                <span className="kb-ask-step-num">2</span>
                <div>
                  <b>{t('askStep2Title')}</b>
                  <span>{t('askStep2Desc')}</span>
                </div>
              </div>
              <div className="kb-ask-step">
                <span className="kb-ask-step-num">3</span>
                <div>
                  <b>{t('askStep3Title')}</b>
                  <span>{t('askStep3Desc')}</span>
                </div>
              </div>
            </div>
            <button className="kb-ask-cta" type="button" onClick={onGoToRecents}>
              <IconStar />
              {t('askCtaRecents')}
            </button>
          </div>
        )}
        {messages.length === 0 && corpusCount !== null && corpusCount > 0 && (
          <p className="kb-ask-hint">{t('askIntro')}</p>
        )}
        {messages.map((m, i) => (
          <div key={i} className={`kb-ask-msg kb-ask-${m.role}${m.isError ? ' kb-ask-error' : ''}`}>
            {m.role === 'assistant' && m.toolNotes.length > 0 && (
              <div className="kb-ask-tools">
                {m.toolNotes.map((note, j) => (
                  <span key={j} className="kb-ask-tool-note">
                    {note.kind === 'search' ? (
                      <IconSearch />
                    ) : note.kind === 'create' ? (
                      <IconFilePlus />
                    ) : (
                      <IconFileText />
                    )}
                    {note.label}
                    {m.streaming && j === m.toolNotes.length - 1 ? '…' : ''}
                  </span>
                ))}
              </div>
            )}
            {m.text ? (
              m.role === 'assistant' ? (
                <Markdown text={m.text} navs={[fileNav]} />
              ) : (
                m.text
              )
            ) : m.streaming ? (
              <AiTypingIndicator label={t('askHintBusy')} />
            ) : null}
            {m.role === 'assistant' && m.usage && (
              <div className="ai-message-usage">
                <AiUsageChip usage={m.usage} lang={lang} />
              </div>
            )}
            {m.role === 'assistant' && m.text && !m.streaming && !m.isError && (
              <div className="kb-ask-actions">
                {m.saved && typeof m.saved === 'object' && 'path' in m.saved ? (
                  <button
                    type="button"
                    className="kb-ask-action kb-ask-action-saved"
                    onClick={() => onOpenPath((m.saved as { path: string }).path)}
                  >
                    <IconFileText />
                    {t('askSavedAs', { name: m.saved.name })}
                  </button>
                ) : (
                  <>
                    <button
                      type="button"
                      className="kb-ask-action"
                      disabled={m.saved === 'saving'}
                      onClick={() => void saveAnswer(i, 'md')}
                    >
                      <IconFilePlus />
                      {t('askSaveMd')}
                    </button>
                    <button
                      type="button"
                      className="kb-ask-action"
                      disabled={m.saved === 'saving'}
                      onClick={() => void saveAnswer(i, 'docx')}
                    >
                      <IconFilePlus />
                      {t('askSaveDocx')}
                    </button>
                    {m.saved && typeof m.saved === 'object' && 'error' in m.saved && (
                      <span className="kb-ask-action-error">
                        {t('askSaveFailed', { error: m.saved.error })}
                      </span>
                    )}
                  </>
                )}
              </div>
            )}
          </div>
        ))}
      </div>
      {messages.length > 0 && (
        <p className="kb-ask-nosave">
          <IconInfo />
          {t('askNoSave')}
        </p>
      )}
      {gskEmail === null && <p className="kb-ask-login">{t('askNeedsLogin')}</p>}
      <AiComposer
        value={draft}
        busy={busy}
        placeholder={t('askPlaceholder')}
        hintIdle={t('askHintIdle')}
        hintBusy={t('askHintBusy')}
        sendLabel={t('askSend')}
        stopLabel={t('askStop')}
        ariaLabel={t('askPlaceholder')}
        iconOnly
        sendIconEnabled={<img src={sendEnterOn} alt="" aria-hidden />}
        sendIconDisabled={<img src={sendEnterOff} alt="" aria-hidden />}
        stopIcon={<img src={sendStop} alt="" aria-hidden />}
        footerStart={
          sessionUsage ? <AiUsageChip usage={sessionUsage} session lang={lang} /> : undefined
        }
        onChange={setDraft}
        onSend={() => void ask()}
        onStop={() => {
          askSeqRef.current++
          loopRef.current?.cancel()
          finishLast()
        }}
      />
    </div>
  )
}

/** open book for the empty-corpus hero (inline, like every Home icon) */
function IconBookOpen(): React.JSX.Element {
  return (
    <svg
      width="26"
      height="26"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M12 6.8C10.2 5.4 7.9 4.7 5.5 4.7c-.4 0-.8 0-1.2.1v12.9c.4-.1.8-.1 1.2-.1 2.4 0 4.7.7 6.5 2.1 1.8-1.4 4.1-2.1 6.5-2.1.4 0 .8 0 1.2.1V4.8c-.4-.1-.8-.1-1.2-.1-2.4 0-4.7.7-6.5 2.1z" />
      <path d="M12 6.8v12.9" />
    </svg>
  )
}

function IconStar(): React.JSX.Element {
  return (
    <svg
      width="15"
      height="15"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M12 3.6l2.6 5.2 5.8.9-4.2 4.1 1 5.8-5.2-2.7-5.2 2.7 1-5.8L3.6 9.7l5.8-.9z" />
    </svg>
  )
}

function IconInfo(): React.JSX.Element {
  return (
    <svg
      width="13"
      height="13"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      aria-hidden="true"
    >
      <circle cx="12" cy="12" r="9" />
      <path d="M12 8.2v.01M12 11.5V16" />
    </svg>
  )
}

function IconSearch(): React.JSX.Element {
  return (
    <svg
      width="11"
      height="11"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.2"
      strokeLinecap="round"
      aria-hidden="true"
    >
      <circle cx="11" cy="11" r="7" />
      <path d="m20 20-3.8-3.8" />
    </svg>
  )
}

function IconFilePlus(): React.JSX.Element {
  return (
    <svg
      width="11"
      height="11"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.2"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <polyline points="14 2 14 8 20 8" />
      <line x1="12" y1="18" x2="12" y2="12" />
      <line x1="9" y1="15" x2="15" y2="15" />
    </svg>
  )
}

function IconFileText(): React.JSX.Element {
  return (
    <svg
      width="11"
      height="11"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.2"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z" />
      <path d="M14 3v5h5M9 13h6M9 17h6" />
    </svg>
  )
}
