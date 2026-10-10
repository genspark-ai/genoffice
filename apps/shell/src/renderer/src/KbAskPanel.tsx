import { useEffect, useRef, useState } from 'react'
import {
  AgentLoop,
  createIpcTransport,
  createKnowledgeBaseSkill,
  type KbFileInfo,
} from '@genoffice/agent-core'
import { AiComposer, AiTypingIndicator, createFileNav, Markdown } from '@genoffice/ui'
import type { AiSettings } from '@genoffice/ai-provider'
import { useI18n } from './locale'

/**
 * The home screen's knowledge-base ask panel: chat with the reader's starred
 * corpus without opening a document. First app-side consumer of the shell's
 * kb:* gateway — the loop, streaming and rendering reuse what every editor's
 * AI panel uses (AgentLoop + createIpcTransport + Markdown with a filenav
 * nav), so answers cite the folder and file they came from and clicking a
 * citation opens that file in the app that owns it.
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
  kind: 'search' | 'read'
}

interface AskMessage {
  role: 'user' | 'assistant'
  text: string
  toolNotes: ToolNote[]
  streaming?: boolean
  isError?: boolean
}

export function KbAskPanel({
  onOpenPath,
  onGoToRecents,
}: {
  onOpenPath: (path: string) => void
  onGoToRecents: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [messages, setMessages] = useState<AskMessage[]>([])
  const [draft, setDraft] = useState('')
  const [busy, setBusy] = useState(false)
  const [corpusCount, setCorpusCount] = useState<number | null>(null)
  const [gskEmail, setGskEmail] = useState<string | null>(null)
  const corpusRef = useRef<KbFileInfo[]>([])
  const settingsRef = useRef<AiSettings | null>(null)
  const loopRef = useRef<AgentLoop<unknown> | null>(null)
  const scrollRef = useRef<HTMLDivElement>(null)

  const refreshSettings = (): void => {
    void window.aiOffice
      .aiGetSettings()
      .then((settings) => {
        settingsRef.current = settings
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
    void window.aiOffice
      .kbList()
      .then((files) => {
        corpusRef.current = files
        setCorpusCount(files.length)
      })
      .catch(() => setCorpusCount(0))
    refreshSettings()
    void window.aiOffice
      .aiGskStatus()
      .then((status) => setGskEmail(status.loggedIn ? (status.email ?? '') : null))
      .catch(() => undefined)
    loopRef.current = new AgentLoop({
      transport: createIpcTransport<AiSettings>({
        onStream: (listener) => window.aiOffice.onAiStream(listener),
        start: (request) => window.aiOffice.aiStream(request),
        cancel: (requestId) => void window.aiOffice.aiStreamCancel(requestId),
        getSettings: () => settingsRef.current as AiSettings,
        unknownErrorText: () => t('askUnknownError'),
      }),
      skill: createKnowledgeBaseSkill({
        listFiles: () => corpusRef.current,
        search: (query, limit) => window.aiOffice.kbSearch({ q: query, limit }),
        read: (path, offset) => window.aiOffice.kbRead(path, offset),
      }),
      events: {
        onText: (text) => patchLast({ text }),
        onToolStart: (call) =>
          patchLast((last) => ({
            toolNotes: [
              ...last.toolNotes,
              call.name === 'search_knowledge_base'
                ? { label: t('askToolSearch'), kind: 'search' as const }
                : { label: t('askToolRead'), kind: 'read' as const },
            ],
          })),
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
    // freshly starred files are answerable this very turn
    void window.aiOffice
      .kbList()
      .then((files) => {
        corpusRef.current = files
        setCorpusCount(files.length)
      })
      .catch(() => undefined)
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

  const fileNav = createFileNav((path) => onOpenPath(path))

  return (
    <div className="kb-ask">
      {corpusCount !== null && corpusCount > 0 && (
        <div className="kb-ask-headbar">
          <b>{t('askTitle')}</b>
          <span>{t('askCorpusCount', { n: String(corpusCount) })}</span>
        </div>
      )}
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
                    {note.kind === 'search' ? <IconSearch /> : <IconFileText />}
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
        iconOnly
        sendIconEnabled={<IconArrowUp />}
        placeholder={t('askPlaceholder')}
        hintIdle={t('askHintIdle')}
        hintBusy={t('askHintBusy')}
        footerStart={
          <span className="ai-input-hint kb-ask-kbdhint">
            <kbd>Enter</kbd>
            {t('askSend')}
          </span>
        }
        sendLabel={t('askSend')}
        stopLabel={t('askStop')}
        ariaLabel={t('askPlaceholder')}
        onChange={setDraft}
        onSend={() => void ask()}
        onStop={() => {
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

/** upward arrow for the icon-only send button (IconEnter is a return key) */
function IconArrowUp(): React.JSX.Element {
  return (
    <svg
      width="16"
      height="16"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.2"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M12 19V5" />
      <path d="M5.5 11.5 12 5l6.5 6.5" />
    </svg>
  )
}
