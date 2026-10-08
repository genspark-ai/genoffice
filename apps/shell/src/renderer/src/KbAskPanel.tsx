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
 * deliberate later step, not an oversight).
 */

interface AskMessage {
  role: 'user' | 'assistant'
  text: string
  /** short labels of the tools this answer used, in order */
  toolNotes: string[]
  streaming?: boolean
  isError?: boolean
}

export function KbAskPanel({
  onOpenPath,
}: {
  onOpenPath: (path: string) => void
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
    // one corpus snapshot per mount; sends refresh it again (see ask())
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
              call.name === 'search_knowledge_base' ? t('askToolSearch') : t('askToolRead'),
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
      <div className="kb-ask-head">
        <h2>{t('askTitle')}</h2>
        <p className="kb-ask-sub">
          {corpusCount === null
            ? ''
            : corpusCount === 0
              ? t('askEmptyCorpusHint')
              : t('askCorpusCount', { n: String(corpusCount) })}
        </p>
      </div>
      <div className="kb-ask-thread" ref={scrollRef}>
        {messages.length === 0 && corpusCount !== null && corpusCount > 0 && (
          <p className="kb-ask-hint">{t('askIntro')}</p>
        )}
        {messages.map((m, i) => (
          <div key={i} className={`kb-ask-msg kb-ask-${m.role}${m.isError ? ' kb-ask-error' : ''}`}>
            {m.role === 'assistant' && m.toolNotes.length > 0 && (
              <div className="kb-ask-tools">
                {m.toolNotes.map((note, j) => (
                  <span key={j} className="kb-ask-tool-note">
                    {note}
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
        {busy && messages[messages.length - 1]?.text === '' && messages.length > 0 && null}
      </div>
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
