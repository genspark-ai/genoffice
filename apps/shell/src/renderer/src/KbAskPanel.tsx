import { useEffect, useRef, useState } from 'react'
import {
  AgentLoop,
  composeSkills,
  CREATE_DOCUMENT_TOOL,
  createDocumentSkill,
  createIpcTransport,
  createKnowledgeBaseSkill,
  type CreateDocumentFormat,
  type KbFileInfo,
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
import type { ChatMessage, ChatMeta, ProjectApi } from '@genoffice/project-store'
import { useI18n } from './locale'
import sendEnterOn from './assets/send-enter-on.png'
import sendEnterOff from './assets/send-enter-off.png'
import sendStop from './assets/send-stop.png'

/**
 * The home screen's knowledge-base ask panel: chat with the reader's starred
 * corpus without opening a document. First app-side consumer of the shell's
 * kb:* gateway — the loop, streaming and rendering reuse what every editor's
 * AI panel uses (AgentLoop + createIpcTransport + Markdown with a filenav
 * nav), so answers cite the folder and file they came from and clicking a
 * citation opens that file in the app that owns it.
 *
 * The conversation persists through project-store under a fixed chat id, the
 * one place in the app where a conversation belongs to no file: with no
 * `filePath` the store takes the caller's id verbatim, so the panel passes a
 * constant and finds its own transcript again on the next launch. An empty
 * corpus gets an onboarding hero instead of a bare hint: the only useful move
 * there is going to Recents and starring files, so the panel says so and
 * offers the jump as a button.
 *
 * That constant is also the limit of what exists so far: one transcript, no
 * list, no name, and nothing to start a second conversation. A reader with two
 * questions on two days gets one long thread — a list, per-conversation ids
 * and a name are the next layer, not this one.
 */

/** The one transcript the panel keeps. Stable across launches by design. */
const CHAT_ID = 'kb-ask'

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

/**
 * The panel's conversations, newest first, labelled by their own first question.
 *
 * A stored conversation has no title — the store writes `<chat-id>.jsonl` and
 * nothing else — so the first thing said in it is the only honest label. A
 * blank one falls back to the id, which at least is unique and copyable.
 *
 * The list is sorted here rather than trusting the store to: `listChats`
 * returns directory order, and a reader opening the panel expects the thread
 * they were just in to be on top.
 */
export interface ListedThread {
  readonly chatId: string
  readonly label: string
  readonly updatedAt: string
  readonly approxCount: number
}

export function listThreads(
  metas: readonly ChatMeta[],
  firstQuestionOf: (chatId: string) => string | undefined,
): ListedThread[] {
  return metas
    .map((meta) => ({
      chatId: meta.chatId,
      label: firstQuestionOf(meta.chatId)?.replace(/\s+/g, ' ').trim().slice(0, 60) || meta.chatId,
      updatedAt: meta.updatedAt,
      approxCount: meta.approxCount,
    }))
    .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt))
}

/** The first question in a stored conversation, which is how it is labelled. */
export function firstQuestionOf(stored: readonly ChatMessage[]): string | undefined {
  return stored.find((message) => message.role === 'user')?.text
}

/**
 * Claims the panel's transcript: resolves its ids, writes whatever turns were
 * asked before the ids landed, then reads back what is already stored.
 *
 * Split out of the effect so the ordering is testable without a DOM — the
 * panel is the only conversation in the app that belongs to no file, so
 * nothing else exercises the file-less `resolveChat` path end to end.
 */
export async function claimThread(
  api: ProjectApi,
  chatId: string,
  queued: Array<{ role: 'user' | 'assistant'; text: string }>,
  show: (messages: AskMessage[]) => void,
  stillCurrent: () => boolean,
): Promise<{ projectId: string; chatId: string }> {
  const ids = await api.resolveChat({ filePath: null, tempChatId: chatId })
  for (const entry of queued) {
    await api.appendChat({ ...ids, ...entry })
  }
  const stored = await api.loadChat({ ...ids, limit: 200 })
  if (stored.length > 0 && stillCurrent()) show(restoreAskMessages(stored))
  return ids
}

/**
 * The stored turns as the panel renders them. Only role and text come back:
 * a turn read from disk has this session's badges, token chip and "save as"
 * outcome long gone, and rendering them as absent is truer than guessing.
 */
export function restoreAskMessages(stored: ChatMessage[]): AskMessage[] {
  return stored
    .filter((message) => message.role === 'user' || message.role === 'assistant')
    .map((message) => ({ role: message.role, text: message.text, toolNotes: [] }))
}

export function KbAskPanel({
  onOpenPath,
  onGoToRecents,
}: {
  onOpenPath: (path: string) => void
  onGoToRecents: () => void
}): React.JSX.Element {
  const { t, lang } = useI18n()
  const [messages, setMessages] = useState<AskMessage[]>([])
  const [draft, setDraft] = useState('')
  const [busy, setBusy] = useState(false)
  const [corpusCount, setCorpusCount] = useState<number | null>(null)
  const [gskEmail, setGskEmail] = useState<string | null>(null)
  const corpusRef = useRef<KbFileInfo[]>([])
  const settingsRef = useRef<AiSettings | null>(null)
  const loopRef = useRef<AgentLoop<unknown> | null>(null)
  const scrollRef = useRef<HTMLDivElement>(null)
  /** bumped by Stop: an ask still refreshing the corpus must not start its run */
  const askSeqRef = useRef(0)
  /** tokens billed so far in the in-flight answer (summed across turns) */
  const runUsageRef = useRef<AiTokenUsageView | undefined>(undefined)
  /** tokens billed across the panel's session (the composer chip) */
  const [sessionUsage, setSessionUsage] = useState<AiTokenUsageView | undefined>(undefined)
  /** the rendered thread, readable from callbacks that close over an older one */
  const messagesRef = useRef<AskMessage[]>([])
  messagesRef.current = messages
  /** where this panel's transcript lives, once project-store has answered */
  const [activeChat, setActiveChat] = useState<string>(CHAT_ID)
  const [threads, setThreads] = useState<ListedThread[]>([])
  const chatIdsRef = useRef<{ projectId: string; chatId: string } | null>(null)
  /** turns asked before the ids came back, written once they do */
  const pendingRef = useRef<Array<{ role: 'user' | 'assistant'; text: string }>>([])

  const persist = (entry: { role: 'user' | 'assistant'; text: string }): void => {
    const ids = chatIdsRef.current
    if (!ids) {
      pendingRef.current.push(entry)
      return
    }
    void window.projectApi?.appendChat({ ...ids, ...entry }).catch(() => undefined)
  }

  useEffect(() => {
    // Reopening the panel should show the thread that was already there. The
    // ids resolve in the background, so a turn asked before they land is
    // queued by persist() and flushed here rather than lost.
    const api = window.projectApi
    if (!api) return
    let current = true
    void claimThread(api, CHAT_ID, pendingRef.current, setMessages, () => current)
      .then((ids) => {
        if (current) chatIdsRef.current = ids
      })
      .catch(() => undefined)
    void refreshThreads(api)
    return () => {
      current = false
    }
  }, [activeChat])

  const refreshThreads = async (api: ProjectApi): Promise<void> => {
    try {
      const ids = await api.resolveChat({ filePath: null, tempChatId: CHAT_ID })
      const metas = await api.listChats({ projectId: ids.projectId })
      // The label needs each thread's own first question, so every id is read
      // back; a chat with hundreds of turns still costs one short read each.
      const labelled = await Promise.all(
        metas.map(async (meta) => {
          let label: string | undefined
          try {
            label = firstQuestionOf(await api.loadChat({ ...ids, chatId: meta.chatId, limit: 20 }))
          } catch {
            label = undefined
          }
          return [meta, label] as const
        }),
      )
      setThreads(
        listThreads(
          labelled.map(([meta]) => meta),
          (chatId) => labelled.find(([meta]) => meta.chatId === chatId)?.[1],
        ),
      )
    } catch {
      setThreads([])
    }
  }

  /** Switching is a state change: the effect re-reads whichever thread is now active. */
  const openThread = (chatId: string): void => {
    setActiveChat(chatId)
    setMessages([])
  }

  /** A fresh id, so a new thread does not append to the one being read. */
  const startThread = (): void => {
    setActiveChat(`kb-ask-${Date.now()}`)
    setMessages([])
  }

  const removeThread = async (chatId: string): Promise<void> => {
    const api = window.projectApi
    if (!api) return
    try {
      const ids = await api.resolveChat({ filePath: null, tempChatId: CHAT_ID })
      await api.deleteChat({ projectId: ids.projectId, chatId })
      if (chatId === activeChat) startThread()
      await refreshThreads(api)
    } catch {
      // A delete that did not happen leaves the list as it was, which is the
      // honest outcome; the reader sees the thread still listed.
    }
  }

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
    // The answer is whole for the first time here; streaming it as it grows
    // would store every intermediate state. The `streaming` guard is what
    // keeps a second finish from storing the same turn twice.
    const last = messagesRef.current[messagesRef.current.length - 1]
    if (last?.role === 'assistant' && last.streaming && last.text) {
      persist({ role: 'assistant', text: last.text })
    }
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
    // freshly starred files are answerable this very turn
    try {
      const files = await window.aiOffice.kbList()
      corpusRef.current = files
      setCorpusCount(files.length)
    } catch {
      // keep the mount snapshot
    }
    if (askSeqRef.current !== seq) return
    if (corpusRef.current.length === 0) {
      setMessages((prev) => [
        ...prev,
        { role: 'user', text: instruction, toolNotes: [] },
        { role: 'assistant', text: t('askEmptyCorpus'), toolNotes: [] },
      ])
      persist({ role: 'user', text: instruction })
      persist({ role: 'assistant', text: t('askEmptyCorpus') })
      setBusy(false)
      return
    }
    setMessages((prev) => [
      ...prev,
      { role: 'user', text: instruction, toolNotes: [] },
      { role: 'assistant', text: '', toolNotes: [], streaming: true },
    ])
    persist({ role: 'user', text: instruction })
    try {
      await loopRef.current!.run(instruction)
    } catch {
      finishLast()
    }
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
      {threads.length > 0 && (
        <div className="kb-ask-threads" role="tablist" aria-label={t('askThreadsTitle')}>
          {threads.map((thread) => (
            <div
              key={thread.chatId}
              className={`kb-ask-thread-chip${thread.chatId === activeChat ? ' active' : ''}`}
            >
              <button
                type="button"
                role="tab"
                aria-selected={thread.chatId === activeChat}
                title={thread.label}
                onClick={() => openThread(thread.chatId)}
              >
                {thread.label}
              </button>
              <button
                type="button"
                className="kb-ask-thread-del"
                aria-label={t('askThreadDelete', { label: thread.label })}
                onClick={() => void removeThread(thread.chatId)}
              >
                ×
              </button>
            </div>
          ))}
          <button
            type="button"
            className="kb-ask-thread-new"
            aria-label={t('askThreadNew')}
            title={t('askThreadNew')}
            onClick={startThread}
          >
            +
          </button>
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
