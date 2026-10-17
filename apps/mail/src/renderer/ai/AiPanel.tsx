import { useEffect, useMemo, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import { AgentLoop, composeSkills, createIpcTransport } from '@genoffice/agent-core'
import type { AgentSkill, AgentToolCall, AgentTransport } from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import { AiComposer, AiTypingIndicator } from '@genoffice/ui'
import { useI18n } from '../i18n/locale'
import type { AiScope } from '../devFakes'
import './ai.css'

/**
 * Mail AI panel (UI lane owns). Mirrors the markdown AiPanel wiring:
 * `AgentLoop` over a `createIpcTransport` shaped like markdown
 * `transport.ts:6` (over `window.mailApi.onAiStream/aiStream/
 * aiStreamCancel`), skills composed with `composeSkills('mail-ai', ...)`.
 *
 * Account scoping: the picker-selected account set travels with EVERY AI
 * call — skills read it live through `getScope`, and a scope change
 * cancels the in-flight run so stale-scope results never render. Thread
 * citations (`[uid]`) in replies open the reader through
 * `onOpenCitation`; the human sends through `onInsertAssist`, which lands
 * EDITABLE text in the compose box (AI proposes, human sends).
 *
 * The four mail skills (K lane) arrive through `createSkills`; until that
 * lane merges the panel streams tool-less text over the same transport.
 * The panel never logs bodies, snippets, subjects, or addresses.
 */

export interface AiCitationRef {
  uid: string
  folderId?: string
}

/** Live readers handed to the K-lane skill factories. */
export interface AiPanelSkillContext {
  /** Picker-selected scope, read fresh on every tool call. */
  getScope: () => AiScope
  /** Current compose draft text (user-typed), when App wires it. */
  getDraft: () => string
}

export interface MailAiPanelProps {
  /** Picker default; the panel follows later selection changes. */
  accountIds: string[]
  accounts: Array<{ accountId: string; label: string }>
  onOpenCitation(citation: AiCitationRef): void
  /** Assist text lands EDITABLE in compose; the human sends. */
  onInsertAssist(text: string): void
  /** K-lane wiring point: factories for the four mail skills. */
  createSkills?: (ctx: AiPanelSkillContext) => AgentSkill[]
  /** Test seam; defaults to the mail IPC transport. */
  transport?: AgentTransport
  /** Test/dev seam for the compose-assist draft context. */
  getDraft?: () => string
}

interface ToolActivity {
  name: string
  summary: string
  running?: boolean
  isError?: boolean
}

interface ChatEntry {
  role: 'user' | 'assistant'
  text: string
  streaming?: boolean
  isError?: boolean
  tools?: ToolActivity[]
}

/** Short `[uid]` citations; longer bracketed prose stays plain text. */
const CITATION_RE = /\[([A-Za-z0-9][A-Za-z0-9._-]*)\]/g

const MAIL_AI_INTRO = [
  'Mail assistant over the account-scoped retrieval bridge.',
  'Triage reads headers and snippets only; bodies are fetched for the cited',
  'set only. Every factual claim carries its per-claim [uid] citation.',
].join('\n')

function createMailTransport(
  getSettings: () => AiSettings,
  unknownErrorText: string,
): AgentTransport {
  return createIpcTransport<AiSettings>({
    onStream: (listener) => window.mailApi.onAiStream(listener),
    start: (request) => window.mailApi.aiStream(request),
    cancel: (requestId) => void window.mailApi.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => unknownErrorText,
  })
}

export function AiPanel({
  accountIds,
  accounts,
  onOpenCitation,
  onInsertAssist,
  createSkills,
  transport,
  getDraft,
}: MailAiPanelProps): ReactElement {
  const { t } = useI18n()
  const [picked, setPicked] = useState<string[]>(accountIds)
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [prompt, setPrompt] = useState('')
  const [busy, setBusy] = useState(false)
  const chatRef = useRef<HTMLDivElement>(null)
  const mountedRef = useRef(true)

  // Follow later folder/account selection; the picker stays user-editable.
  const accountKey = accountIds.slice().sort().join(',')
  useEffect(() => {
    setPicked(accountIds)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [accountKey])

  // Live readers so skills always see the current picker, never a capture.
  const pickedRef = useRef(picked)
  pickedRef.current = picked
  const getDraftRef = useRef(getDraft)
  getDraftRef.current = getDraft
  const skillContext = useMemo<AiPanelSkillContext>(
    () => ({
      getScope: (): AiScope => ({ accountIds: [...pickedRef.current] }),
      getDraft: (): string => getDraftRef.current?.() ?? '',
    }),
    [],
  )

  const errorText = t('mail.ai.error')
  const emptyText = t('mail.ai.emptyResult')
  const settingsRef = useRef<AiSettings | null>(null)
  const transportOrDefault = useMemo(
    () => transport ?? createMailTransport(() => settingsRef.current!, errorText),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [transport],
  )

  const skills = useMemo(() => createSkills?.(skillContext) ?? [], [createSkills, skillContext])

  const loopRef = useRef<AgentLoop<unknown> | null>(null)
  if (!loopRef.current) {
    loopRef.current = new AgentLoop<unknown>({
      transport: transportOrDefault,
      skill: composeSkills('mail-ai', MAIL_AI_INTRO, skills),
      events: {
        onText: (text) => patchLast({ text }),
        onToolStart: (call: AgentToolCall) => {
          patchLast((last) => ({
            tools: [
              ...(last.tools ?? []),
              { name: call.name, summary: call.name.replace(/[_-]+/g, ' '), running: true },
            ],
          }))
        },
        onToolExecuted: ({ call, execution }) => {
          patchLast((last) => {
            const tools = [...(last.tools ?? [])]
            if (tools.at(-1)?.running) tools.pop()
            return {
              tools: [
                ...tools,
                { name: call.name, summary: execution.summary, isError: execution.isError },
              ],
            }
          })
        },
        onDone: ({ text, cancelled }) => {
          const scopeNow = pickedRef.current.slice().sort().join(',')
          patchLast((last) => ({
            streaming: false,
            text:
              scopeNow !== sendScopeRef.current ? emptyText : text || (cancelled ? '' : emptyText),
            tools: last.tools?.filter((tool) => !tool.running),
          }))
          setBusy(false)
        },
        onError: (error) => {
          patchLast({ streaming: false, text: error, isError: true })
          setBusy(false)
        },
      },
    })
  }

  const sendScopeRef = useRef('')
  const runInstructionRef = useRef('')

  useEffect(
    () => () => {
      mountedRef.current = false
      loopRef.current?.cancel()
    },
    [],
  )

  // A scope change cancels the in-flight run: results from the old scope
  // must never render under the new picker.
  const scopeKey = picked.slice().sort().join(',')
  const busyRef = useRef(busy)
  busyRef.current = busy
  useEffect(() => {
    if (busyRef.current) loopRef.current?.cancel()
  }, [scopeKey])

  useEffect(() => {
    const el = chatRef.current
    // jsdom has no layout engine: scrollTo is absent there (and on
    // detached nodes), so the stick-to-bottom nudge must tolerate it.
    el?.scrollTo?.({ top: el.scrollHeight })
  }, [chat, busy])

  function patchLast(patch: Partial<ChatEntry> | ((last: ChatEntry) => Partial<ChatEntry>)): void {
    setChat((prev) => {
      const next = [...prev]
      const last = next[next.length - 1]
      if (!last || last.role !== 'assistant') return prev
      next[next.length - 1] = { ...last, ...(typeof patch === 'function' ? patch(last) : patch) }
      return next
    })
  }

  function send(text: string): void {
    const instruction = text.trim()
    const loop = loopRef.current
    if (!instruction || !loop || loop.busy) return
    if (pickedRef.current.length === 0) {
      setChat((prev) => [...prev, { role: 'user', text: instruction }])
      setChat((prev) => [...prev, { role: 'assistant', text: errorText, isError: true }])
      setPrompt('')
      return
    }
    sendScopeRef.current = pickedRef.current.slice().sort().join(',')
    runInstructionRef.current = instruction
    setChat((prev) => [
      ...prev,
      { role: 'user', text: instruction },
      { role: 'assistant', text: '', streaming: true },
    ])
    setPrompt('')
    setBusy(true)
    void (async () => {
      try {
        settingsRef.current = await window.mailApi.getAiSettings()
        if (!mountedRef.current) return
        loop.run(instruction)
      } catch (err) {
        if (!mountedRef.current) return
        patchLast({
          streaming: false,
          text: err instanceof Error ? err.message : String(err),
          isError: true,
        })
        setBusy(false)
      }
    })()
  }

  function stop(): void {
    loopRef.current?.cancel()
  }

  function toggleAccount(accountId: string): void {
    setPicked((prev) => {
      if (prev.includes(accountId)) {
        // Keep at least one account: the scope is never empty.
        if (prev.length <= 1) return prev
        return prev.filter((id) => id !== accountId)
      }
      if (prev.length >= 8) return prev
      return [...prev, accountId]
    })
  }

  function quickSend(preset: string): void {
    if (loopRef.current?.busy) setPrompt(preset)
    else send(preset)
  }

  function renderCitedText(text: string, keyPrefix: string): ReactElement {
    const parts: Array<string | ReactElement> = []
    const pattern = new RegExp(CITATION_RE)
    let lastIndex = 0
    let match: RegExpExecArray | null
    let citeIndex = 0
    while ((match = pattern.exec(text)) !== null) {
      if (match.index > lastIndex) parts.push(text.slice(lastIndex, match.index))
      const uid = match[1] ?? ''
      parts.push(
        <button
          key={`${keyPrefix}-cite-${citeIndex++}`}
          type="button"
          className="mail-ai-cite"
          title={t('mail.ai.cite')}
          onClick={() => onOpenCitation({ uid })}
        >
          [{uid}]
        </button>,
      )
      lastIndex = match.index + match[0].length
    }
    if (lastIndex < text.length) parts.push(text.slice(lastIndex))
    return <span dir="auto">{parts}</span>
  }

  return (
    <section className="mail-ai-panel" aria-label={t('mail.ai.title')} data-scope={scopeKey}>
      <header className="mail-ai-panel-head">
        <span className="mail-ai-panel-title">{t('mail.ai.title')}</span>
      </header>
      <fieldset className="mail-ai-picker">
        <legend className="mail-ai-picker-note">{t('mail.ai.scopedAccount')}</legend>
        {accounts.map((account) => (
          <label key={account.accountId} className="mail-ai-pick">
            <input
              type="checkbox"
              checked={picked.includes(account.accountId)}
              onChange={() => toggleAccount(account.accountId)}
            />
            {account.label}
          </label>
        ))}
      </fieldset>
      <div className="mail-ai-quick">
        <button
          type="button"
          onClick={() =>
            quickSend(
              'Triage the current folder headers for my selected accounts: priority, needs-reply, newsletter, noise.',
            )
          }
        >
          {t('mail.ai.triage')}
        </button>
        <button
          type="button"
          onClick={() =>
            quickSend('Summarize the thread I have open, one cited bullet per message.')
          }
        >
          {t('mail.ai.summarize')}
        </button>
        <button
          type="button"
          onClick={() =>
            quickSend(
              'Draft a short reply from the open thread as editable text. Do not send anything.',
            )
          }
        >
          {t('mail.ai.composeAssist')}
        </button>
        <button
          type="button"
          onClick={() =>
            quickSend(
              'List the open questions in my selected mailbox, each with its message citation.',
            )
          }
        >
          {t('mail.ai.askMail')}
        </button>
      </div>
      <div className="mail-ai-chat" ref={chatRef}>
        {chat.length === 0 && <div className="mail-ai-empty">{t('mail.ai.emptyResult')}</div>}
        {chat.map((entry, i) => {
          if (entry.role === 'user') {
            return (
              <div key={i} className="mail-ai-msg mail-ai-msg-user">
                <span dir="auto">{entry.text}</span>
              </div>
            )
          }
          const hasTools = (entry.tools?.length ?? 0) > 0
          if (!entry.text && !entry.streaming && !hasTools) return null
          const done = !entry.streaming && !entry.isError
          return (
            <div
              key={i}
              className={`mail-ai-msg mail-ai-msg-assistant${entry.isError ? ' mail-ai-msg-error' : ''}${entry.streaming ? ' mail-ai-msg-streaming' : ''}`}
            >
              {!entry.text && entry.streaming ? (
                <AiTypingIndicator label={t('mail.ai.title')} />
              ) : (
                entry.text && renderCitedText(entry.text, `m${i}`)
              )}
              {hasTools && (
                <ul className="mail-ai-tools">
                  {(entry.tools ?? []).map((tool, j) => (
                    <li
                      key={j}
                      className={`mail-ai-tool${tool.running ? ' is-running' : ''}${tool.isError ? ' is-error' : ''}`}
                    >
                      {tool.summary}
                    </li>
                  ))}
                </ul>
              )}
              {done && entry.text && (
                <div className="mail-ai-msg-actions">
                  <button type="button" onClick={() => onInsertAssist(entry.text)}>
                    {t('mail.ai.composeAssist')}
                  </button>
                </div>
              )}
            </div>
          )
        })}
      </div>
      <div className="mail-ai-composer">
        <AiComposer
          value={prompt}
          busy={busy}
          placeholder={t('mail.ai.placeholder')}
          hintIdle={t('mail.ai.scopedAccount')}
          hintBusy={t('mail.ai.scopedAccount')}
          sendLabel={t('mail.ai.send')}
          stopLabel={t('mail.ai.stop')}
          onChange={setPrompt}
          onSend={() => send(prompt)}
          onStop={stop}
        />
      </div>
    </section>
  )
}
