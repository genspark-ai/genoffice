import { useEffect, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent, ReactElement } from 'react'
import { AgentLoop } from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import { AiComposer, AiTypingIndicator, Markdown } from '@genoffice/ui'
import { aiLangDirective, t as tGlobal, useI18n } from '../i18n/locale'
import { fileNameOf, type HangulStudioFacade } from '../studio-text'
import { createHangulSkill, type HangulSkillDeps } from './hangul-skill'
import { createElectronTransport } from './transport'

const PANEL_WIDTH_KEY = 'hwp-ai-panel-width'
const PANEL_WIDTH_DEFAULT = 360
const PANEL_WIDTH_MIN = 280

function clampPanelWidth(w: number): number {
  const max = Math.max(PANEL_WIDTH_MIN, Math.min(720, Math.round(window.innerWidth * 0.6)))
  return Math.min(Math.max(w, PANEL_WIDTH_MIN), max)
}

function loadPanelWidth(): number {
  const saved = Number(localStorage.getItem(PANEL_WIDTH_KEY))
  return Number.isFinite(saved) && saved > 0
    ? Math.min(Math.max(saved, PANEL_WIDTH_MIN), 720)
    : PANEL_WIDTH_DEFAULT
}

interface ToolActivity {
  name: string
  summary: string
  isError?: boolean
}

interface ChatEntry {
  role: 'user' | 'assistant'
  text: string
  streaming?: boolean
  isError?: boolean
  undelivered?: boolean
  loginRequired?: boolean
  tools?: ToolActivity[]
}

type Phase = 'thinking' | 'replying' | 'working'

export function AiPanel({
  facade,
  filePath,
  onCollapse,
}: {
  facade: HangulStudioFacade | null
  filePath: string | null
  onCollapse: () => void
}): ReactElement {
  const { lang, t } = useI18n()
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [prompt, setPrompt] = useState('')
  const [busy, setBusy] = useState(false)
  const [phase, setPhase] = useState<Phase>('thinking')
  const [resizing, setResizing] = useState(false)
  const preferredWidthRef = useRef(loadPanelWidth())
  const [panelWidth, setPanelWidth] = useState(() => clampPanelWidth(preferredWidthRef.current))
  const chatRef = useRef<HTMLDivElement>(null)
  const stickToBottomRef = useRef(true)
  const asideRef = useRef<HTMLElement>(null)
  const mountedRef = useRef(true)
  const settingsRef = useRef<AiSettings | null>(null)
  const langRef = useRef(lang)
  langRef.current = lang
  const facadeRef = useRef(facade)
  facadeRef.current = facade
  const filePathRef = useRef(filePath)
  filePathRef.current = filePath
  const contextRef = useRef({
    pageCount: 0,
    hasSelection: false,
    selectionPreview: null as string | null,
  })

  useEffect(() => {
    const dock = asideRef.current?.closest('.ai-dock') as HTMLElement | null
    dock?.style.setProperty('--ai-panel-width', `${panelWidth}px`)
  }, [panelWidth])

  const patchLast = (patch: Partial<ChatEntry> | ((last: ChatEntry) => Partial<ChatEntry>)) => {
    setChat((prev) => {
      const next = [...prev]
      const last = next[next.length - 1]
      if (!last || last.role !== 'assistant') return prev
      next[next.length - 1] = { ...last, ...(typeof patch === 'function' ? patch(last) : patch) }
      return next
    })
  }

  const loopRef = useRef<AgentLoop | null>(null)
  if (!loopRef.current) {
    const skillDeps = (): HangulSkillDeps => ({
      fileName: () => fileNameOf(filePathRef.current),
      pageCount: () => contextRef.current.pageCount,
      hasSelection: () => contextRef.current.hasSelection,
      selectionPreview: () => contextRef.current.selectionPreview,
      getDocumentText: async () => {
        const studio = facadeRef.current
        if (!studio) throw new Error('Hangul editor is not ready')
        return studio.getPlainText()
      },
      getSelection: async () => {
        const studio = facadeRef.current
        if (!studio) return null
        return studio.getSelectionText()
      },
    })
    loopRef.current = new AgentLoop({
      transport: createElectronTransport(() => settingsRef.current!),
      skill: createHangulSkill(skillDeps),
      systemSuffix: () => aiLangDirective(langRef.current),
      events: {
        onText: (text) => {
          setPhase('replying')
          patchLast({ text })
        },
        onToolStart: (call) => {
          setPhase('working')
          patchLast((last) => ({
            tools: [
              ...(last.tools ?? []),
              { name: call.name, summary: call.name.replace(/[_-]+/g, ' ') },
            ],
          }))
        },
        onToolExecuted: ({ call, execution }) => {
          const activity: ToolActivity = {
            name: call.name,
            summary: execution.summary,
            isError: execution.isError,
          }
          patchLast((last) => {
            const tools = [...(last.tools ?? [])]
            if (tools.at(-1)?.name === call.name) tools.pop()
            return { tools: [...tools, activity] }
          })
        },
        onTurnEnd: () => {
          patchLast({ streaming: false })
          setChat((prev) => [...prev, { role: 'assistant', text: '', streaming: true }])
        },
        onDone: ({ text, cancelled, turnLimit, truncated }) => {
          const base = turnLimit
            ? [text, tGlobal('aiTurnLimit')].filter(Boolean).join('\n\n')
            : text || (cancelled ? tGlobal('aiStopped') : '')
          const final = truncated
            ? [base, tGlobal('aiTruncatedNote')].filter(Boolean).join('\n\n')
            : base
          patchLast((last) => ({
            streaming: false,
            text: final || (last.tools?.length ? last.text : tGlobal('aiNoReply')),
          }))
          setBusy(false)
        },
        onError: (error) => {
          setChat((prev) => {
            const next = [...prev]
            for (let i = next.length - 1; i >= 0; i--) {
              const entry = next[i]!
              if (entry.role === 'user') {
                next[i] = { ...entry, undelivered: true }
                break
              }
            }
            const last = next.at(-1)
            if (last?.role === 'assistant') {
              next[next.length - 1] = { ...last, streaming: false, text: error, isError: true }
            }
            return next
          })
          void window.hwpApi
            .aiGskStatus()
            .then((status) => {
              if (status.loggedIn) return
              setChat((prev) => {
                const next = [...prev]
                const last = next.at(-1)
                if (last?.role === 'assistant' && last.isError) {
                  next[next.length - 1] = { ...last, loginRequired: true }
                }
                return next
              })
            })
            .catch(() => undefined)
          setBusy(false)
        },
      },
    })
  }

  useEffect(() => {
    mountedRef.current = true
    return () => {
      mountedRef.current = false
      loopRef.current?.cancel()
    }
  }, [])

  useEffect(() => {
    if (stickToBottomRef.current) {
      chatRef.current?.scrollTo({ top: chatRef.current.scrollHeight })
    }
  }, [chat, busy])

  const onChatScroll = (): void => {
    const el = chatRef.current
    if (!el) return
    stickToBottomRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 48
  }

  const refreshContext = async (): Promise<void> => {
    const studio = facadeRef.current
    if (!studio) {
      contextRef.current = { pageCount: 0, hasSelection: false, selectionPreview: null }
      return
    }
    const [pageCount, preview, selected] = await Promise.all([
      studio.pageCount().catch(() => 0),
      studio.getSelectionText().catch(() => null),
      studio.hasSelection().catch(() => false),
    ])
    contextRef.current = {
      pageCount,
      hasSelection: Boolean(preview?.trim()) || selected,
      selectionPreview: preview,
    }
  }

  const send = (text: string): void => {
    const instruction = text.trim()
    const loop = loopRef.current
    if (!instruction || !loop || loop.busy || !facadeRef.current) return
    stickToBottomRef.current = true
    setPhase('thinking')
    setChat((prev) => [
      ...prev,
      { role: 'user', text: instruction },
      { role: 'assistant', text: '', streaming: true },
    ])
    setPrompt('')
    setBusy(true)
    void (async () => {
      try {
        await refreshContext()
        settingsRef.current = await window.hwpApi.getAiSettings()
        if (!mountedRef.current) return
        await loop.run(instruction)
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

  const stop = (): void => loopRef.current?.cancel()

  useEffect(() => {
    const onResize = (): void => setPanelWidth(clampPanelWidth(preferredWidthRef.current))
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])

  const resizeCleanupRef = useRef<(() => void) | null>(null)
  useEffect(() => () => resizeCleanupRef.current?.(), [])

  const startResize = (e: ReactPointerEvent<HTMLDivElement>): void => {
    e.preventDefault()
    const resizer = e.currentTarget
    setResizing(true)
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    const onMove = (ev: PointerEvent): void => {
      const w = clampPanelWidth(ev.clientX)
      preferredWidthRef.current = w
      setPanelWidth(w)
    }
    let done = false
    const cleanup = (): void => {
      if (done) return
      done = true
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', cleanup)
      window.removeEventListener('pointercancel', cleanup)
      resizer.removeEventListener('lostpointercapture', cleanup)
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
      setResizing(false)
      localStorage.setItem(PANEL_WIDTH_KEY, String(Math.round(preferredWidthRef.current)))
    }
    resizeCleanupRef.current = cleanup
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', cleanup)
    window.addEventListener('pointercancel', cleanup)
    resizer.addEventListener('lostpointercapture', cleanup)
    resizer.setPointerCapture(e.pointerId)
  }

  const typingLabel =
    phase === 'replying' ? t('aiReplying') : phase === 'working' ? t('aiWorking') : t('aiThinking')

  return (
    <aside
      ref={asideRef}
      className={`copilot${resizing ? ' ai-panel-resizing' : ''}`}
      style={{ width: '100%' }}
      dir={lang === 'ar' || lang === 'he' ? 'rtl' : undefined}
    >
      <div
        className="ai-panel-resizer"
        onPointerDown={startResize}
        role="separator"
        aria-orientation="vertical"
        aria-label="Genspark"
      />
      <header className="ai-panel-header">
        <span className="ai-panel-title">
          <GensparkMark size={22} />
          Genspark
        </span>
        <div className="ai-panel-header-actions">
          {chat.length > 0 && (
            <button
              className="ai-header-btn"
              onClick={() => {
                stop()
                loopRef.current?.reset()
                setBusy(false)
                setChat([])
              }}
              data-tip={t('aiNewChat')}
              aria-label={t('aiNewChat')}
            >
              <IconNewChat />
            </button>
          )}
          <button
            className="ai-header-btn"
            onClick={onCollapse}
            data-tip={t('aiCollapsePanel')}
            aria-label={t('aiCollapsePanel')}
          >
            <IconCollapse />
          </button>
        </div>
      </header>

      <div className="ai-chat" ref={chatRef} onScroll={onChatScroll}>
        {chat.length === 0 && (
          <div className="ai-chat-empty">
            <div className="ai-chat-empty-title">{t('aiEmptyTitle')}</div>
            <div className="ai-chat-empty-body">{t('aiEmptyBody')}</div>
          </div>
        )}
        {chat.map((entry, i) => {
          if (entry.role === 'user') {
            return (
              <div key={i} className="ai-msg ai-msg-user">
                <span dir="auto">{entry.text}</span>
                {entry.undelivered && (
                  <div className="ai-msg-undelivered">
                    {t('aiUndelivered')}
                    {!busy && (
                      <button className="ai-retry-btn" onClick={() => send(entry.text)}>
                        {t('aiRetry')}
                      </button>
                    )}
                  </div>
                )}
              </div>
            )
          }
          const hasTools = (entry.tools?.length ?? 0) > 0
          if (!entry.text && !hasTools) return null
          return (
            <div
              key={i}
              className={`ai-msg ai-msg-assistant${entry.isError ? ' ai-msg-error' : ''}`}
            >
              {hasTools &&
                entry.tools!.map((tool, ti) => (
                  <div key={ti} className={`ai-tool-row${tool.isError ? ' error' : ''}`}>
                    {tool.summary}
                  </div>
                ))}
              {entry.text && (
                <div dir="auto">
                  <Markdown text={entry.text} />
                </div>
              )}
              {entry.loginRequired && (
                <button className="ai-login-btn" onClick={() => void window.hwpApi.aiGskLogin()}>
                  {t('aiGskLoginBtn')}
                </button>
              )}
            </div>
          )
        })}
        {busy && <AiTypingIndicator label={typingLabel} />}
      </div>

      <div className="ai-composer">
        <AiComposer
          value={prompt}
          busy={busy}
          placeholder={facade ? t('aiComposerPlaceholder') : t('aiEditorNotReady')}
          hintIdle={t('aiHintIdle')}
          hintBusy={t('aiHintBusy')}
          sendLabel={t('aiSend')}
          stopLabel={t('aiStop')}
          onChange={setPrompt}
          onSend={() => send(prompt)}
          onStop={stop}
        />
      </div>
    </aside>
  )
}

function IconNewChat() {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden>
      <path
        d="M3 3.5h7.5A1.5 1.5 0 0 1 12 5v2.2M3 3.5v7A1.5 1.5 0 0 0 4.5 12H8"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
      />
      <path d="M10 10.5h4M12 8.5v4" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
    </svg>
  )
}

function IconCollapse() {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden>
      <path d="M10 3.5 6 8l4 4.5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
    </svg>
  )
}

export function GensparkMark({ size = 18 }: { size?: number }): React.JSX.Element {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 130 130.025"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      aria-hidden
    >
      <path
        d="M105.115 0H24.6428C11.0443 0 0 11.0686 0 24.6915V105.334C0 118.981 11.0199 130.025 24.6428 130.025H105.115C118.714 130.025 129.758 118.957 129.758 105.334V24.6915C129.758 11.0443 118.714 0 105.115 0ZM71.5201 35.2735C85.5078 33.1571 86.7729 31.9164 88.865 17.88C88.938 17.4421 89.3028 17.1259 89.7407 17.1259C90.1786 17.1259 90.5435 17.4421 90.6164 17.88C92.7328 31.8921 93.9735 33.1571 107.961 35.2735C108.399 35.3465 108.715 35.7114 108.715 36.1493C108.715 36.5871 108.399 36.952 107.961 37.025C93.9249 39.1414 92.7085 40.4064 90.5677 54.6131C90.5191 54.9537 90.2516 55.197 89.911 55.197C89.5704 55.197 89.3028 54.9537 89.2542 54.6131C87.1134 40.4064 85.5565 39.1658 71.4958 37.025C71.0579 36.952 70.7417 36.5871 70.7417 36.1493C70.7417 35.7114 71.0579 35.3465 71.4958 35.2735H71.5201ZM101.758 78.5261C101.758 78.8181 101.563 79.037 101.271 79.0856C92.3193 80.4236 91.5652 81.2264 90.2029 90.2759C90.1786 90.4948 89.9839 90.6408 89.7893 90.6408C89.5703 90.6408 89.4001 90.4948 89.3758 90.2759C88.0135 81.2507 87.0161 80.4479 78.0883 79.0856C77.7964 79.037 77.6017 78.7937 77.6017 78.5261C77.6017 78.2342 77.7964 78.0153 78.0883 77.9666C86.9918 76.6287 87.7703 75.8259 89.1326 66.898C89.1812 66.6061 89.4244 66.4115 89.692 66.4115C89.9839 66.4115 90.2028 66.6061 90.2515 66.898C91.5894 75.8259 92.3923 76.6043 101.296 77.9666C101.588 78.0153 101.782 78.2585 101.782 78.5261H101.758ZM16.5178 54.8077C16.5178 54.1023 17.0286 53.4941 17.7341 53.3968C40.1388 50.0154 42.1093 47.9963 45.4907 25.5672C45.588 24.8861 46.1961 24.3509 46.9016 24.3509C47.6071 24.3509 48.191 24.8617 48.3126 25.5672C51.694 47.9963 53.6887 50.0154 76.0691 53.3968C76.7503 53.4941 77.2855 54.1023 77.2855 54.8077C77.2855 55.5132 76.7746 56.1214 76.0691 56.2187C53.5914 59.6244 51.6696 61.6192 48.2639 84.3645C48.1909 84.8754 47.7287 85.2889 47.2179 85.2889C46.707 85.2889 46.2448 84.8997 46.1718 84.3645C42.7418 61.6435 40.2604 59.6244 17.7584 56.2187C17.0772 56.1214 16.542 55.5132 16.542 54.8077H16.5178ZM112.097 109.591C112.097 111.416 110.613 112.9 108.813 112.9H21.2614C19.4369 112.9 17.9774 111.416 17.9774 109.591V102.658C17.9774 100.834 19.4612 99.3497 21.2614 99.3497H108.813C110.637 99.3497 112.097 100.834 112.097 102.658V109.591Z"
        fill="currentColor"
      />
    </svg>
  )
}
