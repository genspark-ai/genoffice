import { useCallback, useEffect, useRef, useState } from 'react'
import { useI18n } from './locale'
import type { TFunc } from './locale'

/**
 * AiChatPane — minimal in-app AI chat backed by the real `POST /api/ai/stream`
 * SSE endpoint exposed by apps/web-server. Lives in the Settings modal so
 * users can verify the GenOffice chat pipeline end-to-end (browser → IPC →
 * web-server → MiniMax → SSE → UI render) without opening a docs tab.
 *
 * Why this exists:
 *  - apps/docs has its own rich AiPanel tied to a document. There's no
 *    "floating" GenOffice chat yet.
 *  - We need a place to type a free-form prompt, watch SSE tokens stream in,
 *    and prove the LLM call works in web mode (W29 verifier).
 */

interface ChatMessage {
  id: string
  role: 'user' | 'assistant' | 'system'
  content: string
  /** timestamp ms */
  ts: number
  /** streaming state for assistant turns */
  streaming?: boolean
}

const ROLE_LABEL: Record<ChatMessage['role'], string> = {
  user: '你',
  assistant: 'GenOffice AI',
  system: '系统',
}

function newId(): string {
  return Math.random().toString(36).slice(2, 10)
}

export function AiChatPane({ t: _t }: { t: TFunc }) {
  const { t } = useI18n()
  const [messages, setMessages] = useState<ChatMessage[]>([])
  const [input, setInput] = useState('')
  const [busy, setBusy] = useState(false)
  const [systemPrompt, setSystemPrompt] = useState('你是一个简洁有用的助手,使用中文回答。')
  const [error, setError] = useState<string | null>(null)
  const abortRef = useRef<AbortController | null>(null)
  const scrollerRef = useRef<HTMLDivElement | null>(null)

  // auto-scroll to bottom on new content
  useEffect(() => {
    const el = scrollerRef.current
    if (el) el.scrollTop = el.scrollHeight
  }, [messages])

  const send = useCallback(async () => {
    const trimmed = input.trim()
    if (!trimmed || busy) return

    const userMsg: ChatMessage = {
      id: newId(),
      role: 'user',
      content: trimmed,
      ts: Date.now(),
    }
    const assistantId = newId()
    const assistantMsg: ChatMessage = {
      id: assistantId,
      role: 'assistant',
      content: '',
      ts: Date.now(),
      streaming: true,
    }

    setMessages((prev) => [...prev, userMsg, assistantMsg])
    setInput('')
    setBusy(true)
    setError(null)

    const ac = new AbortController()
    abortRef.current = ac

    try {
      const res = await fetch('/api/ai/stream', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          system: systemPrompt,
          messages: [{ role: 'user', content: trimmed }],
        }),
        signal: ac.signal,
      })

      if (!res.ok || !res.body) {
        throw new Error(`HTTP ${res.status}: ${res.statusText}`)
      }

      const reader = res.body.getReader()
      const decoder = new TextDecoder()
      let buffer = ''
      interface StreamPayload {
        type?: string
        text?: string
        error?: string
        requestId?: string
      }
      let pending: StreamPayload[] = []

      while (true) {
        const { value, done } = await reader.read()
        if (done) break
        buffer += decoder.decode(value, { stream: true })

        // Split into SSE events (each starts with "data: ")
        const parts = buffer.split('\n\n')
        buffer = parts.pop() ?? ''
        for (const evt of parts) {
          for (const line of evt.split('\n')) {
            const m = /^data:\s?(.*)$/.exec(line)
            if (!m) continue
            try {
              const payload = JSON.parse(m[1])
              pending.push(payload)
            } catch {
              // ignore non-JSON lines (heartbeats)
            }
          }
        }

        // Apply pending updates to the streaming message
        if (pending.length > 0) {
          const toApply = pending
          pending = []
          setMessages((prev) =>
            prev.map((msg) => {
              if (msg.id !== assistantId) return msg
              let next = msg.content
              let errorText: string | undefined
              let finished = false
              for (const p of toApply) {
                if (p.type === 'delta' && typeof p.text === 'string') {
                  next += p.text
                } else if (p.type === 'error') {
                  errorText = p.error || '未知错误'
                  finished = true
                } else if (p.type === 'done') {
                  finished = true
                }
              }
              return {
                ...msg,
                content: next,
                streaming: !finished,
                error: errorText,
              }
            }),
          )
        }
      }

      // Mark streaming complete on EOF
      setMessages((prev) =>
        prev.map((msg) =>
          msg.id === assistantId ? { ...msg, streaming: false } : msg,
        ),
      )
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e)
      setError(msg)
      setMessages((prev) =>
        prev.map((m) =>
          m.id === assistantId
            ? { ...m, content: m.content || '(empty)', streaming: false }
            : m,
        ),
      )
    } finally {
      setBusy(false)
      abortRef.current = null
    }
  }, [input, busy, systemPrompt])

  const stop = useCallback(() => {
    abortRef.current?.abort()
  }, [])

  const clear = useCallback(() => {
    abortRef.current?.abort()
    setMessages([])
    setError(null)
  }, [])

  return (
    <div className="ai-chat-pane">
      <div className="ai-chat-header">
        <h3>{t('setSecAiChat')}</h3>
        <div className="ai-chat-meta">
          POST /api/ai/stream (SSE) → MiniMax
        </div>
      </div>

      <div className="ai-chat-system">
        <label className="set-field-label">系统提示</label>
        <textarea
          className="ai-chat-system-input"
          value={systemPrompt}
          onChange={(e) => setSystemPrompt(e.target.value)}
          rows={2}
          disabled={busy}
        />
      </div>

      <div className="ai-chat-scroller" ref={scrollerRef}>
        {messages.length === 0 && (
          <div className="ai-chat-empty">
            输入问题并按"发送",观察 SSE 流式响应从 web-server 流到浏览器。
          </div>
        )}
        {messages.map((m) => (
          <div key={m.id} className={`ai-chat-msg ai-chat-msg-${m.role}`}>
            <div className="ai-chat-role">{ROLE_LABEL[m.role]}</div>
            <div className="ai-chat-content">
              {m.content || (m.streaming ? '⏳ 思考中…' : '')}
            </div>
            {m.streaming && <span className="ai-chat-cursor">▍</span>}
          </div>
        ))}
      </div>

      {error && <div className="ai-chat-error">⚠️ {error}</div>}

      <div className="ai-chat-input-row">
        <textarea
          className="ai-chat-input"
          placeholder="向 GenOffice AI 提问…(Shift+Enter 换行,Enter 发送)"
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault()
              void send()
            }
          }}
          rows={2}
          disabled={busy}
        />
        <div className="ai-chat-actions">
          {busy ? (
            <button type="button" className="set-btn" onClick={stop}>
              停止
            </button>
          ) : (
            <button
              type="button"
              className="set-btn set-btn-primary"
              onClick={() => void send()}
              disabled={!input.trim()}
            >
              发送
            </button>
          )}
          <button type="button" className="set-btn" onClick={clear} disabled={busy && messages.length === 0}>
            清空
          </button>
        </div>
      </div>
    </div>
  )
}
