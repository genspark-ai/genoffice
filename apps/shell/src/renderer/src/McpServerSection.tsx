import { useCallback, useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import type { McpStatus } from '../../shared/home-api'
import type { TFunc } from './locale'

// ── Settings → Integrations → MCP → local HTTP server ─────────
// The in-app MCP server that drives the live Word editor. Lives in its own
// component so the settings modal does not carry the server state for the
// other panes; polling the log only while this block is mounted.

/**
 * mcp.json snippet a client needs to reach the local server. Named apart from
 * the stdio `genoffice` entry (part A) so both can live in one config file.
 */
export function mcpConfigExample(port: string): string {
  return JSON.stringify(
    { mcpServers: { 'genoffice-editor': { url: `http://127.0.0.1:${port}/mcp` } } },
    null,
    2,
  )
}

function Field({ label, value, action }: { label: string; value: string; action?: ReactNode }) {
  return (
    <div className="set-field">
      <div className="set-field-text">
        <div className="set-field-label">{label}</div>
        <div className="set-field-value">{value}</div>
      </div>
      {action}
    </div>
  )
}

/** copies the given text; the label flips to "Copied" for a moment as feedback */
function CopyTextButton({ t, text }: { t: TFunc; text: string }) {
  const [copied, setCopied] = useState(false)
  const timerRef = useRef<number | null>(null)
  useEffect(
    () => () => {
      if (timerRef.current !== null) window.clearTimeout(timerRef.current)
    },
    [],
  )
  return (
    <button
      className="set-btn"
      onClick={() => {
        void navigator.clipboard.writeText(text).then(
          () => {
            setCopied(true)
            if (timerRef.current !== null) window.clearTimeout(timerRef.current)
            timerRef.current = window.setTimeout(() => setCopied(false), 1600)
          },
          () => {},
        )
      }}
    >
      {copied ? t('setMcpCopied') : t('setMcpCopy')}
    </button>
  )
}

export function McpServerSection({ t }: { t: TFunc }) {
  const [running, setRunning] = useState(false)
  const [enabled, setEnabled] = useState(false)
  const [port, setPort] = useState('3093')
  // last port the main process confirmed; an invalid edit reverts to it, not to a
  // literal default that may disagree with a server already running elsewhere
  const savedPort = useRef('3093')
  const [error, setError] = useState('')
  const [saving, setSaving] = useState(false)
  const [background, setBackground] = useState(false)
  const [logging, setLogging] = useState(false)
  const [logs, setLogs] = useState<string[]>([])
  const logViewRef = useRef<HTMLPreElement | null>(null)

  const takeStatus = (s: McpStatus) => {
    setRunning(s.running)
    setEnabled(s.enabled)
    setPort(String(s.port))
    savedPort.current = String(s.port)
    setBackground(s.background)
    setLogging(s.logging)
    setError(s.error ?? '')
  }

  useEffect(() => {
    let alive = true
    void window.aiOffice.getMcpStatus?.().then((s) => {
      if (alive) takeStatus(s)
    })
    return () => {
      alive = false
    }
  }, [])

  /**
   * Always sends the full current snapshot plus the changed field: the main
   * process persists and applies atomically, restarting the server when the
   * port or the exposed tool set changes.
   */
  const apply = (patch: {
    enabled?: boolean
    port?: number
    background?: boolean
    logging?: boolean
  }) => {
    setSaving(true)
    void window.aiOffice
      .setMcpSettings({
        enabled: patch.enabled ?? enabled,
        port: patch.port ?? (Number(port) || 3093),
        background: patch.background ?? background,
        logging: patch.logging ?? logging,
      })
      .then(takeStatus)
      .catch(() => {})
      .finally(() => setSaving(false))
  }

  const fetchLogs = useCallback(() => {
    void window.aiOffice
      .getMcpLogs?.()
      .then((lines) => setLogs(Array.isArray(lines) ? lines : []))
      .catch(() => {})
  }, [])

  // poll while the log viewer is visible: entries land as tools run
  useEffect(() => {
    if (!logging) return
    fetchLogs()
    const timer = window.setInterval(fetchLogs, 2000)
    return () => window.clearInterval(timer)
  }, [logging, fetchLogs])

  // follow the tail (newest entries last), but never fight a reader who scrolled up
  useEffect(() => {
    const el = logViewRef.current
    if (!el) return
    if (el.scrollHeight - el.scrollTop - el.clientHeight < 48) el.scrollTop = el.scrollHeight
  }, [logs])

  const url = (path: string) => `http://127.0.0.1:${port}/${path}`

  return (
    <div className="set-mcp-server">
      <div className="set-field set-field-top">
        <div className="set-field-text">
          <div className="set-field-stack">
            <div className="set-field-label">{t('setMcp')}</div>
            <div className="set-field-desc">{t('setMcpDesc')}</div>
            <div className="set-field-desc set-mcp-status">
              <span
                className={`set-status-dot${error ? ' error' : running ? ' running' : ''}`}
                aria-hidden="true"
              />
              <span>{error || (running ? t('setMcpRunning') : t('setMcpStopped'))}</span>
            </div>
          </div>
        </div>
        <button
          className="set-switch"
          role="switch"
          aria-checked={enabled}
          aria-label={t('setMcp')}
          disabled={saving}
          onClick={() => apply({ enabled: !enabled })}
        />
      </div>
      <div className="set-field">
        <div className="set-field-text">
          <div className="set-field-stack">
            <label className="set-field-label" htmlFor="set-mcp-port">
              {t('setMcpPort')}
            </label>
          </div>
        </div>
        <input
          id="set-mcp-port"
          className="set-input"
          type="number"
          min={1024}
          max={65535}
          value={port}
          onChange={(e) => setPort(e.target.value)}
          onBlur={() => {
            const next = Number(port)
            if (!Number.isInteger(next) || next < 1024 || next > 65535) {
              setPort(savedPort.current)
              return
            }
            if (String(next) === savedPort.current) return
            // persisted while the server is off too, so the port is set before it starts
            apply({ port: next })
          }}
        />
      </div>
      <div className="set-field set-field-top">
        <div className="set-field-text">
          <div className="set-field-stack">
            <div className="set-field-label">{t('setMcpBg')}</div>
            <div className="set-field-desc">{t('setMcpBgDesc')}</div>
          </div>
        </div>
        <button
          className="set-switch"
          role="switch"
          aria-checked={background}
          aria-label={t('setMcpBg')}
          disabled={saving}
          onClick={() => apply({ background: !background })}
        />
      </div>
      <div className="set-config-block">
        <div className="set-config-head">
          <div className="set-field-stack">
            <div className="set-field-label">{t('setMcpConfig')}</div>
            <div className="set-field-desc">{t('setMcpConfigDesc')}</div>
          </div>
          <CopyTextButton t={t} text={mcpConfigExample(port)} />
        </div>
        <pre className="set-code">{mcpConfigExample(port)}</pre>
      </div>

      <details className="set-intg-details set-mcp-advanced">
        <summary>{t('intgMcpHttpAdvanced')}</summary>
        <Field
          label={t('setMcpUrlHttp')}
          value={url('mcp')}
          action={<CopyTextButton t={t} text={url('mcp')} />}
        />
        <Field
          label={t('setMcpUrlSse')}
          value={url('sse')}
          action={<CopyTextButton t={t} text={url('sse')} />}
        />
        <Field
          label={t('setMcpHealth')}
          value={url('health')}
          action={<CopyTextButton t={t} text={url('health')} />}
        />
        <div className="set-field set-field-top">
          <div className="set-field-text">
            <div className="set-field-stack">
              <div className="set-field-label">{t('setMcpLog')}</div>
              <div className="set-field-desc">{t('setMcpLogDesc')}</div>
            </div>
          </div>
          <button
            className="set-switch"
            role="switch"
            aria-checked={logging}
            aria-label={t('setMcpLog')}
            disabled={saving}
            onClick={() => apply({ logging: !logging })}
          />
        </div>
        {logging && (
          <div className="set-config-block">
            <div className="set-config-head">
              <div className="set-field-stack">
                <div className="set-field-label">{t('setMcpLogFile')}</div>
                <div className="set-field-desc">mcp-log.txt</div>
              </div>
              <div className="set-btn-row">
                <button className="set-btn" onClick={fetchLogs}>
                  {t('setMcpLogRefresh')}
                </button>
                <button className="set-btn" onClick={() => void window.aiOffice.openMcpLogFile?.()}>
                  {t('setMcpLogOpen')}
                </button>
                <button
                  className="set-btn"
                  onClick={() => {
                    void window.aiOffice.clearMcpLogs?.().then(fetchLogs)
                  }}
                >
                  {t('setMcpLogClear')}
                </button>
              </div>
            </div>
            <pre ref={logViewRef} className="set-code set-log">
              {logs.length > 0 ? (
                logs.join('\n')
              ) : (
                <span className="set-log-empty">{t('setMcpLogEmpty')}</span>
              )}
            </pre>
          </div>
        )}
      </details>
    </div>
  )
}
