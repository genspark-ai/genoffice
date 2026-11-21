import React, { useEffect, useState } from 'react'
import type { Lang } from '@genoffice/i18n'
import { WEB_SHARE_LABELS } from './strings-web-share'
import './web-share-chip.css'

/** The preload bridge (installed by installWebShareBridge in every editor
 *  preload); typed loosely here so the ui package stays Electron-free. */
interface WebShareApi {
  status(): Promise<WebShareState>
  start(): Promise<WebShareState>
  stop(): Promise<WebShareState>
  onChanged(handler: (status: WebShareState) => void): () => void
}
interface WebShareState {
  running: boolean
  url?: string
  localUrl?: string
  port?: number
  error?: string
}
const api = (): WebShareApi | null =>
  typeof window === 'undefined'
    ? null
    : (((window as unknown as Record<string, unknown>).webShareApi as WebShareApi | undefined) ??
      null)

/**
 * Status-bar chip: one press starts the browser-access service and copies
 * the share link; the ✕ stops it. Hidden inside browser tabs (they ARE the
 * shared session — the service is controlled from the desktop app; http
 * covers both the served page and a vite dev server, which is fine: dev has
 * no reason to share and can still use the CLI flag).
 */
export function WebShareChip({ lang }: { lang: Lang }) {
  const [state, setState] = useState<WebShareState | null>(null)
  const [copied, setCopied] = useState(false)
  const [busy, setBusy] = useState(false)
  useEffect(() => {
    if (typeof window === 'undefined') return
    if (location.protocol === 'http:' || location.protocol === 'https:') return
    const bridge = api()
    if (!bridge) return
    let alive = true
    void bridge.status().then((s) => {
      if (alive) setState(s)
    })
    const off = bridge.onChanged((s) => {
      if (alive) setState(s)
    })
    return () => {
      alive = false
      off()
    }
  }, [])
  if (state === null) return null

  const labels = WEB_SHARE_LABELS[lang]
  const copyUrl = async (): Promise<void> => {
    const url = state.url ?? state.localUrl
    if (!url) return
    try {
      await navigator.clipboard.writeText(url)
      setCopied(true)
      window.setTimeout(() => setCopied(false), 1600)
    } catch {
      // clipboard refused (permissions): the tooltip still carries the URL
    }
  }
  const toggle = async (): Promise<void> => {
    const bridge = api()
    if (!bridge || busy) return
    setBusy(true)
    try {
      setState(state.running ? await bridge.stop() : await bridge.start())
      if (!state.running) void copyUrl()
    } finally {
      setBusy(false)
    }
  }

  if (!state.running) {
    return (
      <button
        type="button"
        className="web-share-chip status-btn"
        data-tip={state.error ? `${labels.shareTip} — ${state.error}` : labels.shareTip}
        disabled={busy}
        onClick={() => void toggle()}
      >
        <svg
          width="13"
          height="13"
          viewBox="0 0 16 16"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
          aria-hidden="true"
        >
          <circle cx="8" cy="8" r="1.6" fill="currentColor" stroke="none" />
          <ellipse cx="8" cy="8" rx="6.2" ry="2.6" />
          <ellipse cx="8" cy="8" rx="6.2" ry="2.6" transform="rotate(60 8 8)" />
          <ellipse cx="8" cy="8" rx="6.2" ry="2.6" transform="rotate(120 8 8)" />
        </svg>
        {labels.share}
      </button>
    )
  }
  return (
    <span className="web-share-chip web-share-chip-on">
      <button
        type="button"
        className="status-btn web-share-port"
        data-tip={copied ? '✓' : (state.url ?? state.localUrl ?? '')}
        disabled={busy}
        onClick={() => void copyUrl()}
      >
        <svg
          width="13"
          height="13"
          viewBox="0 0 16 16"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
          aria-hidden="true"
        >
          <circle cx="8" cy="8" r="1.6" fill="currentColor" stroke="none" />
          <ellipse cx="8" cy="8" rx="6.2" ry="2.6" />
          <ellipse cx="8" cy="8" rx="6.2" ry="2.6" transform="rotate(60 8 8)" />
          <ellipse cx="8" cy="8" rx="6.2" ry="2.6" transform="rotate(120 8 8)" />
        </svg>
        {copied ? '✓' : `:${state.port ?? ''}`}
      </button>
      <button
        type="button"
        className="status-btn web-share-stop"
        data-tip={labels.stopTip}
        aria-label={labels.stopTip}
        disabled={busy}
        onClick={() => void toggle()}
      >
        ×
      </button>
    </span>
  )
}
