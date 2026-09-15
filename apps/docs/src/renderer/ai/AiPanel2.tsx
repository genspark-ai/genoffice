/**
 * AiPanel2 — minimal demonstration of integrating `apps/docs` with the
 * new `@genoffice/agent-runtime` (pi-backed) event stream.
 *
 * This file is intentionally short and is the reference integration pattern
 * that the full `AiPanel.tsx` will be migrated to in a follow-up. The full
 * migration is deferred because the legacy panel depends on chat-runtime /
 * agent-core / ai-provider, which are being phased out per `agent1.md` §5.3.
 *
 * What it demonstrates:
 *   1. <PiSessionProvider> owns the session + UI adapter lifecycle
 *   2. <PiDialogHost /> and <NotificationToaster /> float above the panel
 *   3. Session events stream into React state via `session.subscribe(...)`
 *   4. Tool dialogs (e.g. read_blocks → confirm) auto-resolve when the user
 *      clicks OK / Cancel in the dialog host
 *   5. Provider extensions (docs-skill) are wired via `extensionFactories`
 *
 * To run interactively, mount <AiPanel2 cwd={process.cwd()} /> in your
 * App.tsx alongside the existing <AiPanel /> during the migration window.
 */

import { useEffect, useMemo, useRef, useState } from 'react'
import {
  PiSessionProvider,
  PiDialogHost,
  NotificationToaster,
  ReactUIAdapter,
  usePiSession,
  usePiNotifications,
  useUiAdapter,
} from '@genoffice/agent-runtime'
import { createDocsSkillExtension } from '@genoffice/agent-skills'
import type { ExtensionAPI } from '@earendil-works/pi-coding-agent'

interface Props {
  cwd: string
}

// ---------------------------------------------------------------------------
// Inner panel — must be inside <PiSessionProvider>
// ---------------------------------------------------------------------------

function AiPanelInner(_props: Props) {
  const { session } = usePiSession()
  const adapter = useUiAdapter()
  const notifications = usePiNotifications()
  const [input, setInput] = useState('Summarize this document.')
  const [events, setEvents] = useState<string[]>([])
  const [running, setRunning] = useState(false)
  const unsubRef = useRef<(() => void) | null>(null)

  useEffect(() => () => {
    unsubRef.current?.()
    unsubRef.current = null
  }, [])

  const onSend = async () => {
    if (running) return
    setRunning(true)
    setEvents([])
    const unsub = session.subscribe((ev) => {
      setEvents((prev) => [...prev.slice(-199), ev.type])
    })
    unsubRef.current = unsub
    try {
      await session.prompt(input)
    } catch (err) {
      adapter.notify(`Prompt failed: ${err instanceof Error ? err.message : String(err)}`, 'error')
    } finally {
      unsub()
      unsubRef.current = null
      setRunning(false)
    }
  }

  return (
    <div className="aipanel2" data-testid="aipanel2">
      <h2>AI Panel (pi EventStream demo)</h2>

      {/* The dialog host must be inside the provider to subscribe to usePiDialogs */}
      <PiDialogHost />
      <NotificationToaster />

      <textarea
        value={input}
        onChange={(e) => setInput(e.target.value)}
        rows={3}
        style={{ width: '100%' }}
        disabled={running}
        data-testid="aipanel2-input"
      />
      <button onClick={onSend} disabled={running} data-testid="aipanel2-send">
        {running ? 'Running…' : 'Send'}
      </button>

      <h3>Notifications ({notifications.length})</h3>
      <ul data-testid="aipanel2-notifications">
        {notifications.map((n) => (
          <li key={n.id} data-type={n.type}>[{n.type}] {n.message}</li>
        ))}
      </ul>

      <h3>Event stream ({events.length})</h3>
      <ol
        data-testid="aipanel2-events"
        style={{ maxHeight: 200, overflow: 'auto', fontFamily: 'monospace', fontSize: 12 }}
      >
        {events.map((e, i) => <li key={i}>{e}</li>)}
      </ol>
    </div>
  )
}

// ---------------------------------------------------------------------------
// Outer wrapper — owns the session lifecycle via PiSessionProvider
// ---------------------------------------------------------------------------

/**
 * Extension factory: wires the docs-skill (read_blocks, replace_document, ...)
 * against the live pi AgentSession. The factory closure captures the
 * ReactUIAdapter so dialogs/notifications are shared between the React UI
 * and the tool implementations.
 */
function docsSkillFactory(adapter: ReactUIAdapter) {
  return (pi: ExtensionAPI) => {
    createDocsSkillExtension({ uiAdapter: adapter })(pi)
  }
}

export function AiPanel2(props: Props) {
  // Stable adapter instance per <AiPanel2 /> mount so dialogs/notifications
  // survive re-renders. The same instance is passed to PiSessionProvider
  // and captured by the docs-skill extension factory.
  const adapter = useMemo(() => new ReactUIAdapter(), [])

  return (
    <PiSessionProvider
      cwd={props.cwd}
      uiAdapter={adapter}
      extensionFactories={[docsSkillFactory(adapter)]}
    >
      <AiPanelInner {...props} />
    </PiSessionProvider>
  )
}
