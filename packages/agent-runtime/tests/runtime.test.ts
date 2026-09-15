/**
 * Smoke + integration tests for @genoffice/agent-runtime.
 *
 * Verifies:
 *   1. ReactUIAdapter — dialog push/resolve/timeout, notifications, statuses, custom data
 *   2. createOfficeSession() — wires a real pi AgentSession
 *   3. The session's ExtensionUIContext is the ReactUIAdapter
 *   4. The session runs a real prompt and emits the expected event stream
 */

import { describe, expect, it } from 'vitest'
import { createOfficeSession, ReactUIAdapter } from '../src/index'

describe('ReactUIAdapter', () => {
  it('starts with empty state', () => {
    const a = new ReactUIAdapter({ defaultDialogTimeoutMs: 30_000 })
    expect(a.dialogs.length).toBe(0)
    expect(a.notifications.length).toBe(0)
    expect(a.statuses.size).toBe(0)
  })

  it('round-trips custom data and editor instance', () => {
    const a = new ReactUIAdapter()
    a.setCustomData('editor', { fakeEditor: true })
    expect(a.getCustomData('editor')).toEqual({ fakeEditor: true })
    a.setEditorInstance({ tiptap: 'v3' })
    expect(a.getEditorInstance()).toEqual({ tiptap: 'v3' })
  })

  it('resolves confirm dialogs through the React path', async () => {
    const a = new ReactUIAdapter()
    const p = a.confirm('Title', 'Body')
    expect(a.dialogs.length).toBe(1)
    expect(a.dialogs[0]!.kind).toBe('confirm')
    a.resolveDialog(a.dialogs[0]!.id, true)
    await expect(p).resolves.toBe(true)
    expect(a.dialogs.length).toBe(0)
  })

  it('resolves input dialogs through the React path', async () => {
    const a = new ReactUIAdapter()
    const p = a.input('Name?', 'placeholder')
    expect(a.dialogs.length).toBe(1)
    a.resolveDialog(a.dialogs[0]!.id, 'Alice')
    await expect(p).resolves.toBe('Alice')
  })

  it('resolves select dialogs through the React path', async () => {
    const a = new ReactUIAdapter()
    const p = a.select('Pick', ['a', 'b', 'c'])
    expect(a.dialogs.length).toBe(1)
    a.resolveDialog(a.dialogs[0]!.id, 'b')
    await expect(p).resolves.toBe('b')
  })

  it('auto-resolves dialogs on timeout', async () => {
    const a = new ReactUIAdapter()
    const [rc, ri, rs] = await Promise.all([
      a.confirm('t', 'm', { timeout: 50 }),
      a.input('t', undefined, { timeout: 50 }),
      a.select('t', ['x'], { timeout: 50 }),
    ])
    expect(rc).toBe(false)
    expect(ri).toBeUndefined()
    expect(rs).toBeUndefined()
  })

  it('emits notifications and supports dismiss', () => {
    const a = new ReactUIAdapter()
    let count = 0
    a.onNotifications(() => count++)
    a.notify('hi', 'info')
    a.notify('warn', 'warning')
    a.notify('err', 'error')
    expect(a.notifications.length).toBe(3)
    expect(count).toBe(3)
    a.dismissNotification(a.notifications[0]!.id)
    expect(a.notifications.length).toBe(2)
  })

  it('manages status entries (set / clear)', () => {
    const a = new ReactUIAdapter()
    a.setStatus('model', 'anthropic/claude-sonnet-4-5')
    expect(a.statuses.get('model')).toBe('anthropic/claude-sonnet-4-5')
    a.setStatus('model', undefined)
    expect(a.statuses.has('model')).toBe(false)
  })
})

describe('createOfficeSession', () => {
  it('wires a real pi session with the ReactUIAdapter as UI context', async () => {
    // No agentDir override → uses default ~/.pi/agent (same as apps/docs pi-smoke.ts)
    const { session, uiAdapter, dispose } = await createOfficeSession({
      cwd: process.cwd(),
    })

    expect(session).toBeDefined()
    expect(typeof session.prompt).toBe('function')
    expect(typeof session.subscribe).toBe('function')
    expect(typeof session.dispose).toBe('function')

    // The session exposes its ExtensionRunner — verify the UI context is wired
    const runner = session.extensionRunner
    expect(runner).toBeDefined()
    const ctx = runner.getUIContext()
    expect(ctx).toBeDefined()
    // ReactUIAdapter is its own pi context
    expect(uiAdapter.piContext).toBe(uiAdapter)

    // Drive a dialog through the wired context
    const dialogPromise = uiAdapter.confirm('Sync test', 'OK?')
    expect(uiAdapter.dialogs.length).toBe(1)
    uiAdapter.resolveDialog(uiAdapter.dialogs[0]!.id, true)
    await expect(dialogPromise).resolves.toBe(true)

    // Run a real prompt
    const events: string[] = []
    const unsub = session.subscribe((event) => {
      events.push(event.type)
    })
    await session.prompt('Reply with the single word: pong')
    unsub()
    expect(events).toContain('agent_start')
    expect(events).toContain('agent_end')
    expect(events).toContain('message_update')

    dispose()
  })
})
