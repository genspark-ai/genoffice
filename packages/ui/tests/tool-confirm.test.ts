/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { AiConfirmCard } from '../src/AiConfirmCard'
import { AI_CONFIRM_LABELS } from '../src/strings-ai-confirm'
import { useToolConfirm } from '../src/tool-confirm'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

/** a host that mounts the hook and exposes the card + decisions */
let host: HTMLDivElement
let root: ReturnType<typeof createRoot>
let hook: ReturnType<typeof useToolConfirm>

beforeEach(() => {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})
afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

/** renders the card the way a panel does: card + the two decisions */
function mountGate(timeoutMs?: number): void {
  function Gate() {
    hook = useToolConfirm(timeoutMs ? { timeoutMs } : undefined)
    if (!hook.card) return null
    return createElement(AiConfirmCard, {
      labels: AI_CONFIRM_LABELS.en,
      card: hook.card,
      ...(timeoutMs ? { timeoutMinutes: 1 } : {}),
      onAllow: hook.allow,
      onDeny: hook.deny,
    })
  }
  act(() => root.render(createElement(Gate)))
}

async function settle(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
  })
}

const click = (selector: string) => {
  const el = host.querySelector(selector)
  expect(el).toBeTruthy()
  act(() => (el as HTMLElement).click())
}

describe('useToolConfirm + AiConfirmCard', () => {
  it('asks, then allows: the promise resolves true and the card keeps the outcome', async () => {
    mountGate()
    let decided: boolean | undefined
    const pending = hook.confirm({ name: 'run_genoffice', input: { args: ['guide', 'docs'] } })
    void pending.then((allowed) => (decided = allowed))
    await settle()
    expect(hook.card?.outcome).toBe('awaiting')
    expect(host.querySelector('.ai-confirm-card')).toBeTruthy()

    click('.ai-confirm-allow')
    await settle()
    expect(decided).toBe(true)
    // the settled card stays in the flow as the record of the decision
    expect(hook.card?.outcome).toBe('allowed')
    expect(host.querySelector('.ai-confirm-allow')).toBeNull()
  })

  it('denying resolves false and disables further clicks (one decision per request)', async () => {
    mountGate()
    const decisions: boolean[] = []
    const pending = hook.confirm({ name: 'create_document', input: { title: 'Report' } })
    void pending.then((allowed) => decisions.push(allowed))
    await settle()
    click('.ai-confirm-deny')
    await settle()
    expect(decisions).toEqual([false])
    expect(hook.card?.outcome).toBe('denied')
    // the buttons are gone: the settled request cannot be re-decided
    expect(host.querySelector('.ai-confirm-deny')).toBeNull()
  })

  it('a click after the outcome must not double-settle the promise', async () => {
    mountGate()
    const decisions: boolean[] = []
    const pending = hook.confirm({ name: 'run_genoffice', input: {} })
    void pending.then((allowed) => decisions.push(allowed))
    await settle()
    click('.ai-confirm-deny')
    await settle()
    // late clicks (a zombie handler, a second card instance) settle nothing
    hook.deny()
    hook.allow()
    await settle()
    expect(decisions).toEqual([false])
  })

  it('silence past the timeout counts as denial', async () => {
    vi.useFakeTimers()
    try {
      mountGate(1_000)
      let decided: boolean | undefined
      const pending = hook.confirm({ name: 'run_genoffice', input: {} })
      void pending.then((allowed) => (decided = allowed))
      await settle()
      act(() => vi.advanceTimersByTime(1_001))
      await settle()
      expect(decided).toBe(false)
      expect(hook.card?.outcome).toBe('denied')
      expect(hook.card?.timedOut).toBe(true)
    } finally {
      vi.useRealTimers()
    }
  })

  it('settleAll(false) answers a pending request the safe way (Stop, thread switch)', async () => {
    mountGate()
    let decided: boolean | undefined
    const pending = hook.confirm({ name: 'run_genoffice', input: {} })
    void pending.then((allowed) => (decided = allowed))
    await settle()
    hook.settleAll(false)
    await settle()
    expect(decided).toBe(false)
    expect(hook.card?.outcome).toBe('denied')
  })

  it('unmount denies whatever is still pending, so the loop never hangs', async () => {
    mountGate()
    let decided: boolean | undefined
    const pending = hook.confirm({ name: 'run_genoffice', input: {} })
    void pending.then((allowed) => (decided = allowed))
    await settle()
    act(() => root.unmount())
    // the unmount cleanup resolves the pending promise on the next microtask
    await Promise.resolve()
    expect(decided).toBe(false)
  })

  it('a second request while one is pending settles the first as denied', async () => {
    mountGate()
    const decisions: boolean[] = []
    const first = hook.confirm({ name: 'run_genoffice', input: {} })
    void first.then((allowed) => decisions.push(allowed))
    await settle()
    const second = hook.confirm({ name: 'create_document', input: {} })
    await settle()
    expect(decisions).toEqual([false])
    // the card now shows the new request
    expect(hook.card?.toolName).toBe('create_document')
    click('.ai-confirm-allow')
    await settle()
    await expect(second).resolves.toBe(true)
  })

  it('clear drops a settled card from the flow (a new run owns the stream)', async () => {
    mountGate()
    const pending = hook.confirm({ name: 'run_genoffice', input: {} })
    await settle()
    click('.ai-confirm-deny')
    await settle()
    expect(hook.card).not.toBeNull()
    act(() => hook.clear())
    expect(hook.card).toBeNull()
    await expect(pending).resolves.toBe(false)
  })

  it('shows the tool name and a capped arguments preview', async () => {
    mountGate()
    void hook.confirm({ name: 'apply_ops', input: { ops: [{ op: 'setFont', color: '#FF0000' }] } })
    await settle()
    expect(host.querySelector('.ai-confirm-tool')?.textContent).toContain('apply_ops')
    const pre = host.querySelector('.ai-confirm-args pre')
    expect(pre?.textContent).toContain('setFont')
  })
})
