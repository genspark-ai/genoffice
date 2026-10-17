import { act } from 'react'
import type { ReactElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { AgentSkill, AgentStreamCallbacks, AgentTransport } from '@genoffice/agent-core'
import { LocaleProvider } from '../src/renderer/i18n/locale'
import { FakeMailApi } from '../src/renderer/devFakes'
import type { AiPanelSkillContext } from '../src/renderer/ai/AiPanel'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import { ComposeHost } from '../src/renderer/compose/ComposeHost'

/**
 * UI-lane AI panel tests over stub transports and the devFakes retrieval
 * fake: account-picker scoping, cited replies opening the reader path,
 * compose-assist insertion landing editable, error surfacing, and the
 * spec-faithful aiRetrieve fake (scope echo, cited bodies, rejections).
 * All fixtures are synthetic example.com mail; nothing leaves the fake.
 */

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

let containers: Array<{ container: HTMLElement; root: Root }> = []

function render(node: ReactElement): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  containers.push({ container, root })
  act(() => {
    root.render(<LocaleProvider initial="en">{node}</LocaleProvider>)
  })
  return container
}

async function flush(rounds = 20): Promise<void> {
  await act(async () => {
    for (let i = 0; i < rounds; i += 1) {
      await new Promise((resolve) => setTimeout(resolve, 0))
    }
  })
}

function buttonByText(container: HTMLElement, text: string): HTMLButtonElement | undefined {
  return [...container.querySelectorAll('button')].find((button) => button.textContent === text) as
    HTMLButtonElement | undefined
}

function stubBridge(): FakeMailApi {
  const api = new FakeMailApi()
  api.getAiSettings = (async () => ({})) as FakeMailApi['getAiSettings']
  ;(window as unknown as { mailApi: FakeMailApi }).mailApi = api
  return api
}

function scriptedTransport(
  behaviors: Array<(cb: AgentStreamCallbacks) => void>,
  onCancel?: () => void,
): AgentTransport {
  let calls = 0
  return {
    stream: (_request, callbacks) => {
      const behavior = behaviors[Math.min(calls, behaviors.length - 1)]
      calls += 1
      queueMicrotask(() => behavior?.(callbacks))
      return {
        cancel: () => {
          onCancel?.()
          queueMicrotask(() => callbacks.onDone())
        },
      }
    },
  }
}

const DEMO_ACCOUNTS = [
  { accountId: 'demo', label: 'Demo Account' },
  { accountId: 'work', label: 'Work Account' },
]

beforeEach(() => {
  stubBridge()
})

afterEach(() => {
  for (const { container, root } of containers) {
    act(() => {
      root.unmount()
    })
    container.remove()
  }
  containers = []
})

describe('AiPanel', () => {
  it('renders the picker, quick actions, and composer', () => {
    const container = render(
      <AiPanel
        accountIds={['demo']}
        accounts={DEMO_ACCOUNTS}
        onOpenCitation={() => undefined}
        onInsertAssist={() => undefined}
        transport={scriptedTransport([])}
      />,
    )
    expect(container.querySelector('.mail-ai-panel')?.getAttribute('data-scope')).toBe('demo')
    expect(container.querySelector('.mail-ai-panel-title')?.textContent).toBe('Mail assistant')
    expect(buttonByText(container, 'Triage')).toBeDefined()
    expect(buttonByText(container, 'Summarize')).toBeDefined()
    expect(buttonByText(container, 'Writing help')).toBeDefined()
    expect(buttonByText(container, 'Ask about this mail')).toBeDefined()
    const box = container.querySelector('textarea')
    expect(box?.getAttribute('placeholder')).toBe('Ask about this mailbox…')
  })

  it('streams a reply and opens citations through onOpenCitation', async () => {
    const onOpenCitation = vi.fn()
    const container = render(
      <AiPanel
        accountIds={['demo']}
        accounts={DEMO_ACCOUNTS}
        onOpenCitation={onOpenCitation}
        onInsertAssist={() => undefined}
        transport={scriptedTransport([
          (cb) => {
            cb.onDelta('Two unpaid invoices need a reply [m1] and [m2].')
            cb.onDone()
          },
        ])}
      />,
    )
    act(() => {
      buttonByText(container, 'Triage')?.click()
    })
    await flush()
    const cites = [...container.querySelectorAll('.mail-ai-cite')]
    expect(cites.map((cite) => cite.textContent)).toEqual(['[m1]', '[m2]'])
    act(() => {
      ;(cites[0] as HTMLButtonElement).click()
    })
    expect(onOpenCitation).toHaveBeenCalledWith({ uid: 'm1' })
  })

  it('inserts the reply into compose through onInsertAssist', async () => {
    const onInsertAssist = vi.fn()
    const container = render(
      <AiPanel
        accountIds={['demo']}
        accounts={DEMO_ACCOUNTS}
        onOpenCitation={() => undefined}
        onInsertAssist={onInsertAssist}
        transport={scriptedTransport([
          (cb) => {
            cb.onDelta('Draft ready [m1].')
            cb.onDone()
          },
        ])}
      />,
    )
    act(() => {
      buttonByText(container, 'Writing help')?.click()
    })
    await flush()
    const insert = [...container.querySelectorAll('.mail-ai-msg-actions button')].at(-1)
    expect(insert).toBeDefined()
    act(() => {
      ;(insert as HTMLButtonElement).click()
    })
    expect(onInsertAssist).toHaveBeenCalledWith('Draft ready [m1].')
  })

  it('reads the live picker scope on every skill call', async () => {
    const seenScopes: string[][] = []
    const captured: { current: AiPanelSkillContext | null } = { current: null }
    const createSkills = (ctx: AiPanelSkillContext): AgentSkill[] => {
      captured.current = ctx
      return [
        {
          id: 'test-mail',
          systemPrompt: 'test skill',
          tools: [{ name: 'rank_folder', description: 'rank', inputSchema: { type: 'object' } }],
          executeTool: () => {
            seenScopes.push(ctx.getScope().accountIds)
            return { output: 'ranked [m1]', summary: 'ranked' }
          },
        },
      ]
    }
    const container = render(
      <AiPanel
        accountIds={['demo']}
        accounts={DEMO_ACCOUNTS}
        onOpenCitation={() => undefined}
        onInsertAssist={() => undefined}
        createSkills={createSkills}
        transport={scriptedTransport([
          (cb) => {
            cb.onToolCall({ id: 'c1', name: 'rank_folder', input: {} })
            cb.onDone()
          },
          (cb) => {
            cb.onDelta('Ranked [m1].')
            cb.onDone()
          },
        ])}
      />,
    )
    expect(captured.current).not.toBeNull()
    // The picker follows the prop scope, then tracks user edits live.
    expect(captured.current?.getScope()).toEqual({ accountIds: ['demo'] })
    const workBox = [...container.querySelectorAll('.mail-ai-pick input')][1] as HTMLInputElement
    act(() => {
      workBox.click()
    })
    expect(container.querySelector('.mail-ai-panel')?.getAttribute('data-scope')).toBe('demo,work')
    expect(captured.current?.getScope()).toEqual({ accountIds: ['demo', 'work'] })
    act(() => {
      buttonByText(container, 'Triage')?.click()
    })
    await flush()
    expect(seenScopes).toEqual([['demo', 'work']])
  })

  it('keeps at least one account picked', () => {
    const container = render(
      <AiPanel
        accountIds={['demo']}
        accounts={DEMO_ACCOUNTS}
        onOpenCitation={() => undefined}
        onInsertAssist={() => undefined}
        transport={scriptedTransport([])}
      />,
    )
    const demoBox = [...container.querySelectorAll('.mail-ai-pick input')][0] as HTMLInputElement
    act(() => {
      demoBox.click()
    })
    // Unchecking the last account is refused: the scope never empties.
    expect(demoBox.checked).toBe(true)
    expect(container.querySelector('.mail-ai-panel')?.getAttribute('data-scope')).toBe('demo')
  })

  it('surfaces transport errors with the mail.ai.error text', async () => {
    const container = render(
      <AiPanel
        accountIds={['demo']}
        accounts={DEMO_ACCOUNTS}
        onOpenCitation={() => undefined}
        onInsertAssist={() => undefined}
        transport={scriptedTransport([
          (cb) => {
            cb.onError('The assistant request failed.')
          },
        ])}
      />,
    )
    act(() => {
      buttonByText(container, 'Summarize')?.click()
    })
    await flush()
    expect(container.querySelector('.mail-ai-msg-error')?.textContent).toContain(
      'The assistant request failed.',
    )
  })

  it('stops the in-flight run through the composer stop button', async () => {
    const onCancel = vi.fn()
    const container = render(
      <AiPanel
        accountIds={['demo']}
        accounts={DEMO_ACCOUNTS}
        onOpenCitation={() => undefined}
        onInsertAssist={() => undefined}
        transport={scriptedTransport([], onCancel)}
      />,
    )
    act(() => {
      buttonByText(container, 'Triage')?.click()
    })
    await flush(2)
    const stop = container.querySelector('button[aria-label="Stop"]') as HTMLButtonElement | null
    expect(stop).not.toBeNull()
    act(() => {
      stop?.click()
    })
    await flush()
    expect(onCancel).toHaveBeenCalled()
  })
})

describe('ComposeHost assist insertion', () => {
  it('appends assist text editable and consumes it once', async () => {
    stubBridge()
    const onAssistConsumed = vi.fn()
    const container = render(
      <ComposeHost
        accountId="demo"
        onDirtyChange={() => undefined}
        onSend={() => undefined}
        assistText={'Thanks — paid this morning.\nReceipt appreciated.'}
        onAssistConsumed={onAssistConsumed}
      />,
    )
    await flush()
    const body = container.querySelector('.mail-compose-body')
    expect(body?.innerHTML).toContain('Thanks — paid this morning.')
    expect(onAssistConsumed).toHaveBeenCalledTimes(1)
  })
})

describe('FakeMailApi aiRetrieve', () => {
  it('serves triage envelopes with a scope echo and no bodies', async () => {
    const api = new FakeMailApi()
    const res = await api.aiRetrieve({
      scope: { accountIds: ['demo'] },
      op: 'triage-headers',
      query: 'invoice',
    })
    expect(res.scope).toEqual({ accountIds: ['demo'] })
    expect(res.envelopes.length).toBeGreaterThan(0)
    expect(res.envelopes.every((env) => env.ref.uid.length > 0)).toBe(true)
    expect(res.bodies).toBeUndefined()
    expect(res.citations[0]?.uid).toBe(res.envelopes[0]?.ref.uid)
  })

  it('serves cited bodies on fetch-bodies and headers on cite-resolve', async () => {
    const api = new FakeMailApi()
    const ref = { uid: 'm1', folderId: 'inbox', threadId: 'thread-m1' }
    const bodies = await api.aiRetrieve({
      scope: { accountIds: ['demo'] },
      op: 'fetch-bodies',
      refs: [ref],
    })
    expect(bodies.bodies?.[0]?.text).toContain('INV-1042')
    const resolved = await api.aiRetrieve({
      scope: { accountIds: ['demo'] },
      op: 'cite-resolve',
      refs: [ref],
    })
    expect(resolved.bodies).toBeUndefined()
    expect(resolved.citations).toEqual([
      { uid: 'm1', folderId: 'inbox', msgIdHeader: '<m1@example.com>', accountId: 'demo' },
    ])
  })

  it('rejects unknown accounts, empty scopes, and missing params', async () => {
    const api = new FakeMailApi()
    await expect(
      api.aiRetrieve({ scope: { accountIds: ['ghost'] }, op: 'triage-headers', query: 'x' }),
    ).rejects.toThrow('unknown account')
    await expect(
      api.aiRetrieve({ scope: { accountIds: [] }, op: 'triage-headers', query: 'x' }),
    ).rejects.toThrow()
    await expect(
      api.aiRetrieve({ scope: { accountIds: ['demo'] }, op: 'triage-headers' }),
    ).rejects.toThrow('requires a query')
    await expect(
      api.aiRetrieve({ scope: { accountIds: ['demo'] }, op: 'fetch-bodies' }),
    ).rejects.toThrow('requires refs')
  })
})
