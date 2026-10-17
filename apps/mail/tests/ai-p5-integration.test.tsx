import { act } from 'react'
import type { ReactElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { AgentSkill, AgentStreamCallbacks, AgentTransport } from '@genoffice/agent-core'
import type { AgentToolCall } from '@genoffice/agent-core'
import { LocaleProvider } from '../src/renderer/i18n/locale'
import App, { createMailAiSkills } from '../src/renderer/App'
import { AiPanel, type AiPanelSkillContext } from '../src/renderer/ai/AiPanel'
import { FakeMailApi, type FakeSeed } from '../src/renderer/devFakes'

/**
 * P5 QA integration (owns this file): the three lanes wired together.
 * K skills (via the App `createMailAiSkills` glue) run over the
 * production structural path (`window.mailApi.aiRetrieve`, UI-lane fake
 * standing in for the W-lane `mail:ai-retrieve` bridge) with the stubbed
 * provider transport: triage digest on headers-only, thread summarize
 * with working citation-open, compose-assist as editable text, ask-mail
 * answers with per-claim citations, two-account no-leak, and silent logs.
 * All fixtures are synthetic example.com mail; nothing leaves the fake.
 */

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const QA_SEEDS: FakeSeed[] = [
  {
    accountId: 'qa-a',
    label: 'QA Alpha',
    folders: [{ id: 'inbox', name: 'Inbox', role: 'inbox', unread: 2 }],
    messages: [
      {
        uid: 'qa-m1',
        folderId: 'inbox',
        from: ['Acme Billing <billing@example.com>'],
        to: ['Reader <reader@example.net>'],
        subject: 'Invoice INV-1042 ready for review',
        date: '2026-10-10T09:00:00.000Z',
        snippet: 'Invoice INV-1042 for October is attached for review.',
        text: 'Hello,\n\nBODY-MARKER-A1 invoice terms, bank account 88112233, due Friday.\n\nAcme Billing',
      },
      {
        uid: 'qa-m2',
        folderId: 'inbox',
        from: ['Team Chat <chat@example.com>'],
        to: ['Reader <reader@example.net>'],
        subject: 'Weekly newsletter notes',
        date: '2026-10-11T12:30:00.000Z',
        seen: true,
        snippet: 'This week shipped the new reader with remote images blocked.',
        text: 'BODY-MARKER-A2 standup transcript with action items.',
      },
    ],
  },
  {
    accountId: 'qa-b',
    label: 'QA Beta',
    folders: [{ id: 'inbox', name: 'Inbox', role: 'inbox', unread: 1 }],
    messages: [
      {
        uid: 'qb-m1',
        folderId: 'inbox',
        from: ['Payroll <payroll@example.net>'],
        to: ['Reader <reader@example.net>'],
        subject: 'Payslip September is available',
        date: '2026-10-10T10:00:00.000Z',
        snippet: 'Your September payslip is ready to download.',
        text: 'BODY-MARKER-B1 salary breakdown and tax withholding figures.',
      },
    ],
  },
]

function seedBridge(): FakeMailApi {
  const api = new FakeMailApi(QA_SEEDS)
  api.getAiSettings = (async () => ({})) as FakeMailApi['getAiSettings']
  ;(window as unknown as { mailApi: FakeMailApi }).mailApi = api
  return api
}

function toolCall(name: string, input: Record<string, unknown> = {}): AgentToolCall {
  return { id: `qa-${name}`, name, input }
}

function ctxFor(accountIds: string[]): AiPanelSkillContext {
  return {
    getScope: () => ({ accountIds: [...accountIds] }),
    getDraft: () => '',
  }
}

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

async function flush(rounds = 30): Promise<void> {
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

function scriptedTransport(behaviors: Array<(cb: AgentStreamCallbacks) => void>): AgentTransport {
  let calls = 0
  return {
    stream: (_request, callbacks) => {
      const behavior = behaviors[Math.min(calls, behaviors.length - 1)]
      calls += 1
      queueMicrotask(() => behavior?.(callbacks))
      return { cancel: () => queueMicrotask(() => callbacks.onDone()) }
    },
  }
}

beforeEach(() => {
  seedBridge()
})

afterEach(() => {
  for (const { container, root } of containers) {
    act(() => {
      root.unmount()
    })
    container.remove()
  }
  containers = []
  vi.restoreAllMocks()
})

describe('P5 QA integration: App skill wiring', () => {
  it('exposes the four K skills with live scope readers', async () => {
    const skills = createMailAiSkills(ctxFor(['qa-a']))
    expect(skills.map((skill) => skill.id)).toEqual([
      'mail-triage',
      'mail-summarize',
      'mail-compose-assist',
      'mail-search',
    ])
    // Same instances follow scope edits: no stale capture.
    const mutable = ctxFor(['qa-a'])
    const [triage] = createMailAiSkills(mutable)
    const first = await triage.executeTool?.(toolCall('rank_folder', { query: 'invoice' }))
    expect(first?.isError).toBeFalsy()
    expect(first?.output).toContain('[qa-m1]')
    mutable.getScope = () => ({ accountIds: ['qa-b'] })
    const second = await triage.executeTool?.(toolCall('rank_folder', { query: 'invoice' }))
    expect(second?.isError).toBeFalsy()
    expect(second?.output).not.toContain('[qa-m1]')
  })

  it('mounts the panel inside App without crashing the skill loop', async () => {
    const api = new FakeMailApi(QA_SEEDS)
    api.getAiSettings = (async () => ({})) as FakeMailApi['getAiSettings']
    const container = render(<App api={api} />)
    await flush(10)
    expect(container.querySelector('.mail-ai-panel')).not.toBeNull()
    const picks = [...container.querySelectorAll('.mail-ai-pick')].map((el) => el.textContent)
    expect(picks.join(' ')).toContain('QA Alpha')
    expect(picks.join(' ')).toContain('QA Beta')
  })
})

describe('P5 QA integration: stubbed-provider AI e2e over the bridge', () => {
  it('triage digest reads headers-only and stays inside the scope', async () => {
    const [triage] = createMailAiSkills(ctxFor(['qa-a']))
    // Single-token query: the dev fake matches one substring while the
    // production lane delegates multi-token queries to FTS (retrieve.ts).
    const result = await triage.executeTool?.(toolCall('rank_folder', { query: 'invoice' }))
    expect(result?.isError).toBeFalsy()
    expect(result?.output).toContain('[qa-m1]')
    expect(result?.output).not.toContain('BODY-MARKER-A1')
    expect(result?.output).not.toContain('BODY-MARKER-B1')
    // The other account is invisible under this scope.
    const other = await triage.executeTool?.(toolCall('rank_folder', { query: 'payslip' }))
    expect(other?.isError).toBeFalsy()
    expect(other?.output).not.toContain('[qb-m1]')
    expect(other?.output).not.toContain('BODY-MARKER-B1')
  })

  it('thread summarize cites per claim and opens through the panel', async () => {
    const skills = createMailAiSkills(ctxFor(['qa-a']))
    const summarize = skills.find((skill) => skill.id === 'mail-summarize') as AgentSkill
    const result = await summarize.executeTool?.(
      toolCall('summarize_thread', { uids: ['qa-m1'], folderId: 'inbox' }),
    )
    expect(result?.isError).toBeFalsy()
    expect(result?.output).toContain('[qa-m1]')
    // Uncited factual claims force one corrective turn naming the citation.
    const correction = summarize.verifyResponse?.('- Terms are due Friday with no citation', [
      { name: 'summarize_thread', ok: true },
    ])
    expect(correction).toContain('[qa-m1]')

    // Panel streams a cited reply; the citation opens the reader path.
    const onOpenCitation = vi.fn()
    const container = render(
      <AiPanel
        accountIds={['qa-a']}
        accounts={[
          { accountId: 'qa-a', label: 'QA Alpha' },
          { accountId: 'qa-b', label: 'QA Beta' },
        ]}
        onOpenCitation={onOpenCitation}
        onInsertAssist={() => undefined}
        createSkills={createMailAiSkills}
        transport={scriptedTransport([
          (cb) => {
            cb.onToolCall({ id: 'c1', name: 'rank_folder', input: { query: 'invoice' } })
            cb.onDone()
          },
          (cb) => {
            cb.onDelta('Top item: invoice ready [qa-m1].')
            cb.onDone()
          },
        ])}
      />,
    )
    act(() => {
      buttonByText(container, 'Triage')?.click()
    })
    await flush()
    const cite = container.querySelector('.mail-ai-cite') as HTMLButtonElement | null
    expect(cite?.textContent).toBe('[qa-m1]')
    act(() => {
      cite?.click()
    })
    expect(onOpenCitation).toHaveBeenCalledWith({ uid: 'qa-m1' })
  })

  it('compose-assist inserts editable text and never sends', async () => {
    const skills = createMailAiSkills(ctxFor(['qa-a']))
    const assist = skills.find((skill) => skill.id === 'mail-compose-assist') as AgentSkill
    const result = await assist.executeTool?.(
      toolCall('draft_from_bullets', { bullets: ['Confirm receipt', 'Ask for a W-8 form'] }),
    )
    expect(result?.isError).toBeFalsy()
    expect(result?.mutated).toBeFalsy()
    expect(result?.output).toContain('- Confirm receipt')

    // The panel hands the reply to the composer; the human sends.
    const onInsertAssist = vi.fn()
    const container = render(
      <AiPanel
        accountIds={['qa-a']}
        accounts={[{ accountId: 'qa-a', label: 'QA Alpha' }]}
        onOpenCitation={() => undefined}
        onInsertAssist={onInsertAssist}
        createSkills={createMailAiSkills}
        transport={scriptedTransport([
          (cb) => {
            cb.onToolCall({
              id: 'c1',
              name: 'draft_from_bullets',
              input: { bullets: ['Confirm receipt'] },
            })
            cb.onDone()
          },
          (cb) => {
            cb.onDelta('Draft ready [qa-m1].')
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
    expect(onInsertAssist).toHaveBeenCalledWith('Draft ready [qa-m1].')
  })

  it('ask-mail answers carry per-claim citations', async () => {
    const skills = createMailAiSkills(ctxFor(['qa-a']))
    const search = skills.find((skill) => skill.id === 'mail-search') as AgentSkill
    const hits = await search.executeTool?.(toolCall('search_mail', { query: 'invoice' }))
    expect(hits?.isError).toBeFalsy()
    expect(hits?.output).toContain('[qa-m1]')
    expect(hits?.output).not.toContain('BODY-MARKER-A1')
    const quote = await search.executeTool?.(
      toolCall('quote_message', { uids: ['qa-m1'], folderId: 'inbox' }),
    )
    expect(quote?.isError).toBeFalsy()
    expect(quote?.output).toContain('[qa-m1]')
    const correction = search.verifyResponse?.('- The invoice is due Friday, no source given', [
      { name: 'quote_message', ok: true },
    ])
    expect(correction).toContain('[qa-m1]')
  })
})

describe('P5 QA integration: no cross-account leakage', () => {
  it('bridge tags every citation and body with its scope account', async () => {
    const api = new FakeMailApi(QA_SEEDS)
    const cited = await api.aiRetrieve({
      scope: { accountIds: ['qa-a'] },
      op: 'fetch-bodies',
      refs: [{ uid: 'qa-m1', folderId: 'inbox', threadId: 'thread-qa-m1' }],
    })
    expect(cited.citations).toEqual([
      { uid: 'qa-m1', folderId: 'inbox', msgIdHeader: '<qa-m1@example.com>', accountId: 'qa-a' },
    ])
    expect(cited.bodies?.map((entry) => entry.accountId)).toEqual(['qa-a'])
    const digest = await api.aiRetrieve({
      scope: { accountIds: ['qa-a'] },
      op: 'triage-headers',
      query: 'invoice',
    })
    expect(digest.citations.every((citation) => citation.accountId === 'qa-a')).toBe(true)
  })
  it('wrong-account bodies never surface under a foreign scope', async () => {
    const skills = createMailAiSkills(ctxFor(['qa-a']))
    const summarize = skills.find((skill) => skill.id === 'mail-summarize') as AgentSkill
    const search = skills.find((skill) => skill.id === 'mail-search') as AgentSkill
    const leak = await summarize.executeTool?.(
      toolCall('summarize_thread', { uids: ['qb-m1'], folderId: 'inbox' }),
    )
    expect(leak?.isError).toBe(true)
    expect(leak?.output).not.toContain('BODY-MARKER-B1')
    const quoteLeak = await search.executeTool?.(
      toolCall('quote_message', { uids: ['qb-m1'], folderId: 'inbox' }),
    )
    expect(quoteLeak?.isError).toBe(true)
    expect(quoteLeak?.output).not.toContain('BODY-MARKER-B1')
    // Unknown accounts are rejected before any store opens.
    const wide = createMailAiSkills(ctxFor(['ghost-acct']))
    const [triage] = wide
    const rejected = await triage.executeTool?.(toolCall('rank_folder', { query: 'invoice' }))
    expect(rejected?.isError).toBe(true)
    expect(rejected?.output).toMatch(/unknown account|scope/i)
  })
})

describe('P5 QA integration: redaction and log silence', () => {
  it('skill execution logs no mail content', async () => {
    const seen: string[] = []
    for (const method of ['log', 'warn', 'error', 'debug'] as const) {
      vi.spyOn(console, method).mockImplementation((...args: unknown[]) => {
        seen.push(args.map((arg) => String(arg)).join(' '))
      })
    }
    const skills = createMailAiSkills(ctxFor(['qa-a']))
    for (const skill of skills) {
      for (const tool of skill.tools) {
        if (tool.name === 'rank_folder') {
          await skill.executeTool?.(toolCall(tool.name, { query: 'invoice' }))
        } else if (tool.name === 'summarize_thread' || tool.name === 'quote_message') {
          await skill.executeTool?.(toolCall(tool.name, { uids: ['qa-m1'], folderId: 'inbox' }))
        } else if (tool.name === 'search_mail') {
          await skill.executeTool?.(toolCall(tool.name, { query: 'invoice' }))
        } else if (tool.name === 'draft_from_bullets') {
          await skill.executeTool?.(toolCall(tool.name, { bullets: ['Confirm receipt'] }))
        } else if (tool.name === 'rewrite_tone') {
          await skill.executeTool?.(toolCall(tool.name, { text: 'Send it now.', tone: 'formal' }))
        } else if (tool.name === 'reply_from_thread') {
          await skill.executeTool?.(
            toolCall(tool.name, { uids: ['qa-m1'], folderId: 'inbox', focus: 'payment date' }),
          )
        } else if (tool.name === 'apply_label') {
          await skill.executeTool?.(
            toolCall(tool.name, { uids: ['qa-m1'], folderId: 'inbox', label: 'follow-up' }),
          )
        } else if (tool.name === 'catch_up_folder') {
          await skill.executeTool?.(
            toolCall(tool.name, { query: 'invoice', folderId: 'inbox', limit: 5 }),
          )
        }
      }
    }
    const dump = seen.join('\n')
    for (const marker of [
      'BODY-MARKER-A1',
      'BODY-MARKER-A2',
      'BODY-MARKER-B1',
      'INV-1042',
      'billing@example.com',
      '88112233',
    ]) {
      expect(dump).not.toContain(marker)
    }
  })
})
