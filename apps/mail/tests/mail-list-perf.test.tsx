import { act } from 'react'
import type { ReactElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { MailEnvelope } from '@genoffice/mail-protocols'
import { LocaleProvider } from '../src/renderer/i18n/locale'
import { FakeMailApi } from '../src/renderer/devFakes'
import { MessageList } from '../src/renderer/list/MessageList'
import { SearchBox } from '../src/renderer/list/SearchBox'
import { MessageReader } from '../src/renderer/reader/MessageReader'
import { sanitizeMailHtml } from '../src/renderer/reader/quarantine'
import type { MailApi } from '../src/shared/ipc'
import {
  LIST_OVERSCAN,
  LIST_ROW_HEIGHT,
  LIST_VIEWPORT_HEIGHT,
  SEARCH_DEBOUNCE_MAX_MS,
  SEARCH_DEBOUNCE_MIN_MS,
  SEARCH_DEBOUNCE_MS,
  DebouncedSearch,
  SearchQueryTracker,
  computeListWindow,
  maxWindowedRows,
  percentile,
} from '../src/renderer/list/list-perf'

/**
 * H6 list-side perf budget pins (H1 hardening spec section 4.1,
 * P-1..P-5). Synthetic `example.com` envelopes only; no real mail.
 * P-1..P-4 are CI pins; P-5 records harness numbers against the
 * reported (non-CI-gating) budgets with a generous anti-regression
 * ceiling so jsdom timing jitter never flakes the suite.
 */

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const MAILBOX_SIZE = 100_000

function syntheticMailbox(total: number): MailEnvelope[] {
  const out = new Array<MailEnvelope>(total)
  for (let i = 0; i < total; i += 1) {
    const uid = `uid-${i}`
    out[i] = {
      from: [`sender${i % 997}@example.com`],
      to: ['reader@example.com'],
      cc: [],
      subject: `Subject ${i}`,
      date: '2026-10-14T00:00:00.000Z',
      labels: [],
      snippet: `snippet ${i}`,
      hasAttach: false,
      ref: { uid, folderId: 'inbox', threadId: `thread-${i % 1000}` },
      flags: { seen: i % 2 === 0, flagged: false, answered: false, draft: false },
    }
  }
  return out
}

let containers: HTMLElement[] = []

function render(node: ReactElement): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  containers.push(container)
  const root = createRoot(container)
  act(() => {
    root.render(<LocaleProvider initial="en">{node}</LocaleProvider>)
  })
  return container
}

function renderRoot(node: ReactElement): { container: HTMLElement; root: Root } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  containers.push(container)
  const root = createRoot(container)
  act(() => {
    root.render(<LocaleProvider initial="en">{node}</LocaleProvider>)
  })
  return { container, root }
}

async function settle(rounds = 8): Promise<void> {
  await act(async () => {
    for (let i = 0; i < rounds; i += 1) {
      await Promise.resolve()
    }
  })
}

function scrollListTo(container: HTMLElement, top: number): void {
  const list = container.querySelector('.mail-list') as HTMLElement
  list.scrollTop = top
  act(() => {
    list.dispatchEvent(new Event('scroll', { bubbles: true }))
  })
}

function uidIndex(uid: string): number {
  return Number.parseInt(uid.slice('uid-'.length), 10)
}

beforeEach(() => {
  vi.useRealTimers()
  ;(window as unknown as { mailApi: FakeMailApi }).mailApi = new FakeMailApi()
})

afterEach(() => {
  for (const container of containers) container.remove()
  containers = []
  vi.useRealTimers()
  vi.restoreAllMocks()
})

describe('P-1 window bound at 100k envelopes', () => {
  it('pins the spec-default bound to ~15 mounted rows', () => {
    expect(LIST_ROW_HEIGHT).toBe(76)
    expect(LIST_OVERSCAN).toBe(4)
    expect(LIST_VIEWPORT_HEIGHT).toBe(480)
    expect(maxWindowedRows(LIST_VIEWPORT_HEIGHT, LIST_ROW_HEIGHT, LIST_OVERSCAN)).toBe(15)
  })

  it('holds the bound at every scroll position with consistent spacers', () => {
    const totalHeight = MAILBOX_SIZE * LIST_ROW_HEIGHT
    const bound = maxWindowedRows(LIST_VIEWPORT_HEIGHT, LIST_ROW_HEIGHT, LIST_OVERSCAN)
    for (let top = 0; top <= totalHeight; top += 76_913) {
      const window = computeListWindow({
        total: MAILBOX_SIZE,
        scrollTop: top,
        viewportHeight: LIST_VIEWPORT_HEIGHT,
        rowHeight: LIST_ROW_HEIGHT,
        overscan: LIST_OVERSCAN,
      })
      expect(window.endIndex - window.startIndex).toBeLessThanOrEqual(bound)
      expect(window.startIndex).toBeGreaterThanOrEqual(0)
      expect(window.endIndex).toBeLessThanOrEqual(MAILBOX_SIZE)
      expect(
        window.topSpacer +
          (window.endIndex - window.startIndex) * LIST_ROW_HEIGHT +
          window.bottomSpacer,
      ).toBe(totalHeight)
    }
    // Bottom edge clamps exactly onto the tail rows.
    const bottom = computeListWindow({
      total: MAILBOX_SIZE,
      scrollTop: totalHeight - LIST_VIEWPORT_HEIGHT,
      viewportHeight: LIST_VIEWPORT_HEIGHT,
      rowHeight: LIST_ROW_HEIGHT,
      overscan: LIST_OVERSCAN,
    })
    expect(bottom.endIndex).toBe(MAILBOX_SIZE)
    expect(bottom.bottomSpacer).toBe(0)
  })
})

describe('P-2 scroll-handler cost at 100k envelopes', () => {
  it('runs the slice + spacer handler p95 under 8 ms', () => {
    const box = syntheticMailbox(MAILBOX_SIZE)
    const samples: number[] = []
    for (let i = 0; i < 500; i += 1) {
      const top = (i * 7919) % (MAILBOX_SIZE * LIST_ROW_HEIGHT)
      const started = performance.now()
      const window = computeListWindow({
        total: box.length,
        scrollTop: top,
        viewportHeight: LIST_VIEWPORT_HEIGHT,
        rowHeight: LIST_ROW_HEIGHT,
        overscan: LIST_OVERSCAN,
      })
      void box.slice(window.startIndex, window.endIndex)
      samples.push(performance.now() - started)
    }
    const p95 = percentile(samples, 95)
    const mean = samples.reduce((sum, value) => sum + value, 0) / samples.length
    console.info(
      `[h6-perf] P-2 scroll handler over 100k: mean ${mean.toFixed(3)} ms, p95 ${p95.toFixed(3)} ms (budget p95 < 8 ms)`,
    )
    expect(p95).toBeLessThan(8)
  })

  it('never renders off-window rows (render-counter probe)', () => {
    const box = syntheticMailbox(MAILBOX_SIZE)
    let rendered: string[] = []
    const container = render(
      <MessageList
        messages={box}
        selectedUid={null}
        loading={false}
        hasMore={false}
        onSelect={() => undefined}
        onLoadMore={() => undefined}
        onRenderRow={(uid) => {
          rendered.push(uid)
        }}
      />,
    )
    const bound = maxWindowedRows(LIST_VIEWPORT_HEIGHT, LIST_ROW_HEIGHT, LIST_OVERSCAN)
    expect(rendered.length).toBeLessThanOrEqual(bound)
    expect(container.querySelectorAll('.mail-row').length).toBeLessThanOrEqual(bound)

    rendered = []
    const targetTop = 3_800_000
    scrollListTo(container, targetTop)
    const expectedStart = Math.floor(targetTop / LIST_ROW_HEIGHT) - LIST_OVERSCAN
    expect(container.querySelectorAll('.mail-row').length).toBeLessThanOrEqual(bound)
    expect(rendered.length).toBeGreaterThan(0)
    for (const uid of rendered) {
      const index = uidIndex(uid)
      expect(index).toBeGreaterThanOrEqual(expectedStart)
      expect(index).toBeLessThan(expectedStart + bound)
    }
  })
})

describe('P-3 initial mount at 100k envelopes', () => {
  it('first-paints with at most 30 mounted row nodes', () => {
    const box = syntheticMailbox(MAILBOX_SIZE)
    const container = render(
      <MessageList
        messages={box}
        selectedUid={null}
        loading={false}
        hasMore={false}
        onSelect={() => undefined}
        onLoadMore={() => undefined}
      />,
    )
    const rows = container.querySelectorAll('.mail-row').length
    console.info(`[h6-perf] P-3 initial mount over 100k: ${rows} row nodes (budget <= 30)`)
    expect(rows).toBeLessThanOrEqual(30)
  })
})

describe('P-4 search-as-you-type debounce', () => {
  it('pins the single debounce constant inside 150-250 ms', () => {
    expect(SEARCH_DEBOUNCE_MIN_MS).toBe(150)
    expect(SEARCH_DEBOUNCE_MAX_MS).toBe(250)
    expect(SEARCH_DEBOUNCE_MS).toBeGreaterThanOrEqual(SEARCH_DEBOUNCE_MIN_MS)
    expect(SEARCH_DEBOUNCE_MS).toBeLessThanOrEqual(SEARCH_DEBOUNCE_MAX_MS)
  })

  it('fires once with the latest query; keystrokes supersede the pending query', () => {
    vi.useFakeTimers()
    try {
      const fired: Array<[string, number]> = []
      let superseded = 0
      const debounced = new DebouncedSearch(
        SEARCH_DEBOUNCE_MS,
        (query, generation) => {
          fired.push([query, generation])
        },
        () => {
          superseded += 1
        },
      )
      debounced.push('a')
      debounced.push('ab')
      debounced.push('abc')
      expect(debounced.pending).toBe(true)
      vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS - 1)
      expect(fired).toEqual([])
      vi.advanceTimersByTime(1)
      expect(fired).toEqual([['abc', 3]])
      expect(superseded).toBe(2)
    } finally {
      vi.useRealTimers()
    }
  })

  it('suppresses refires for unchanged values and cancels cleanly', () => {
    vi.useFakeTimers()
    try {
      const fired: string[] = []
      let superseded = 0
      const debounced = new DebouncedSearch(
        SEARCH_DEBOUNCE_MS,
        (query) => {
          fired.push(query)
        },
        () => {
          superseded += 1
        },
      )
      debounced.push('abc')
      vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS)
      expect(fired).toEqual(['abc'])
      debounced.push('abc')
      vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS)
      expect(fired).toEqual(['abc'])
      debounced.push('abcd')
      debounced.cancel()
      vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS * 3)
      expect(fired).toEqual(['abc'])
      expect(superseded).toBeGreaterThanOrEqual(1)
    } finally {
      vi.useRealTimers()
    }
  })

  it('discards stale async results via the generation tracker', () => {
    const tracker = new SearchQueryTracker()
    const first = tracker.start()
    const second = tracker.start()
    expect(tracker.isCurrent(first)).toBe(false)
    expect(tracker.isCurrent(second)).toBe(true)
  })

  it('keeps legacy submit-only behavior when no debounce is set', () => {
    let submitted = 0
    const container = render(
      <SearchBox
        value="abc"
        searching={false}
        onChange={() => undefined}
        onSubmit={() => {
          submitted += 1
        }}
        onClear={() => undefined}
      />,
    )
    const form = container.querySelector('.mail-search') as HTMLFormElement
    act(() => {
      form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }))
    })
    expect(submitted).toBe(1)
  })

  it('debounces keystrokes through the box and reports supersedes', () => {
    vi.useFakeTimers()
    try {
      const fired: string[] = []
      let superseded = 0
      const props = {
        searching: false,
        onChange: () => undefined,
        onSubmit: () => undefined,
        onClear: () => undefined,
        debounceMs: SEARCH_DEBOUNCE_MS,
        onDebouncedQuery: (query: string) => {
          fired.push(query)
        },
        onQuerySuperseded: () => {
          superseded += 1
        },
      }
      const { root } = renderRoot(<SearchBox value="a" {...props} />)
      act(() => {
        root.render(
          <LocaleProvider initial="en">
            <SearchBox value="ab" {...props} />
          </LocaleProvider>,
        )
      })
      act(() => {
        root.render(
          <LocaleProvider initial="en">
            <SearchBox value="abc" {...props} />
          </LocaleProvider>,
        )
      })
      act(() => {
        vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS)
      })
      expect(fired).toEqual(['abc'])
      expect(superseded).toBe(2)
    } finally {
      vi.useRealTimers()
    }
  })
})

describe('P-5 harness targets (reported numbers)', () => {
  it('opens a 1 MiB reader body with sanitize + first paint measured', async () => {
    const bodyText = `Synthetic line from sender@example.com for scale testing. `.repeat(18_204)
    expect(bodyText.length).toBeGreaterThan(1_024 * 1_024)
    const sanitizeStart = performance.now()
    const sanitized = sanitizeMailHtml(`<pre>${bodyText}</pre>`, { blockedLabel: 'Blocked' })
    const sanitizeMs = performance.now() - sanitizeStart
    expect(sanitized.html).toContain('Synthetic line')

    const envelope = syntheticMailbox(1)[0]
    const api = { fetchBody: async () => ({ text: bodyText }) } as unknown as MailApi
    const paintStart = performance.now()
    const container = render(<MessageReader api={api} envelope={envelope} />)
    await settle()
    const firstPaintMs = performance.now() - paintStart
    console.info(
      `[h6-perf] P-5 reader open of a 1 MiB body: sanitize ${sanitizeMs.toFixed(1)} ms, first paint ${firstPaintMs.toFixed(1)} ms (budget < 200 ms)`,
    )
    expect(container.querySelector('.mq-body')).not.toBeNull()
    // Anti-hang guard only (not a budget signal): the 200 ms P-5
    // budget is verified against the reported numbers on reference
    // hardware, not gated in jsdom where timer jitter dominates.
    // Spot-check 2026-10-15 on this machine's Chromium: DOMParser
    // parse + innerHTML set + tree walk + serialize of the same ~1 MiB
    // payload takes ~8-9 ms, so the sanitize path sits ~20x inside
    // the 200 ms budget on a real engine; the ~1-5 s jsdom figure is
    // harness overhead (worse under parallel workers).
    expect(firstPaintMs).toBeLessThan(15000)
  })
})
