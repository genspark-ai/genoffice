/** @vitest-environment jsdom */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { AiQueueStrip } from '../src/AiQueueStrip'
import { AI_QUEUE_LABELS } from '../src/strings-ai-queue'

let host: HTMLDivElement
let root: ReturnType<typeof createRoot>
beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})
afterEach(() => {
  act(() => root.unmount())
  host.remove()
  vi.restoreAllMocks()
})

const labels = AI_QUEUE_LABELS.en

it('renders nothing for an empty queue', () => {
  act(() =>
    root.render(
      createElement(AiQueueStrip, {
        items: [],
        labels,
        onUpdate: () => {},
        onRemove: () => {},
        onClear: () => {},
      }),
    ),
  )
  expect(host.querySelector('.ai-queue-strip')).toBeNull()
})

it('shows the collapsed count, expands to editable rows, and wires every action', () => {
  const onUpdate = vi.fn()
  const onRemove = vi.fn()
  const onClear = vi.fn()
  const items = [
    { id: 'q1', text: 'first message' },
    { id: 'q2', text: 'second message' },
  ]
  act(() =>
    root.render(createElement(AiQueueStrip, { items, labels, onUpdate, onRemove, onClear })),
  )
  expect(host.querySelector('.ai-queue-toggle')!.textContent).toContain('2 queued')
  // collapsed: no rows until the toggle opens
  expect(host.querySelectorAll('.ai-queue-item')).toHaveLength(0)
  act(() => host.querySelector<HTMLButtonElement>('.ai-queue-toggle')!.click())
  expect(host.querySelectorAll('.ai-queue-item')).toHaveLength(2)

  // edit: pencil fills the inline input, Enter commits the trimmed text
  const row = host.querySelectorAll<HTMLLIElement>('.ai-queue-item')[1]!
  act(() => row.querySelectorAll('button')[0]!.click())
  const input = row.querySelector<HTMLInputElement>('.ai-queue-edit')!
  expect(input.value).toBe('second message')
  // React tracks the value through its own setter: route the change through it
  const valueSetter = Object.getOwnPropertyDescriptor(
    window.HTMLInputElement.prototype,
    'value',
  )!.set!
  act(() => {
    valueSetter.call(input, '  second, edited ')
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
  })
  expect(onUpdate).toHaveBeenCalledWith('q2', 'second, edited')

  // Esc cancels the edit
  act(() => row.querySelectorAll('button')[0]!.click())
  const input2 = row.querySelector<HTMLInputElement>('.ai-queue-edit')!
  act(() =>
    input2.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true })),
  )
  expect(row.querySelector('.ai-queue-edit')).toBeNull()

  // remove the row, then clear the whole queue
  const secondRow = host.querySelectorAll<HTMLLIElement>('.ai-queue-item')[1]!
  act(() => secondRow.querySelectorAll('button')[1]!.click())
  expect(onRemove).toHaveBeenCalledWith('q2')
  act(() => host.querySelector<HTMLButtonElement>('.ai-queue-clear')!.click())
  expect(onClear).toHaveBeenCalledTimes(1)
})

it('gives the queue a pause button that flips to resume and shows the held state', () => {
  const onTogglePause = vi.fn()
  const items = [
    { id: 'q1', text: 'first' },
    { id: 'q2', text: 'second' },
  ]
  const renderAt = (paused: boolean) =>
    act(() =>
      root.render(
        createElement(AiQueueStrip, {
          items,
          labels,
          paused,
          onTogglePause,
          onUpdate: () => {},
          onRemove: () => {},
          onClear: () => {},
        }),
      ),
    )
  renderAt(false)
  const pause = host.querySelector<HTMLButtonElement>('.ai-queue-pause')!
  expect(pause.getAttribute('aria-pressed')).toBe('false')
  expect(pause.getAttribute('aria-label')).toBe(labels.pauseTitle)
  expect(host.querySelector('.ai-queue-hint')).toBeNull()
  act(() => pause.click())
  expect(onTogglePause).toHaveBeenCalledTimes(1)

  renderAt(true)
  const resume = host.querySelector<HTMLButtonElement>('.ai-queue-pause')!
  expect(resume.getAttribute('aria-pressed')).toBe('true')
  expect(resume.getAttribute('aria-label')).toBe(labels.resumeTitle)
  expect(host.querySelector('.ai-queue-hint')!.textContent).toBe(labels.pausedHint)
  expect(host.querySelector('.ai-queue-strip')!.getAttribute('data-paused')).toBe('true')
  // both messages are still there: pausing holds them, it does not drop them
  act(() => host.querySelector<HTMLButtonElement>('.ai-queue-toggle')!.click())
  expect(host.querySelectorAll('.ai-queue-item')).toHaveLength(2)
})

/** jsdom lays everything out at 0: give the rows the heights the drag maths reads */
function stubRowRects(rows: ArrayLike<Element>) {
  Array.from(rows).forEach((el, i) => {
    el.getBoundingClientRect = () =>
      ({
        top: i * 20,
        bottom: i * 20 + 20,
        height: 20,
        left: 0,
        right: 100,
        width: 100,
        x: 0,
        y: i * 20,
      }) as DOMRect
  })
}

function renderStrip(props: Partial<Parameters<typeof AiQueueStrip>[0]> = {}) {
  const onMove = vi.fn()
  const items = [
    { id: 'q1', text: 'first' },
    { id: 'q2', text: 'second' },
    { id: 'q3', text: 'third' },
  ]
  act(() =>
    root.render(
      createElement(AiQueueStrip, {
        items,
        labels,
        onUpdate: () => {},
        onMove,
        onRemove: () => {},
        onClear: () => {},
        ...props,
      }),
    ),
  )
  act(() => host.querySelector<HTMLButtonElement>('.ai-queue-toggle')!.click())
  stubRowRects(host.querySelectorAll('li.ai-queue-item'))
  return { onMove }
}

it('picks a row up on a long press and drops it where the pointer is', () => {
  vi.useFakeTimers()
  try {
    const { onMove } = renderStrip()
    const rows = host.querySelectorAll<HTMLLIElement>('li.ai-queue-item')
    act(() => {
      rows[2]!.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, clientY: 50 }))
    })
    // still held: the press has not become a drag yet
    act(() => {
      vi.advanceTimersByTime(120)
    })
    expect(host.querySelector('li[data-dragging]')).toBeNull()
    act(() => {
      vi.advanceTimersByTime(150)
    })
    expect(host.querySelector('li[data-dragging]')!.textContent).toContain('third')
    // dragged over the first row
    act(() => {
      window.dispatchEvent(new MouseEvent('pointermove', { clientY: 5 }))
    })
    act(() => {
      window.dispatchEvent(new MouseEvent('pointerup'))
    })
    expect(onMove).toHaveBeenCalledWith('q3', 0)
    expect(host.querySelector('li[data-dragging]')).toBeNull()
  } finally {
    vi.useRealTimers()
  }
})

it('does not start a drag for a click or a scroll', () => {
  vi.useFakeTimers()
  try {
    const { onMove } = renderStrip()
    const rows = host.querySelectorAll<HTMLLIElement>('li.ai-queue-item')
    // a short press: released before the hold completes
    act(() => {
      rows[0]!.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, clientY: 10 }))
    })
    act(() => {
      window.dispatchEvent(new MouseEvent('pointerup'))
    })
    act(() => {
      vi.advanceTimersByTime(500)
    })
    expect(onMove).not.toHaveBeenCalled()

    // a press that moves before the hold completes is a scroll, not a drag
    act(() => {
      rows[1]!.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, clientY: 30 }))
    })
    act(() => {
      window.dispatchEvent(new MouseEvent('pointermove', { clientY: 60 }))
    })
    act(() => {
      vi.advanceTimersByTime(500)
    })
    expect(host.querySelector('li[data-dragging]')).toBeNull()
    expect(onMove).not.toHaveBeenCalled()
  } finally {
    vi.useRealTimers()
  }
})

it('moves a row with Alt+arrow keys for keyboards that cannot drag', () => {
  const { onMove } = renderStrip()
  const rows = host.querySelectorAll<HTMLLIElement>('li.ai-queue-item')
  act(() => {
    rows[2]!.dispatchEvent(
      new window.KeyboardEvent('keydown', { key: 'ArrowUp', altKey: true, bubbles: true }),
    )
  })
  expect(onMove).toHaveBeenCalledWith('q3', 1)
  act(() => {
    rows[0]!.dispatchEvent(
      new window.KeyboardEvent('keydown', { key: 'ArrowDown', altKey: true, bubbles: true }),
    )
  })
  expect(onMove).toHaveBeenLastCalledWith('q1', 1)
  // plain arrows stay with the list
  act(() => {
    rows[0]!.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }))
  })
  expect(onMove).toHaveBeenCalledTimes(2)
})

it('localizes the strip labels', () => {
  act(() =>
    root.render(
      createElement(AiQueueStrip, {
        items: [{ id: 'q1', text: 'x' }],
        labels: AI_QUEUE_LABELS.zh,
        onUpdate: () => {},
        onRemove: () => {},
        onClear: () => {},
      }),
    ),
  )
  expect(host.querySelector('.ai-queue-toggle')!.textContent).toContain('1 条排队中')
})

/**
 * A cascade guard, asserted on the text because jsdom cannot see it: this
 * stylesheet used to write `.ai-queue-head`, `.ai-queue-hint` and
 * `.ai-queue-text` unqualified, and the Docs, HTML, Markdown and Slides
 * `styles.css` files already style those three names for the queued *edits*
 * panel. The cascade then ran both ways — our rules restyled their panel, and
 * their sheet (loaded last by every renderer entry) overrode ours, so the same
 * strip rendered differently per app. No render test can observe that, since
 * none of them load an app stylesheet.
 */
const STRIP_CSS = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/ai-queue-strip.css'),
  'utf8',
)

const selectorsIn = (css: string) =>
  css
    .replace(/\/\*[\s\S]*?\*\//g, '') // a comment above a rule is not part of it
    .split('{')
    .slice(0, -1)
    .flatMap((block) => block.split(',').map((s) => s.trim().split('\n').pop()!.trim()))

it('scopes every rule to the strip, so it cannot reach another component', () => {
  const selectors = selectorsIn(STRIP_CSS)
  // vacuous-pass guard: the parse has to see the rules that are there
  expect(selectors.length).toBeGreaterThanOrEqual(22)
  const escaped = selectors.filter((s) => !s.startsWith('.ai-queue-strip'))
  expect(escaped).toEqual([])
})

it('does not redefine a class the editors already style', () => {
  // the three names that collided; scoped, they can only match inside the strip
  for (const contested of ['.ai-queue-head', '.ai-queue-hint', '.ai-queue-text']) {
    const bare = selectorsIn(STRIP_CSS).filter((s) => s === contested)
    expect(bare, `${contested} must not be selectable outside .ai-queue-strip`).toEqual([])
  }
})
