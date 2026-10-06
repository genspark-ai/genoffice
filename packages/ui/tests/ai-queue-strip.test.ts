/** @vitest-environment jsdom */
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
