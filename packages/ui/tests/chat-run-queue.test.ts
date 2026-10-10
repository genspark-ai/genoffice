/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { useChatRunQueue, type ChatRunQueue } from '../src/chat-run-queue'

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

interface ProbeProps {
  busy: boolean
  submit: (text: string, meta: string | undefined) => boolean
  resetKey?: string | number
}

let latest: ChatRunQueue<string | undefined>
/** mounts the hook through a probe component so busy/submit are drivable per render */
function Probe(props: ProbeProps) {
  latest = useChatRunQueue<string | undefined>({
    busy: props.busy,
    submit: props.submit,
    resetKey: props.resetKey,
  })
  return null
}

function render(props: ProbeProps) {
  act(() => root.render(createElement(Probe, props)))
}

it('queues a draft while busy and pumps it when the run settles', () => {
  const submitted: string[] = []
  render({
    busy: true,
    submit: (text) => {
      submitted.push(text)
      return true
    },
  })
  act(() => latest.enqueue('first', undefined))
  expect(latest.queued.map((m) => m.text)).toEqual(['first'])
  // the run settles: the pump starts the queued message and drops it from the strip
  render({
    busy: false,
    submit: (text) => {
      submitted.push(text)
      return true
    },
  })
  expect(submitted).toEqual(['first'])
  expect(latest.queued).toHaveLength(0)
})

it('pumps the queue in order across consecutive settles and passes the meta through', () => {
  const seen: Array<[string, string | undefined]> = []
  let busy = true
  const renderAt = () =>
    render({
      busy,
      submit: (text, meta) => {
        seen.push([text, meta])
        return true
      },
    })
  renderAt()
  act(() => {
    latest.enqueue('a', 'meta-a')
    latest.enqueue('b', undefined)
  })
  expect(latest.queued.map((m) => [m.text, m.meta])).toEqual([
    ['a', 'meta-a'],
    ['b', undefined],
  ])
  // two run cycles: each settle pumps exactly one queued message
  for (let i = 0; i < 2; i++) {
    busy = false
    renderAt()
    busy = true
    renderAt()
  }
  expect(seen).toEqual([
    ['a', 'meta-a'],
    ['b', undefined],
  ])
  expect(latest.queued).toHaveLength(0)
})

it('submits straight away when enqueue lands on an idle panel', () => {
  const submitted: string[] = []
  render({
    busy: false,
    submit: (text) => {
      submitted.push(text)
      return true
    },
  })
  act(() => latest.enqueue('now', undefined))
  expect(submitted).toEqual(['now'])
  expect(latest.queued).toHaveLength(0)
})

it('keeps the message queued when submit refuses, then retries it', () => {
  vi.useFakeTimers()
  try {
    let accept = false
    const submitted: string[] = []
    let busy = true
    const renderAt = () =>
      render({
        busy,
        submit: (text) => {
          if (!accept) return false
          submitted.push(text)
          return true
        },
      })
    renderAt()
    act(() => latest.enqueue('patient', undefined))
    // settle with submit refusing: the message stays queued for the retry net
    busy = false
    renderAt()
    expect(submitted).toEqual([])
    expect(latest.queued).toHaveLength(1)
    act(() => {
      vi.advanceTimersByTime(900)
    })
    expect(submitted).toEqual([])
    accept = true
    act(() => {
      vi.advanceTimersByTime(900)
    })
    expect(submitted).toEqual(['patient'])
    expect(latest.queued).toHaveLength(0)
  } finally {
    vi.useRealTimers()
  }
})

it('edits, removes, and clears queued messages', () => {
  render({ busy: true, submit: () => true })
  act(() => {
    latest.enqueue('one', undefined)
    latest.enqueue('two', undefined)
  })
  act(() => latest.update('q1', 'one!'))
  expect(latest.queued.map((m) => m.text)).toEqual(['one!', 'two'])
  // empty edits are ignored, not destructive
  act(() => latest.update('q1', '   '))
  expect(latest.queued[0]!.text).toBe('one!')
  act(() => latest.remove('q1'))
  expect(latest.queued.map((m) => m.text)).toEqual(['two'])
  act(() => latest.clear())
  expect(latest.queued).toHaveLength(0)
})

it('holds several queued messages while paused and runs them once resumed', () => {
  const submitted: string[] = []
  let busy = true
  const renderAt = () =>
    render({
      busy,
      submit: (text) => {
        submitted.push(text)
        return true
      },
    })
  renderAt()
  act(() => {
    latest.enqueue('first', undefined)
    latest.enqueue('second', undefined)
  })
  expect(latest.queued.map((m) => m.text)).toEqual(['first', 'second'])
  act(() => latest.setPaused(true))
  expect(latest.paused).toBe(true)
  // the run in flight settles: a paused queue does not start the next one
  busy = false
  renderAt()
  expect(submitted).toEqual([])
  expect(latest.queued).toHaveLength(2)
  // resuming opens the gate and the panel is idle, so the head starts at once
  act(() => latest.setPaused(false))
  expect(submitted).toEqual(['first'])
  expect(latest.queued.map((m) => m.text)).toEqual(['second'])
})

it('never starts a message while paused, not even after an edit or a retry tick', () => {
  vi.useFakeTimers()
  try {
    const submitted: string[] = []
    render({
      busy: false,
      submit: (text) => {
        submitted.push(text)
        return true
      },
    })
    // pause first: on an idle panel, enqueueing would start the message at once
    act(() => latest.setPaused(true))
    act(() => {
      latest.enqueue('one', undefined)
      latest.enqueue('two', undefined)
    })
    expect(latest.queued).toHaveLength(2)
    // the retry net is down while paused: no timer can slip a message out
    act(() => {
      vi.advanceTimersByTime(2000)
    })
    expect(submitted).toEqual([])
    // editing one queued message must not set the other off
    act(() => latest.update('q2', 'two!'))
    act(() => {
      vi.advanceTimersByTime(2000)
    })
    expect(submitted).toEqual([])
    expect(latest.queued.map((m) => m.text)).toEqual(['one', 'two!'])
    act(() => latest.togglePaused())
    expect(submitted).toEqual(['one'])
    expect(latest.queued.map((m) => m.text)).toEqual(['two!'])
  } finally {
    vi.useRealTimers()
  }
})

it('reorders queued messages and clamps out-of-range targets', () => {
  render({ busy: true, submit: () => true })
  act(() => {
    latest.enqueue('one', undefined)
    latest.enqueue('two', undefined)
    latest.enqueue('three', undefined)
  })
  act(() => latest.move('q3', 0))
  expect(latest.queued.map((m) => m.text)).toEqual(['three', 'one', 'two'])
  act(() => latest.move('q3', 2))
  expect(latest.queued.map((m) => m.text)).toEqual(['one', 'two', 'three'])
  // clamping keeps the message in the queue instead of dropping it
  act(() => latest.move('q1', 99))
  expect(latest.queued.map((m) => m.text)).toEqual(['two', 'three', 'one'])
  act(() => latest.move('q1', -5))
  expect(latest.queued.map((m) => m.text)).toEqual(['one', 'two', 'three'])
  // unknown ids are a no-op
  act(() => latest.move('nope', 0))
  expect(latest.queued.map((m) => m.text)).toEqual(['one', 'two', 'three'])
})

it('keeps the meta with the message when it is moved', () => {
  render({ busy: true, submit: () => true })
  act(() => {
    latest.enqueue('first', 'meta-1')
    latest.enqueue('second', 'meta-2')
  })
  act(() => latest.move('q2', 0))
  expect(latest.queued.map((m) => [m.text, m.meta])).toEqual([
    ['second', 'meta-2'],
    ['first', 'meta-1'],
  ])
})

it('wipes the queue when the reset key changes', () => {
  render({ busy: true, submit: () => true, resetKey: 'doc-1' })
  act(() => {
    latest.enqueue('stale', undefined)
  })
  render({ busy: true, submit: () => true, resetKey: 'doc-2' })
  expect(latest.queued).toHaveLength(0)
})

it('ignores blank drafts', () => {
  render({ busy: true, submit: () => true })
  act(() => latest.enqueue('   ', undefined))
  expect(latest.queued).toHaveLength(0)
})
