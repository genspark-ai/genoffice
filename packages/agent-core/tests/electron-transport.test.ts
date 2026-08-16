import { describe, expect, it, vi } from 'vitest'
import {
  createIpcTransport,
  IPC_STREAM_SILENCE_TIMEOUT_MS,
  type IpcStreamChunk,
  type IpcStreamStart,
} from '../src'

interface FakeSettings {
  provider: string
}

function setup(
  startImpl?: (request: IpcStreamStart<FakeSettings>) => void | Promise<unknown>,
  creditsErrorText?: () => string,
  networkErrorText?: () => string,
) {
  let listener: ((chunk: IpcStreamChunk) => void) | undefined
  const unsubscribe = vi.fn(() => {
    listener = undefined
  })
  const started: IpcStreamStart<FakeSettings>[] = []
  const cancelled: string[] = []
  const transport = createIpcTransport<FakeSettings>({
    onStream: (l) => {
      listener = l
      return unsubscribe
    },
    start: (request) => {
      started.push(request)
      return startImpl?.(request)
    },
    cancel: (requestId) => cancelled.push(requestId),
    getSettings: () => ({ provider: 'genspark' }),
    unknownErrorText: () => 'unknown error',
    timeoutErrorText: () => 'timed out',
    ...(creditsErrorText ? { creditsErrorText } : {}),
    ...(networkErrorText ? { networkErrorText } : {}),
  })
  const cb = {
    onDelta: vi.fn(),
    onToolCall: vi.fn(),
    onStopReason: vi.fn(),
    onDone: vi.fn(),
    onError: vi.fn(),
  }
  const handle = transport.stream({ system: 'sys', messages: [], tools: [] }, cb)
  const emit = (chunk: Omit<IpcStreamChunk, 'requestId'> & { requestId?: string }) =>
    listener?.({ requestId: started[0]!.requestId, ...chunk })
  return { started, cancelled, cb, handle, emit, unsubscribe }
}

describe('createIpcTransport', () => {
  it('starts one request with settings and forwards deltas and tool calls', () => {
    const { started, cb, emit } = setup()
    expect(started).toHaveLength(1)
    expect(started[0]!.settings).toEqual({ provider: 'genspark' })
    expect(started[0]!.system).toBe('sys')

    emit({ type: 'delta', text: 'hi' })
    emit({ type: 'delta' })
    emit({ type: 'tool-call', toolCall: { id: 'c1', name: 'read', input: {} } })
    expect(cb.onDelta).toHaveBeenNthCalledWith(1, 'hi')
    expect(cb.onDelta).toHaveBeenNthCalledWith(2, '')
    expect(cb.onToolCall).toHaveBeenCalledWith({ id: 'c1', name: 'read', input: {} })
  })

  it('ignores chunks for other requestIds', () => {
    const { cb, emit } = setup()
    emit({ requestId: 'someone-else', type: 'delta', text: 'nope' })
    expect(cb.onDelta).not.toHaveBeenCalled()
  })

  it('unsubscribes on done', () => {
    const { cb, emit, unsubscribe } = setup()
    emit({ type: 'done' })
    expect(cb.onDone).toHaveBeenCalledTimes(1)
    expect(cb.onStopReason).not.toHaveBeenCalled()
    expect(unsubscribe).toHaveBeenCalledTimes(1)
  })

  it('forwards a stopReason carried on the done chunk before onDone', () => {
    const { cb, emit } = setup()
    emit({ type: 'done', stopReason: 'max_tokens' })
    expect(cb.onStopReason).toHaveBeenCalledWith('max_tokens')
    expect(cb.onDone).toHaveBeenCalledTimes(1)
  })

  it('maps error chunks to onError with the localized fallback', () => {
    const { cb, emit, unsubscribe } = setup()
    emit({ type: 'error' })
    expect(cb.onError).toHaveBeenCalledWith('unknown error')
    expect(unsubscribe).toHaveBeenCalledTimes(1)
  })

  it('cancel forwards the requestId to the bridge', () => {
    const { started, cancelled, handle } = setup()
    handle.cancel()
    expect(cancelled).toEqual([started[0]!.requestId])
  })

  it('maps a timeout error code to the localized timeout message', () => {
    const { cb, emit } = setup()
    emit({ type: 'error', error: 'AI request timed out: no data received', errorCode: 'timeout' })
    expect(cb.onError).toHaveBeenCalledWith('timed out')
  })

  it('maps a credits error code to the localized credits message', () => {
    const { cb, emit } = setup(undefined, () => 'credits used up')
    emit({
      type: 'error',
      error: 'Your Genspark credits have been exhausted.',
      errorCode: 'credits',
    })
    expect(cb.onError).toHaveBeenCalledWith('credits used up')
  })

  it('maps a network error code to the localized network message', () => {
    const { cb, emit } = setup(undefined, undefined, () => 'network problem')
    emit({
      type: 'error',
      error: 'Claude fetch failed: fetch failed cause=ECONNRESET',
      errorCode: 'network',
    })
    expect(cb.onError).toHaveBeenCalledWith('network problem')
  })

  it('a network error code without networkErrorText falls back to the carried text', () => {
    const { cb, emit } = setup()
    emit({
      type: 'error',
      error: 'Claude fetch failed: fetch failed cause=ECONNRESET',
      errorCode: 'network',
    })
    expect(cb.onError).toHaveBeenCalledWith('Claude fetch failed: fetch failed cause=ECONNRESET')
  })

  it('a credits error code without creditsErrorText falls back to the carried text', () => {
    const { cb, emit } = setup()
    emit({
      type: 'error',
      error: 'Your Genspark credits have been exhausted.',
      errorCode: 'credits',
    })
    expect(cb.onError).toHaveBeenCalledWith('Your Genspark credits have been exhausted.')
  })

  it('fails the run after prolonged silence; pings re-arm the watchdog', () => {
    vi.useFakeTimers()
    try {
      const { cb, emit, started, cancelled } = setup()
      emit({ type: 'delta', text: 'x' })
      vi.advanceTimersByTime(IPC_STREAM_SILENCE_TIMEOUT_MS - 1)
      emit({ type: 'ping' })
      vi.advanceTimersByTime(IPC_STREAM_SILENCE_TIMEOUT_MS - 1)
      expect(cb.onError).not.toHaveBeenCalled()
      vi.advanceTimersByTime(IPC_STREAM_SILENCE_TIMEOUT_MS)
      expect(cb.onError).toHaveBeenCalledWith('timed out')
      expect(cancelled).toEqual([started[0]!.requestId])
    } finally {
      vi.useRealTimers()
    }
  })

  it('done disarms the silence watchdog', () => {
    vi.useFakeTimers()
    try {
      const { cb, emit } = setup()
      emit({ type: 'done' })
      vi.advanceTimersByTime(IPC_STREAM_SILENCE_TIMEOUT_MS * 2)
      expect(cb.onError).not.toHaveBeenCalled()
      expect(cb.onDone).toHaveBeenCalledTimes(1)
    } finally {
      vi.useRealTimers()
    }
  })

  it('a rejected start fails the run instead of leaving it pending', async () => {
    const { cb } = setup(() => Promise.reject(new Error('no handler registered')))
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(cb.onError).toHaveBeenCalledWith('no handler registered')
    expect(cb.onDone).not.toHaveBeenCalled()
  })
})
