import { EventEmitter } from 'node:events'
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'

const execFileMock = vi.hoisted(() => vi.fn())
vi.mock('node:child_process', () => ({ execFile: execFileMock }))

import { gskLogin, gskLoginStart, type GskLoginProgress } from '../src/gsk'

class FakeStream extends EventEmitter {
  setEncoding(): this {
    return this
  }
}

interface FakeChild {
  stdout: FakeStream
  stderr: FakeStream
  kill: ReturnType<typeof vi.fn>
  /** invokes the execFile completion callback (fires on child exit) */
  exit: (error: (Error & { killed?: boolean }) | null) => void
}

let spawned: FakeChild[]

beforeEach(() => {
  spawned = []
  process.env.GSK_CLI_PATH = '/fake/gsk/index.js'
  execFileMock.mockImplementation((...args: unknown[]) => {
    const callback = args[args.length - 1] as FakeChild['exit']
    const child: FakeChild = {
      stdout: new FakeStream(),
      stderr: new FakeStream(),
      kill: vi.fn(),
      exit: callback,
    }
    spawned.push(child)
    return child
  })
})

afterEach(() => {
  // drain module state: make sure the last spawned CLI has exited cleanly so
  // the next test starts without an in-flight login session
  const last = spawned[spawned.length - 1]
  if (last && !last.kill.mock.calls.length) last.exit(null)
  delete process.env.GSK_CLI_PATH
  execFileMock.mockReset()
})

describe('gskLoginStart retry', () => {
  it('silences the superseded session before its pipes drain', () => {
    const eventsA: GskLoginProgress[] = []
    const eventsB: GskLoginProgress[] = []
    expect(gskLoginStart((p) => eventsA.push(p))).toBe(true)
    const childA = spawned[0]!
    childA.stderr.emit('data', '[INFO] Login URL: https://genspark.ai/cli-auth?code=old\n')
    expect(eventsA).toEqual([{ phase: 'url', url: 'https://genspark.ai/cli-auth?code=old' }])

    // retry: A is killed, but its buffered output and exit arrive *after* B starts
    expect(gskLoginStart((p) => eventsB.push(p))).toBe(true)
    expect(childA.kill).toHaveBeenCalled()
    const childB = spawned[1]!
    childA.stderr.emit('data', '[INFO] Login URL: https://genspark.ai/cli-auth?code=stale\n')
    childA.stderr.emit('data', '[ERROR] Authorization expired. Please try again.\n')
    childA.exit(Object.assign(new Error('killed'), { killed: true }))
    expect(eventsA).toHaveLength(1) // nothing new leaked into the old callback

    childB.stderr.emit('data', '[INFO] Login URL: https://genspark.ai/cli-auth?code=new\n')
    childB.stdout.emit('data', '"message": "Login successful"\n')
    expect(eventsB).toEqual([
      { phase: 'url', url: 'https://genspark.ai/cli-auth?code=new' },
      { phase: 'success' },
    ])
  })

  it("does not let the killed child's exit clear the new session", () => {
    expect(gskLoginStart()).toBe(true)
    expect(gskLoginStart()).toBe(true)
    spawned[0]!.exit(Object.assign(new Error('killed'), { killed: true }))
    // session B must still be tracked as in flight: gskLogin() reuses it
    expect(gskLogin()).toBe(true)
    expect(execFileMock).toHaveBeenCalledTimes(2)
  })
})

describe('gskLogin', () => {
  it('reuses an in-flight login instead of killing it', () => {
    const events: GskLoginProgress[] = []
    expect(gskLoginStart((p) => events.push(p))).toBe(true)
    const child = spawned[0]!

    expect(gskLogin()).toBe(true)
    expect(execFileMock).toHaveBeenCalledTimes(1)
    expect(child.kill).not.toHaveBeenCalled()

    // the original session still reports progress to its caller
    child.stderr.emit('data', '[INFO] Login successful! API key saved.\n')
    expect(events).toEqual([{ phase: 'success' }])
  })

  it('starts a fresh login once the previous one has exited', () => {
    expect(gskLogin()).toBe(true)
    spawned[0]!.exit(null)
    expect(gskLogin()).toBe(true)
    expect(execFileMock).toHaveBeenCalledTimes(2)
  })

  it('starts a fresh login after a terminal error, even if the old CLI lingers', () => {
    const events: GskLoginProgress[] = []
    expect(gskLoginStart((p) => events.push(p))).toBe(true)
    const child = spawned[0]!
    child.stderr.emit('data', '[INFO] Login URL: https://genspark.ai/cli-auth?code=x\n')
    child.stderr.emit('data', '[ERROR] Authorization expired. Please try again.\n')
    expect(events.at(-1)).toEqual({ phase: 'error', error: 'expired' })
    expect(child.kill).toHaveBeenCalled() // lingering CLI is reaped with the session

    // the errored child has not exited yet, but the session is terminal:
    // a passive gskLogin() must spawn a new CLI instead of reusing the corpse
    expect(gskLogin()).toBe(true)
    expect(execFileMock).toHaveBeenCalledTimes(2)
  })
})
