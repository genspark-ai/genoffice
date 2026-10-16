import { describe, expect, it } from 'vitest'
import {
  HOOK_ACTIVITY_MAX,
  createHookRunner,
  installSaveGuard,
  runHooks,
  saveTargetOf,
  setLiveHookRunner,
  shouldGuardSave,
} from '../src/main/feature-hooks'
import type { HookRunner } from '../src/main/feature-hooks'
import type { FeatureRecord } from '@genoffice/agent-core'

/**
 * The hook host: which hooks run for an event, what is recorded, and what a hook
 * cannot do. The runner is injected, so this is about the host, not the sandbox.
 */

function hook(id: string, on: 'document.opened' | 'document.closed'): FeatureRecord {
  return {
    id,
    name: id,
    description: 'x',
    kind: 'hook',
    hook: { on, script: { source: '$.log("" + input.event.path)' } },
    createdAt: 'x',
    updatedAt: 'x',
  }
}

describe('hook host', () => {
  it('runs only the hooks that observe the event, in registry order', () => {
    const ran: string[] = []
    const activity = runHooks(
      [hook('a', 'document.opened'), hook('b', 'document.closed'), hook('c', 'document.opened')],
      'document.opened',
      '/tmp/report.docx',
      (id) => {
        ran.push(id)
        return { log: ['saw ' + id] }
      },
      () => '2026-01-01T00:00:00.000Z',
    )
    expect(ran).toEqual(['a', 'c'])
    expect(activity.map((entry) => entry.hookId)).toEqual(['a', 'c'])
    expect(activity[0]).toMatchObject({
      event: 'document.opened',
      path: '/tmp/report.docx',
      log: ['saw a'],
      at: '2026-01-01T00:00:00.000Z',
    })
  })

  it('records a hook that throws, and still runs the others', () => {
    const activity = runHooks(
      [hook('a', 'document.opened'), hook('b', 'document.opened')],
      'document.opened',
      '/tmp/report.docx',
      (id) => {
        if (id === 'a') throw new Error('the hook blew up')
        return { log: ['ok'] }
      },
    )
    expect(activity[0]!.error).toContain('the hook blew up')
    expect(activity[0]!.log).toEqual([])
    expect(activity[1]!.error).toBeUndefined()
    expect(activity[1]!.log).toEqual(['ok'])
  })

  it('does nothing when no feature observes the event', () => {
    const activity = runHooks(
      [hook('a', 'document.closed')],
      'document.opened',
      '/tmp/x.docx',
      () => ({
        log: [],
      }),
    )
    expect(activity).toEqual([])
  })

  it('keeps a bounded record, newest last', () => {
    let count = 0
    const runner = createHookRunner({
      list: () => [hook('a', 'document.opened')],
      run: () => ({ log: ['line ' + ++count] }),
      now: () => '2026-01-01T00:00:00.000Z',
    })
    for (let i = 0; i < HOOK_ACTIVITY_MAX + 10; i++) {
      runner.emit('document.opened', '/tmp/x' + i + '.docx')
    }
    const activity = runner.activity()
    expect(activity).toHaveLength(HOOK_ACTIVITY_MAX)
    expect(activity[activity.length - 1]!.path).toBe('/tmp/x' + (HOOK_ACTIVITY_MAX + 9) + '.docx')
    runner.dispose()
    expect(runner.activity()).toEqual([])
  })
})
/** An ipcMain stand-in: registering a handler records it, as electron would. */
function fakeIpc(): {
  handlers: Map<string, (...args: unknown[]) => unknown>
  handle(channel: string, listener: (...args: unknown[]) => unknown): void
} {
  const handlers = new Map<string, (...args: unknown[]) => unknown>()
  return {
    handlers,
    handle(channel, listener) {
      handlers.set(channel, listener)
    },
  }
}

function runnerReturning(vetoes: string[] | null): HookRunner {
  return {
    emit: () => {
      if (vetoes === null) throw new Error('the hook blew up')
      return vetoes.length
        ? [
            {
              at: '2026-01-01T00:00:00.000Z',
              hookId: 'no-untitled-reports',
              event: 'document.will-save',
              path: '/tmp/report.docx',
              log: [],
              vetoes,
            },
          ]
        : []
    },
    activity: () => [],
    dispose: () => undefined,
  }
}

describe('the save guard', () => {
  it('watches save-ish channels and leaves everything else alone', () => {
    for (const channel of ['docs:save', 'markdown:save-image', 'docs:save-new', 'pdf:save']) {
      expect(shouldGuardSave(channel)).toBe(true)
    }
    for (const channel of [
      'app:get-auto-save-default',
      'features:list',
      'features:create',
      'docs:pick-export-images-target',
    ]) {
      expect(shouldGuardSave(channel)).toBe(false)
    }
  })

  it('reads the document a save is about, and never invents one', () => {
    expect(saveTargetOf(['/tmp/a.docx'])).toBe('/tmp/a.docx')
    expect(saveTargetOf([{ path: '/tmp/b.docx' }])).toBe('/tmp/b.docx')
    expect(saveTargetOf([{ filePath: '/tmp/c.docx' }])).toBe('/tmp/c.docx')
    expect(saveTargetOf([{ nothing: true }])).toBe('')
  })

  it('calls straight through while no hook is registered', async () => {
    setLiveHookRunner(null)
    const ipc = fakeIpc()
    installSaveGuard(ipc)
    let saved = 0
    ipc.handle('docs:save', () => {
      saved++
      return 'saved'
    })
    expect(await ipc.handlers.get('docs:save')!({}, { path: '/tmp/a.docx' })).toBe('saved')
    expect(saved).toBe(1)
  })

  it('lets a hook stop a save, with its reason, and does not reach the writer', async () => {
    setLiveHookRunner(runnerReturning(['the report has no title']))
    const ipc = fakeIpc()
    installSaveGuard(ipc)
    let saved = 0
    ipc.handle('docs:save', () => {
      saved++
      return 'saved'
    })
    await expect(ipc.handlers.get('docs:save')!({}, { path: '/tmp/a.docx' })).rejects.toThrow(
      'the report has no title',
    )
    expect(saved).toBe(0)
    setLiveHookRunner(null)
  })

  it('never lets a broken hook break a save', async () => {
    setLiveHookRunner(runnerReturning(null))
    const ipc = fakeIpc()
    installSaveGuard(ipc)
    ipc.handle('pdf:save', () => 'saved')
    expect(await ipc.handlers.get('pdf:save')!({})).toBe('saved')
    setLiveHookRunner(null)
  })

  it('does not inspect a channel that is not a save', async () => {
    setLiveHookRunner(runnerReturning(['never mind']))
    const ipc = fakeIpc()
    installSaveGuard(ipc)
    ipc.handle('features:list', () => ['ok'])
    expect(await ipc.handlers.get('features:list')!({})).toEqual(['ok'])
    setLiveHookRunner(null)
  })
})
