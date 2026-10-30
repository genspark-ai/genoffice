import { describe, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  ProjectStore,
  registerProjectRunIpc,
  type ProjectApi,
  type ProjectRunIpcMain,
} from '@genoffice/project-store'
import type { RunSummary } from '@genoffice/project-store'
import {
  interruptedOf,
  resolveRunsProject,
  RunRecorder,
  runsBridgeOf,
  type RunsBridge,
} from '../src/renderer/src/ask-runs'

/**
 * The ask panel's runs transport over the first-class runs IPC. The
 * recorder-facing tests drive the real handlers: registerProjectRunIpc over a
 * real ProjectStore in a temp dir, with a fake ipcMain capturing the channels
 * the way ipcRenderer.invoke would reach them — so "the panel reads
 * interrupted tasks" is the actual contract the shell preload forwards.
 */

type Handler = (event: unknown, args: unknown) => unknown

function fakeIpc(): ProjectRunIpcMain & {
  invoke(channel: string, args?: unknown): Promise<unknown>
} {
  const handlers = new Map<string, Handler>()
  return {
    handle: (channel, listener) => handlers.set(channel, listener),
    invoke: (channel, args) => {
      const handler = handlers.get(channel)
      if (!handler) return Promise.reject(new Error(`no handler for ${channel}`))
      try {
        return Promise.resolve(handler({}, args))
      } catch (error) {
        return Promise.reject(error instanceof Error ? error : new Error(String(error)))
      }
    },
  }
}

/** the renderer bridges as the panel would see them: home runs methods + the chat resolver */
function bridgesOver(ipc: ReturnType<typeof fakeIpc>): {
  home: Parameters<typeof runsBridgeOf>[0]
  project: Pick<ProjectApi, 'resolveChat'>
} {
  return {
    home: {
      appendRunEvent: (args) => ipc.invoke('project:appendRunEvent', args),
      listRuns: (args) => ipc.invoke('project:listRuns', args) as Promise<RunSummary[]>,
      deleteRun: (args) => ipc.invoke('project:deleteRun', args),
    },
    project: {
      resolveChat: async () => ({ projectId: 'default', chatId: 'kb-runs' }),
    },
  }
}

/** recorder + store over a fresh temp dir, wired through the real handlers */
function recorderWithStore(): {
  recorder: RunRecorder
  bridge: RunsBridge
  store: ProjectStore
  tmpDir: string
} {
  const tmpDir = mkdtempSync(join(tmpdir(), 'genoffice-ask-runs-'))
  const store = new ProjectStore(tmpDir)
  store.ensureDefaultProject()
  const ipc = fakeIpc()
  registerProjectRunIpc(ipc, () => store)
  const { home, project } = bridgesOver(ipc)
  const bridge = runsBridgeOf(home, project)
  if (!bridge) throw new Error('bridge should be available')
  return { recorder: new RunRecorder(bridge), bridge, store, tmpDir }
}

describe('runsBridgeOf', () => {
  const home = {
    appendRunEvent: async () => {},
    listRuns: async () => [] as RunSummary[],
    deleteRun: async () => {},
  }
  const project = { resolveChat: async () => ({ projectId: 'default', chatId: 'kb-ask' }) }

  it('needs both halves of the bridge', () => {
    expect(runsBridgeOf(home, project)).toBeDefined()
    expect(runsBridgeOf(undefined, project)).toBeUndefined()
    expect(runsBridgeOf(home, undefined)).toBeUndefined()
    expect(runsBridgeOf(null, null)).toBeUndefined()
  })

  it('a stale preload without the runs methods degrades to no bridge', () => {
    expect(
      runsBridgeOf({ listRuns: home.listRuns, deleteRun: home.deleteRun }, project),
    ).toBeUndefined()
    expect(runsBridgeOf({}, project)).toBeUndefined()
  })

  it('resolves the default project through the file-less chat lookup', async () => {
    const resolveChat = vi.fn(async () => ({ projectId: 'default', chatId: 'kb-ask' }))
    await expect(resolveRunsProject({ resolveChat })).resolves.toBe('default')
    expect(resolveChat).toHaveBeenCalledWith({ filePath: null })
  })
})

describe('RunRecorder over the runs IPC (bridge roundtrip)', () => {
  it('reads back the runs earlier sessions left unfinished', async () => {
    const { recorder, store, bridge, tmpDir } = recorderWithStore()
    try {
      // an earlier session: started, window closed before any terminal
      store.appendRunEvent('default', 'run-7', {
        kind: 'start',
        instruction: 'build the deck',
        chatId: 'kb-ask',
      })
      const interrupted = await recorder.init()
      expect(interrupted).toHaveLength(1)
      expect(interrupted[0]).toMatchObject({
        runId: 'run-7',
        instruction: 'build the deck',
        chatId: 'kb-ask',
        status: 'interrupted',
      })
      expect(interruptedOf(interrupted)).toHaveLength(1)
    } finally {
      rmSync(tmpDir, { recursive: true, force: true })
    }
  })

  it('writes start/finish through the handlers and settles the fold', async () => {
    const { recorder, bridge, tmpDir } = recorderWithStore()
    try {
      await recorder.init()
      recorder.start('run-1', 'summarize the two open documents', 'kb-ask')
      recorder.finish('run-1', 'done')
      const runs = await bridge.listRuns({ projectId: 'default' })
      expect(runs).toHaveLength(1)
      expect(runs[0]).toMatchObject({
        runId: 'run-1',
        status: 'done',
        instruction: 'summarize the two open documents',
      })
    } finally {
      rmSync(tmpDir, { recursive: true, force: true })
    }
  })

  it('dismiss settles an interrupted run as discarded', async () => {
    const { recorder, bridge, tmpDir } = recorderWithStore()
    try {
      await recorder.init()
      recorder.start('run-9', 'build the deck', 'kb-ask')
      recorder.discard('run-9')
      const runs = await bridge.listRuns({ projectId: 'default' })
      expect(runs[0]).toMatchObject({ runId: 'run-9', status: 'discarded' })
      expect(interruptedOf(runs)).toHaveLength(0)
    } finally {
      rmSync(tmpDir, { recursive: true, force: true })
    }
  })

  it('clear deletes every stored run record', async () => {
    const { recorder, bridge, tmpDir } = recorderWithStore()
    try {
      await recorder.init()
      for (const runId of ['a', 'b', 'c']) {
        recorder.start(runId, `run ${runId}`, 'kb-ask')
        recorder.finish(runId, 'done')
      }
      await recorder.clear()
      expect(await bridge.listRuns({ projectId: 'default' })).toEqual([])
    } finally {
      rmSync(tmpDir, { recursive: true, force: true })
    }
  })

  it('queues events written before init and flushes them in order', async () => {
    const { recorder, bridge, tmpDir } = recorderWithStore()
    try {
      recorder.start('run-2', 'started before the project resolved', 'kb-ask')
      recorder.finish('run-2', 'cancelled')
      await recorder.init()
      const runs = await bridge.listRuns({ projectId: 'default' })
      expect(runs).toHaveLength(1)
      expect(runs[0]).toMatchObject({ runId: 'run-2', status: 'cancelled' })
    } finally {
      rmSync(tmpDir, { recursive: true, force: true })
    }
  })

  it('swallows append failures: run records are telemetry, never fatal', async () => {
    const failing: RunsBridge = {
      appendRunEvent: async () => {
        throw new Error('disk full')
      },
      listRuns: async () => [],
      deleteRun: async () => {},
      resolveProject: async () => 'default',
    }
    const recorder = new RunRecorder(failing)
    await recorder.init()
    expect(() => {
      recorder.start('run-3', 'still runs', 'kb-ask')
      recorder.error('run-3', 'provider is down')
    }).not.toThrow()
    // give the swallowed rejections a tick to surface as unhandled if they leak
    await new Promise((resolve) => setTimeout(resolve, 0))
  })
})
