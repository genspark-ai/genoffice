import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { existsSync, mkdtempSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { ProjectStore } from '../src/store.js'
import { PROJECT_RUN_CHANNELS, registerProjectRunIpc, type ProjectRunIpcMain } from '../src/ipc.js'

/**
 * The runs half of the project IPC against a real store: the three channels
 * registered through a capturing fake ipcMain, then driven exactly the way
 * ipcRenderer.invoke drives them. This is the contract the shell preload's
 * runs methods (and through them the ask panel's RunRecorder) rely on.
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
      if (!handler) return Promise.reject(new Error(`no handler registered for ${channel}`))
      try {
        return Promise.resolve(handler({}, args))
      } catch (error) {
        return Promise.reject(error instanceof Error ? error : new Error(String(error)))
      }
    },
  }
}

describe('registerProjectRunIpc', () => {
  let tmpDir: string
  let store: ProjectStore
  let ipc: ReturnType<typeof fakeIpc>

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-run-ipc-'))
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
    ipc = fakeIpc()
    registerProjectRunIpc(ipc, () => store)
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('registers exactly the three runs channels', () => {
    expect(PROJECT_RUN_CHANNELS).toEqual({
      appendRunEvent: 'project:appendRunEvent',
      listRuns: 'project:listRuns',
      deleteRun: 'project:deleteRun',
    })
  })

  it('round-trips a start/finish pair as one folded done run', async () => {
    await ipc.invoke(PROJECT_RUN_CHANNELS.appendRunEvent, {
      projectId: 'default',
      runId: 'run-1',
      kind: 'start',
      instruction: 'summarize the two open documents',
      chatId: 'kb-ask',
    })
    await ipc.invoke(PROJECT_RUN_CHANNELS.appendRunEvent, {
      projectId: 'default',
      runId: 'run-1',
      kind: 'finish',
      outcome: 'done',
    })
    const runs = (await ipc.invoke(PROJECT_RUN_CHANNELS.listRuns, {
      projectId: 'default',
    })) as Array<Record<string, unknown>>
    expect(runs).toHaveLength(1)
    expect(runs[0]).toMatchObject({
      runId: 'run-1',
      chatId: 'kb-ask',
      instruction: 'summarize the two open documents',
      status: 'done',
    })
  })

  it('a start without a terminal reads back interrupted (the closed-window case)', async () => {
    await ipc.invoke(PROJECT_RUN_CHANNELS.appendRunEvent, {
      projectId: 'default',
      runId: 'run-7',
      kind: 'start',
      instruction: 'build the deck',
    })
    const runs = (await ipc.invoke(PROJECT_RUN_CHANNELS.listRuns, {
      projectId: 'default',
    })) as Array<Record<string, unknown>>
    expect(runs[0]).toMatchObject({ runId: 'run-7', status: 'interrupted' })
  })

  it('caps the stored instruction the way the store does', async () => {
    await ipc.invoke(PROJECT_RUN_CHANNELS.appendRunEvent, {
      projectId: 'default',
      runId: 'run-9',
      kind: 'start',
      instruction: 'x'.repeat(600),
    })
    const runs = (await ipc.invoke(PROJECT_RUN_CHANNELS.listRuns, {
      projectId: 'default',
    })) as Array<{ instruction: string }>
    expect(runs[0]!.instruction).toHaveLength(200)
  })

  it('deleteRun removes one run without touching the others', async () => {
    for (const runId of ['a', 'b']) {
      await ipc.invoke(PROJECT_RUN_CHANNELS.appendRunEvent, {
        projectId: 'default',
        runId,
        kind: 'start',
        instruction: `run ${runId}`,
      })
    }
    await ipc.invoke(PROJECT_RUN_CHANNELS.deleteRun, { projectId: 'default', runId: 'a' })
    const runs = (await ipc.invoke(PROJECT_RUN_CHANNELS.listRuns, {
      projectId: 'default',
    })) as Array<{ runId: string }>
    expect(runs.map((r) => r.runId)).toEqual(['b'])
    expect(existsSync(join(storeDir('b')))).toBe(true)
  })

  it('rejects a malformed payload with a named reason, writing nothing', async () => {
    await expect(
      ipc.invoke(PROJECT_RUN_CHANNELS.appendRunEvent, {
        projectId: 'default',
        runId: 'bad-kind',
        kind: 'paused',
      }),
    ).rejects.toThrow(/Invalid run event kind/)
    await expect(
      ipc.invoke(PROJECT_RUN_CHANNELS.appendRunEvent, { projectId: 'default' }),
    ).rejects.toThrow(/Invalid runId/)
    await expect(ipc.invoke(PROJECT_RUN_CHANNELS.appendRunEvent, 'not an object')).rejects.toThrow(
      /expected an object/,
    )
    await expect(
      ipc.invoke(PROJECT_RUN_CHANNELS.appendRunEvent, {
        projectId: 'default',
        runId: 'bad-outcome',
        kind: 'finish',
        outcome: 'paused',
      }),
    ).rejects.toThrow(/Invalid run outcome/)
    await expect(ipc.invoke(PROJECT_RUN_CHANNELS.listRuns, {})).rejects.toThrow(/Invalid projectId/)
    const runs = (await ipc.invoke(PROJECT_RUN_CHANNELS.listRuns, {
      projectId: 'default',
    })) as unknown[]
    expect(runs).toHaveLength(0)
  })

  /** the store keeps its files under <tmp>/projects/<projectId>/runs/<runId>.jsonl */
  function storeDir(runId: string): string {
    return join(tmpDir, 'projects', 'default', 'runs', `${runId}.jsonl`)
  }
})
