import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  AgentLoop,
  activeAppearance,
  createFeaturesSkill,
  slotsAt,
  type AgentStreamCallbacks,
  type AgentStreamRequest,
  type AgentTransport,
} from '@genoffice/agent-core'
import {
  createFeature,
  getFeature,
  listFeatures,
  removeFeature,
  updateFeature,
} from '../src/main/feature-store'

/**
 * The interface end to end, without electron: a scripted model asks for a
 * feature, the loop hands the call to the skill, the skill writes it through the
 * same registry functions the IPC boundary uses, and a fresh read of the
 * directory (a "restart") still sees it.
 */

type Step = (callbacks: AgentStreamCallbacks) => void

function scripted(steps: Step[]): { transport: AgentTransport; requests: AgentStreamRequest[] } {
  const requests: AgentStreamRequest[] = []
  let turn = 0
  const transport: AgentTransport = {
    stream(request, callbacks) {
      requests.push(request)
      const step = steps[turn++]
      if (step) setTimeout(() => step(callbacks), 0)
      return { cancel: () => setTimeout(() => callbacks.onDone(), 0) }
    },
  }
  return { transport, requests }
}

function registry(dir: string) {
  return {
    snapshot: () => listFeatures(dir),
    list: () => listFeatures(dir),
    get: (id: string) => getFeature(dir, id),
    create: (draft: { name: string; description: string; kind?: string }) =>
      createFeature(dir, draft),
    update: (id: string, patch: { name?: string; description?: string; kind?: string }) =>
      updateFeature(dir, id, patch),
    remove: (id: string) => removeFeature(dir, id),
    readDocument: async () => ({ name: 'report.docx', text: 'text', chars: 4 }),
  }
}

describe('mod assistant flow', () => {
  it('records what the model asks for, and the record survives a fresh read', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-flow-'))
    const deps = registry(dir)
    const { transport, requests } = scripted([
      (callbacks) => {
        callbacks.onToolCall({
          id: 'call-1',
          name: 'create_feature',
          input: { name: 'Weekly report formatting', description: 'apply the house style' },
        })
        callbacks.onDone()
      },
      (callbacks) => {
        callbacks.onDelta('Recorded weekly-report-formatting. It does not run yet.')
        callbacks.onDone()
      },
    ])

    const finished = new Promise<string>((resolve) => {
      const loop = new AgentLoop({
        transport,
        skill: createFeaturesSkill(deps),
        events: { onDone: (result) => resolve(result.text) },
      })
      loop.run('add a weekly report formatting feature')
    })

    expect(await finished).toContain('does not run yet')
    // the model was handed the interface, not a hard-coded capability
    expect(requests[0]!.tools.map((tool) => tool.name)).toContain('create_feature')
    expect(listFeatures(dir).map((feature) => feature.name)).toEqual(['Weekly report formatting'])
    // a second registry object over the same directory is what a restart looks like
    expect(getFeature(dir, 'weekly-report-formatting')?.description).toBe('apply the house style')
  })

  it('reports a tool that cannot find its id instead of inventing a record', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-flow-'))
    const deps = registry(dir)
    const { transport } = scripted([
      (callbacks) => {
        callbacks.onToolCall({
          id: 'call-1',
          name: 'delete_feature',
          input: { id: 'does-not-exist' },
        })
        callbacks.onDone()
      },
      (callbacks) => {
        callbacks.onDelta('There is no such feature.')
        callbacks.onDone()
      },
    ])
    const outputs: string[] = []
    const finished = new Promise<void>((resolve) => {
      const loop = new AgentLoop({
        transport,
        skill: createFeaturesSkill(deps),
        events: {
          onToolExecuted: (event) => outputs.push(event.execution.output),
          onDone: () => resolve(),
        },
      })
      loop.run('delete that feature')
    })
    await finished
    expect(outputs).toEqual(['no feature with id does-not-exist'])
    expect(listFeatures(dir)).toEqual([])
  })

  it('registers a slot the host can draw, in the place the model named', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-flow-'))
    const deps = registry(dir)
    const { transport } = scripted([
      (callbacks) => {
        callbacks.onToolCall({
          id: 'call-1',
          name: 'create_feature',
          input: {
            name: 'File switcher',
            description: 'jump to a recent file',
            slot: {
              at: 'home.sidebar',
              element: {
                type: 'button',
                label: 'Files',
                action: { kind: 'open', path: '/tmp/a.docx' },
              },
            },
          },
        })
        callbacks.onDone()
      },
      (callbacks) => {
        callbacks.onDelta('Recorded file-switcher, drawn in the sidebar.')
        callbacks.onDone()
      },
    ])
    const finished = new Promise<void>((resolve) => {
      const loop = new AgentLoop({
        transport,
        skill: createFeaturesSkill(deps),
        events: { onDone: () => resolve() },
      })
      loop.run('add a file switcher to the sidebar')
    })
    await finished
    // a fresh read is what Home's registry renders
    expect(slotsAt(listFeatures(dir), 'home.sidebar').map((f) => f.id)).toEqual(['file-switcher'])
    expect(slotsAt(listFeatures(dir), 'home.card')).toEqual([])
  })

  it('registers an appearance, which the shell then applies to its own chrome', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-flow-'))
    const deps = registry(dir)
    const { transport } = scripted([
      (callbacks) => {
        callbacks.onToolCall({
          id: 'call-1',
          name: 'create_feature',
          input: {
            name: 'Night shift',
            description: 'softer chrome after dark',
            appearance: { background: 'surface-subtle', accent: 'violet' },
          },
        })
        callbacks.onDone()
      },
      (callbacks) => {
        callbacks.onDelta('The shell now uses a violet accent.')
        callbacks.onDone()
      },
    ])
    const finished = new Promise<void>((resolve) => {
      const loop = new AgentLoop({
        transport,
        skill: createFeaturesSkill(deps),
        events: { onDone: () => resolve() },
      })
      loop.run('give the app a violet accent')
    })
    await finished
    // what the renderer's applyAppearance reads
    expect(activeAppearance(listFeatures(dir))).toEqual({
      background: 'surface-subtle',
      accent: 'violet',
    })
  })

  it('registers a host-filled file list, which is how a file sidebar is built', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-flow-'))
    const deps = registry(dir)
    const { transport } = scripted([
      (callbacks) => {
        callbacks.onToolCall({
          id: 'call-1',
          name: 'create_feature',
          input: {
            name: 'Files',
            description: 'jump to a recent file',
            slot: {
              at: 'home.panel',
              element: {
                type: 'file-list',
                source: 'recent-files',
                limit: 6,
                action: { kind: 'open', path: '$path' },
              },
            },
          },
        })
        callbacks.onDone()
      },
      (callbacks) => {
        callbacks.onDelta('Recorded Files, drawn from the recent-files source.')
        callbacks.onDone()
      },
    ])
    const finished = new Promise<void>((resolve) => {
      const loop = new AgentLoop({
        transport,
        skill: createFeaturesSkill(deps),
        events: { onDone: () => resolve() },
      })
      loop.run('give me a file sidebar')
    })
    await finished
    const drawn = slotsAt(listFeatures(dir), 'home.panel')
    expect(drawn.map((f) => f.id)).toEqual(['files'])
    expect(drawn[0]!.slot?.element).toMatchObject({
      type: 'file-list',
      source: 'recent-files',
      limit: 6,
    })
  })
})
