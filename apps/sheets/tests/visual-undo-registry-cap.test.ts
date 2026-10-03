import { ICommandService, IUndoRedoService } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { pushVisualUndo } from '../src/renderer/univer-sync'
import { VISUAL_UNDO_COMMAND_ID } from '../src/renderer/undo-carry'
import type { UniverRuntime } from '../src/renderer/univer-state'

/// The registry's eviction bound (VISUAL_UNDO_REGISTRY_CAP in univer-sync.ts).
/// Pinned here rather than imported so the behavioural assertions below fail
/// on code that never prunes, instead of on a missing export.
const CAP = 100

type Step = { undo: () => void; redo: () => void }
type Handler = (
  accessor: unknown,
  params?: { token: number; direction: 'undo' | 'redo' },
) => boolean

/// A runtime whose injector serves only the two services the visual-undo
/// registry touches, capturing the command handler and every pushed token.
/// A fresh one per test: ensureVisualUndoCommand registers once per runtime.
function fakeRuntime(): {
  runtime: UniverRuntime
  tokens: number[]
  run: (token: number, direction?: 'undo' | 'redo') => boolean
} {
  const tokens: number[] = []
  let handler: Handler | null = null
  const injector = {
    get(token: unknown): unknown {
      if (token === ICommandService) {
        return {
          registerCommand(command: { id: string; handler: Handler }): void {
            if (command.id === VISUAL_UNDO_COMMAND_ID) handler = command.handler
          },
        }
      }
      if (token === IUndoRedoService) {
        return {
          pushUndoRedo(item: {
            undoMutations: Array<{ id: string; params: { token: number } }>
          }): void {
            tokens.push(item.undoMutations[0]!.params.token)
          },
        }
      }
      return undefined
    },
  }
  const runtime = {
    univer: { __getInjector: () => injector },
    univerAPI: { getActiveWorkbook: () => ({ getId: () => 'wb' }) },
  } as unknown as UniverRuntime
  return {
    runtime,
    tokens,
    run: (token, direction = 'undo') => handler?.(null, { token, direction }) ?? false,
  }
}

function step(): { value: Step; undone: () => number; redone: () => number } {
  let undone = 0
  let redone = 0
  return {
    value: {
      undo: () => {
        undone += 1
      },
      redo: () => {
        redone += 1
      },
    },
    undone: () => undone,
    redone: () => redone,
  }
}

describe('visual undo registry is bounded', () => {
  it('drops the oldest step once the registry exceeds its bound', () => {
    const { runtime, tokens, run } = fakeRuntime()
    const steps = Array.from({ length: CAP + 5 }, () => step())
    for (const entry of steps) pushVisualUndo(runtime, entry.value)
    expect(tokens).toHaveLength(CAP + 5)

    // The five oldest are gone: the mutation can no longer resolve a step, so
    // the command reports failure instead of retaining the closure forever.
    for (const token of tokens.slice(0, 5)) {
      expect(run(token)).toBe(false)
      expect(steps[tokens.indexOf(token)]!.undone()).toBe(0)
    }
    // Everything inside the bound still undoes and redoes.
    for (const token of tokens.slice(5)) {
      expect(run(token, 'undo')).toBe(true)
      expect(run(token, 'redo')).toBe(true)
    }
    for (const entry of steps.slice(5)) {
      expect(entry.undone()).toBe(1)
      expect(entry.redone()).toBe(1)
    }
  })

  it('never grows past the bound across a long edit session', () => {
    const { runtime, tokens, run } = fakeRuntime()
    for (let index = 0; index < CAP * 20; index += 1) {
      pushVisualUndo(runtime, step().value)
    }
    // Only the newest CAP tokens can still resolve, so a session that makes
    // thousands of chart/shape edits retains a fixed window, not all of them.
    const resolvable = tokens.filter((token) => run(token)).length
    expect(resolvable).toBeLessThanOrEqual(CAP)
    expect(resolvable).toBe(CAP)
  })

  it('evicts by last use, so a step being stepped through is not the one lost', () => {
    const { runtime, tokens, run } = fakeRuntime()
    for (let index = 0; index < CAP; index += 1) pushVisualUndo(runtime, step().value)
    // Touch the oldest token: Map re-insertion makes it most-recently-used.
    expect(run(tokens[0]!)).toBe(true)
    pushVisualUndo(runtime, step().value)

    // The new step pushed everything else down by one, so the *second* entry
    // is now the least-recently-used and falls out — not the token just used.
    expect(run(tokens[1]!)).toBe(false)
    expect(run(tokens[0]!)).toBe(true)
    expect(run(tokens[tokens.length - 1]!)).toBe(true)
  })
})
