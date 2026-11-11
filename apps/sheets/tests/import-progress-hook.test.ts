// @vitest-environment jsdom
/**
 * The bar belongs to one import: it must not outlive it.
 *
 * The main process reports the position while an import runs and null when
 * it is over. Without the null, the next workbook opens showing the previous
 * file's percentage, with a cancel button that has nothing to cancel — the
 * worst kind of stale: it looks live.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'

import { useImportProgress, type ImportProgress } from '../src/renderer/import-progress'

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
  vi.unstubAllGlobals()
})

/** a preload bridge that hands out the listener and lets the test drive it */
function stubBridge(): { emit: (progress: ImportProgress | null) => void } {
  let listener: ((progress: ImportProgress | null) => void) | null = null
  vi.stubGlobal('window', {
    desktopApi: {
      onImportProgress: (given: (progress: ImportProgress | null) => void) => {
        listener = given
        return () => {
          listener = null
        }
      },
    },
  })
  return {
    emit: (progress) => {
      if (!listener) throw new Error('no listener registered')
      listener(progress)
    },
  }
}

describe('useImportProgress', () => {
  it('follows the import and forgets it when the import is over', async () => {
    const bridge = stubBridge()
    const seen: (ImportProgress | null)[] = []
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
    await act(async () => {
      root?.render(
        createElement(function Probe(): React.JSX.Element {
          const progress = useImportProgress()
          seen.push(progress)
          return createElement('span', null, progress ? progress.phase : 'none')
        }),
      )
    })
    expect(seen[seen.length - 1]).toBe(null)

    await act(async () => {
      bridge.emit({ phase: 'read', fraction: 0.5 })
    })
    expect(seen[seen.length - 1]).toEqual({ phase: 'read', fraction: 0.5 })

    await act(async () => {
      bridge.emit({ phase: 'write', fraction: 0.25 })
    })
    expect(seen[seen.length - 1]).toEqual({ phase: 'write', fraction: 0.25 })

    // the import ends: the bar goes with it
    await act(async () => {
      bridge.emit(null)
    })
    expect(seen[seen.length - 1]).toBe(null)
  })

  it('unsubscribes when the view goes away', async () => {
    let unsubscribed = false
    vi.stubGlobal('window', {
      desktopApi: {
        onImportProgress: () => () => {
          unsubscribed = true
        },
      },
    })
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
    await act(async () => {
      root?.render(
        createElement(function Probe(): React.JSX.Element {
          useImportProgress()
          return createElement('span')
        }),
      )
    })
    await act(async () => {
      root?.unmount()
    })
    root = null
    expect(unsubscribed).toBe(true)
  })
})
