// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { installEditorScrollFollow } from '../src/renderer/editor-scroll-follow'
import type { UniverRuntime } from '../src/renderer/univer-state'

/** The module resolves the bridge first, then the resize service. */
function runtimeWith(bridge: unknown, resize: unknown): UniverRuntime {
  const services = [bridge, resize]
  let next = 0
  return {
    univer: { __getInjector: () => ({ get: () => services[next++] }) },
  } as unknown as UniverRuntime
}

function nextFrame(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof requestAnimationFrame === 'function') requestAnimationFrame(() => resolve())
    else setTimeout(resolve, 0)
  })
}

function bridgeThat(visible: boolean, calls: string[]) {
  return {
    isVisible: () => ({ visible }),
    refreshEditCellPosition: () => calls.push('position'),
  }
}

function resizeThat(calls: string[]) {
  return { fitTextSize: () => calls.push('fit') }
}

describe('installEditorScrollFollow', () => {
  it('recomputes the layout and applies it after a wheel frame', async () => {
    const calls: string[] = []
    const follow = installEditorScrollFollow(
      runtimeWith(bridgeThat(true, calls), resizeThat(calls)),
    )
    document.dispatchEvent(new WheelEvent('wheel'))
    await nextFrame()
    expect(calls).toEqual(['position', 'fit'])
    follow.dispose()
  })

  it('does nothing when no editor is open', async () => {
    const calls: string[] = []
    const follow = installEditorScrollFollow(
      runtimeWith(bridgeThat(false, calls), resizeThat(calls)),
    )
    document.dispatchEvent(new WheelEvent('wheel'))
    await nextFrame()
    expect(calls).toEqual([])
    follow.dispose()
  })

  it('defers the refresh during a ctrl+wheel zoom gesture', async () => {
    const calls: string[] = []
    const follow = installEditorScrollFollow(
      runtimeWith(bridgeThat(true, calls), resizeThat(calls)),
    )
    document.dispatchEvent(new WheelEvent('wheel', { ctrlKey: true }))
    await nextFrame()
    // the zoom gesture must not be re-placed per frame; it settles later
    expect(calls).toEqual([])
    follow.dispose()
  })

  it('stops listening after dispose', async () => {
    const calls: string[] = []
    const follow = installEditorScrollFollow(
      runtimeWith(bridgeThat(true, calls), resizeThat(calls)),
    )
    follow.dispose()
    document.dispatchEvent(new WheelEvent('wheel'))
    await nextFrame()
    expect(calls).toEqual([])
  })

  it('follows the viewport of a workbook opened after install', async () => {
    const calls: string[] = []
    const listeners = new Map<string, () => void>()
    const viewportOf = (id: string) => ({
      onScrollAfter$: {
        subscribe: (fn: () => void) => {
          listeners.set(id, fn)
          return { unsubscribe: () => listeners.delete(id) }
        },
      },
    })
    const viewports: Record<string, unknown> = { placeholder: viewportOf('placeholder') }
    let active = 'placeholder'
    const renderManager = {
      getRenderUnitById: (id: string) => ({ scene: { getViewport: () => viewports[id] ?? null } }),
    }
    const services = [bridgeThat(true, calls), resizeThat(calls), renderManager]
    let next = 0
    const follow = installEditorScrollFollow({
      univer: { __getInjector: () => ({ get: () => services[next++] }) },
      univerAPI: { getActiveWorkbook: () => ({ getId: () => active }) },
    } as unknown as UniverRuntime)
    expect([...listeners.keys()]).toEqual(['placeholder'])

    viewports.file = viewportOf('file')
    active = 'file'
    document.dispatchEvent(new Event('pointerdown'))
    expect([...listeners.keys()]).toEqual(['file'])

    listeners.get('file')!()
    await nextFrame()
    expect(calls).toEqual(['position', 'fit'])
    follow.dispose()
    expect(listeners.size).toBe(0)
  })

  it('is a no-op when a service is missing', () => {
    expect(() =>
      installEditorScrollFollow(runtimeWith(undefined, undefined)).dispose(),
    ).not.toThrow()
  })
})
