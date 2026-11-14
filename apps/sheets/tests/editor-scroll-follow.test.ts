// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest'
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

  it('is a no-op when a service is missing', () => {
    expect(() =>
      installEditorScrollFollow(runtimeWith(undefined, undefined)).dispose(),
    ).not.toThrow()
  })
})
