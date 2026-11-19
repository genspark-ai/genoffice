// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'

import { installContextSubmenuReopenFix } from '../src/renderer/context-submenu-reopen-fix'

const frames = (n: number) =>
  new Promise<void>((resolve) => {
    const step = (left: number) =>
      left === 0 ? resolve() : requestAnimationFrame(() => step(left - 1))
    step(n)
  })

function mountSubmenu(visibility: 'hidden' | 'visible'): HTMLElement {
  const submenu = document.createElement('div')
  submenu.setAttribute('data-u-context-menu-submenu', 'true')
  submenu.style.visibility = visibility
  document.body.appendChild(submenu)
  return submenu
}

// Univer's positioning effect: a capture scroll listener on window that
// re-measures and flips the submenu visible.
function univerPositioningEffect(submenu: HTMLElement): { calls: number; dispose(): void } {
  const state = { calls: 0, dispose: () => {} }
  const handler = () => {
    state.calls += 1
    submenu.style.visibility = 'visible'
  }
  window.addEventListener('scroll', handler, true)
  state.dispose = () => window.removeEventListener('scroll', handler, true)
  return state
}

let disposers: Array<() => void> = []
afterEach(() => {
  for (const dispose of disposers) dispose()
  disposers = []
  document.body.innerHTML = ''
})

describe('installContextSubmenuReopenFix', () => {
  it('nudges a submenu that stays hidden past its first frame', async () => {
    const fix = installContextSubmenuReopenFix()
    disposers.push(fix.dispose)
    const submenu = mountSubmenu('hidden')
    const effect = univerPositioningEffect(submenu)
    disposers.push(effect.dispose)

    await frames(3)

    expect(effect.calls).toBe(1)
    expect(submenu.style.visibility).toBe('visible')
  })

  it('leaves a submenu alone that shows itself on the next frame', async () => {
    const fix = installContextSubmenuReopenFix()
    disposers.push(fix.dispose)
    const submenu = mountSubmenu('hidden')
    const scrolls = vi.fn()
    window.addEventListener('scroll', scrolls, true)
    disposers.push(() => window.removeEventListener('scroll', scrolls, true))
    requestAnimationFrame(() => {
      submenu.style.visibility = 'visible'
    })

    await frames(4)

    expect(scrolls).not.toHaveBeenCalled()
  })

  it('nudges a stuck submenu once, then again after it has shown', async () => {
    const fix = installContextSubmenuReopenFix()
    disposers.push(fix.dispose)
    const submenu = mountSubmenu('hidden')
    const scrolls = vi.fn()
    window.addEventListener('scroll', scrolls, true)
    disposers.push(() => window.removeEventListener('scroll', scrolls, true))

    await frames(3)
    submenu.style.left = '10px'
    await frames(3)
    expect(scrolls).toHaveBeenCalledTimes(1)

    submenu.style.visibility = 'visible'
    await frames(3)
    submenu.style.visibility = 'hidden'
    await frames(3)
    expect(scrolls).toHaveBeenCalledTimes(2)
  })

  it('stops watching after dispose', async () => {
    const fix = installContextSubmenuReopenFix()
    const scrolls = vi.fn()
    window.addEventListener('scroll', scrolls, true)
    disposers.push(() => window.removeEventListener('scroll', scrolls, true))
    fix.dispose()
    mountSubmenu('hidden')

    await frames(3)

    expect(scrolls).not.toHaveBeenCalled()
  })
})
