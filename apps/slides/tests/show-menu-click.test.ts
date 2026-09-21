// Regression: choosing a slide show context-menu item must run that item only —
// the click used to bubble into the show root and advance the slide as well.
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { RenderSlide } from '@genoffice/pptx-render'

vi.mock('react-konva', () => {
  const stub = () => null
  return { Stage: stub, Layer: stub, Rect: stub, Group: stub, Image: stub, Path: stub }
})
vi.mock('../src/renderer/components/AnimatedSlide', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/renderer/components/AnimatedSlide')>()),
  AnimatedSlideStage: () => null,
}))
vi.mock('../src/renderer/SlideThumb', () => ({ SlideThumb: () => null }))

import { SlideShowView } from '../src/renderer/components/SlideShowView'
import { t } from '../src/renderer/i18n/locale'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const resolved =
  <T>(v: T) =>
  () =>
    Promise.resolve(v)
;(window as unknown as { slidesApi: unknown }).slidesApi = {
  getTransition: resolved('none'),
  getAnimations: resolved([]),
  getShapeKeys: resolved([]),
  getSlideLinks: resolved([]),
  getRunLinks: resolved([]),
  getMediaData: resolved(null),
  setShowFullScreen: resolved(undefined),
}

const slides = [0, 1, 2].map(
  () => ({ widthPx: 960, heightPx: 540, nodes: [], hidden: false }) as unknown as RenderSlide,
)

const roots: Array<{ root: Root; container: HTMLElement }> = []
afterEach(() => {
  for (const { root, container } of roots.splice(0)) {
    act(() => root.unmount())
    container.remove()
  }
})

async function mountShow(startAt = 0) {
  const onExit = vi.fn()
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  await act(async () => {
    root.render(createElement(SlideShowView, { slides, images: new Map(), startAt, onExit }))
  })
  // reveal happens on a requestAnimationFrame after the fullscreen IPC settles
  await act(async () => {
    await new Promise((r) => setTimeout(r, 50))
  })
  roots.push({ root, container })
  const show = container.querySelector('.slideshow') as HTMLElement
  expect(show.querySelector('.ss-counter')?.textContent).toBe(`${startAt + 1} / 3`)
  return { show, onExit }
}

const counter = (show: HTMLElement) => show.querySelector('.ss-counter')?.textContent

function openMenu(show: HTMLElement) {
  act(() => {
    show.dispatchEvent(
      new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 40, clientY: 40 }),
    )
  })
  const menu = show.querySelector('.ctx-menu') as HTMLElement
  expect(menu).not.toBeNull()
  return menu
}

function clickItem(menu: HTMLElement, label: string) {
  const btn = Array.from(menu.querySelectorAll<HTMLButtonElement>('.ctx-item')).find(
    (b) => b.querySelector('span')?.textContent === label,
  )!
  act(() => {
    btn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))
    btn.dispatchEvent(new MouseEvent('click', { bubbles: true, button: 0 }))
  })
}

describe('slide show context menu clicks', () => {
  it('Next advances exactly one slide', async () => {
    const { show } = await mountShow()
    clickItem(openMenu(show), t('paneShowMenuNext'))
    expect(counter(show)).toBe('2 / 3')
    expect(show.querySelector('.ctx-menu')).toBeNull()
  })

  it('Previous goes back one slide without a bounce forward', async () => {
    const { show } = await mountShow(2)
    clickItem(openMenu(show), t('paneShowMenuPrev'))
    expect(counter(show)).toBe('2 / 3')
  })

  it('Black Screen stays up after the click', async () => {
    const { show } = await mountShow()
    const menu = openMenu(show)
    act(() => {
      ;(menu.querySelector('.ctx-sub-host') as HTMLElement).dispatchEvent(
        new MouseEvent('mouseover', { bubbles: true }),
      )
    })
    clickItem(menu.querySelector('.ctx-submenu') as HTMLElement, t('paneShowMenuBlack'))
    expect(show.querySelector('.ss-black')).not.toBeNull()
    expect(counter(show)).toBe('1 / 3')
  })

  it('End Show exits once without advancing', async () => {
    const { show, onExit } = await mountShow()
    clickItem(openMenu(show), t('paneShowMenuEnd'))
    expect(onExit).toHaveBeenCalledTimes(1)
    expect(onExit).toHaveBeenCalledWith(0)
  })

  it('the click that dismisses the menu does not advance', async () => {
    const { show } = await mountShow()
    openMenu(show)
    act(() => {
      show.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))
      show.dispatchEvent(new MouseEvent('click', { bubbles: true, button: 0 }))
    })
    expect(show.querySelector('.ctx-menu')).toBeNull()
    expect(counter(show)).toBe('1 / 3')
    act(() => {
      show.dispatchEvent(new MouseEvent('click', { bubbles: true, button: 0 }))
    })
    expect(counter(show)).toBe('2 / 3')
  })

  it('a click while blacked out restores the slide instead of advancing', async () => {
    const { show } = await mountShow()
    act(() => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'b' }))
    })
    expect(show.querySelector('.ss-black')).not.toBeNull()
    act(() => {
      show.dispatchEvent(new MouseEvent('click', { bubbles: true, button: 0 }))
    })
    expect(show.querySelector('.ss-black')).toBeNull()
    expect(counter(show)).toBe('1 / 3')
  })
})
