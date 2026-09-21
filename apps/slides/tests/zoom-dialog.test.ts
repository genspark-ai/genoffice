import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { RenderSlide } from '@genoffice/pptx-render'

vi.mock('../src/renderer/SlideThumb', () => ({ SlideThumb: () => null }))

import { ZoomDialog } from '../src/renderer/components/ZoomDialog'
import { t } from '../src/renderer/i18n/locale'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const slides = [0, 1, 2].map(
  () => ({ widthPx: 1280, heightPx: 720, scale: 1, nodes: [] }) as unknown as RenderSlide,
)
const sections = [{ id: '{A}', name: 'Intro', slideIndices: [1, 2] }]

const roots: Array<{ root: Root; container: HTMLElement }> = []
afterEach(() => {
  for (const { root, container } of roots.splice(0)) {
    act(() => root.unmount())
    container.remove()
  }
})

async function mount(mode: 'slide' | 'section' | 'summary', secs = sections) {
  const onInsert = vi.fn()
  const onClose = vi.fn()
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  await act(async () => {
    root.render(
      createElement(ZoomDialog, {
        mode,
        slides,
        images: new Map(),
        sections: secs,
        currentSlide: 0,
        onInsert,
        onClose,
      }),
    )
  })
  roots.push({ root, container })
  const dlg = container.querySelector('.zoom-dlg') as HTMLElement
  const items = () => Array.from(dlg.querySelectorAll<HTMLButtonElement>('.zoom-dlg-item'))
  const insertBtn = () => dlg.querySelector<HTMLButtonElement>('.modal-actions .primary')!
  const count = () => dlg.querySelector('.zoom-dlg-count')?.textContent
  const click = (el: HTMLElement, toggle = false) =>
    act(() => {
      el.dispatchEvent(new MouseEvent('click', { bubbles: true, metaKey: toggle, ctrlKey: toggle }))
    })
  return { dlg, items, insertBtn, count, click, onInsert, onClose }
}

describe('ZoomDialog', () => {
  it('slide mode: click selects one, toggle-click adds, Insert waits for a pick', async () => {
    const d = await mount('slide')
    expect(d.dlg.querySelector('h2')?.textContent).toBe(t('ribbonZoomSlide'))
    expect(d.items()).toHaveLength(3)
    expect(d.items()[0]!.disabled).toBe(true)
    expect(d.items()[0]!.textContent).toContain(t('ribbonCurrentSlideSuffix').trim())
    expect(d.insertBtn().disabled).toBe(true)
    expect(d.count()).toBe(t('ribbonZoomSelectedSlides', { n: 0 }))

    d.click(d.items()[2]!)
    expect(d.count()).toBe(t('ribbonZoomSelectedSlides', { n: 1 }))
    d.click(d.items()[1]!, true)
    expect(d.count()).toBe(t('ribbonZoomSelectedSlides', { n: 2 }))
    expect(d.items().map((b) => b.classList.contains('selected'))).toEqual([false, true, true])
    d.click(d.items()[2]!, true)
    expect(d.count()).toBe(t('ribbonZoomSelectedSlides', { n: 1 }))
    d.click(d.items()[2]!)
    expect(d.items().map((b) => b.classList.contains('selected'))).toEqual([false, false, true])
    d.click(d.items()[1]!, true)

    expect(d.insertBtn().disabled).toBe(false)
    d.click(d.insertBtn())
    expect(d.onInsert).toHaveBeenCalledWith([1, 2])
  })

  it('section mode: one thumbnail per section captioned with its head slide', async () => {
    const d = await mount('section')
    expect(d.items().map((b) => b.querySelector('.zoom-dlg-caption')?.textContent)).toEqual([
      t('ribbonZoomSectionItem', { n: 1, k: 1, name: t('appSectionDefault') }),
      t('ribbonZoomSectionItem', { n: 2, k: 2, name: 'Intro' }),
    ])
    expect(d.count()).toBe(t('ribbonZoomSelectedSections', { n: 0 }))
    d.click(d.items()[1]!)
    d.click(d.insertBtn())
    expect(d.onInsert).toHaveBeenCalledWith([1])
  })

  it('summary mode preselects every section head and nothing without sections', async () => {
    const d = await mount('summary')
    expect(d.items().map((b) => b.classList.contains('selected'))).toEqual([true, true, false])
    expect(d.items()[0]!.disabled).toBe(false)
    d.click(d.insertBtn())
    expect(d.onInsert).toHaveBeenCalledWith([0, 1])

    const bare = await mount('summary', [])
    expect(bare.items().some((b) => b.classList.contains('selected'))).toBe(false)
    expect(bare.insertBtn().disabled).toBe(true)
    bare.click(bare.dlg.querySelector<HTMLButtonElement>('.modal-actions button')!)
    expect(bare.onClose).toHaveBeenCalled()
  })
})
