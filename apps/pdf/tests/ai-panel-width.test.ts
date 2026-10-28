/**
 * AI panel width rule (shared foundation helpers): without a saved width the
 * panel follows the auto rule and re-derives it on window resize; a dragged
 * width re-clamps into the window cap; legacy sub-360 saved values migrate up
 * to the 360 floor; dragging persists the clamped width. The dock width var
 * (the 180ms slide animation's input) tracks the panel width throughout.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it } from 'vitest'
import { aiPanelAutoWidth, aiPanelClampWidth } from '@genoffice/ui'

import { AiPanel } from '../src/renderer/ai/AiPanel'
import type { PdfAppDeps } from '../src/renderer/ai/tools'

const api = {
  doc: () => null,
  fileName: () => 'sample.pdf',
  pageCount: () => 1,
  currentPage: () => 1,
  readOnly: () => false,
  ocrText: () => null,
  selection: () => null,
  pendingSummary: () => '',
  outline: () => null,
  searchIndex: () => null,
  isDeleted: () => false,
  gotoPage: () => true,
  addMarkup: () => {},
  annotationsOn: () => Promise.resolve({ threads: [], markups: [] }),
  annotationSummary: () => '',
  addNote: () => Promise.resolve(null),
  deletePages: () => Promise.resolve({ ok: true, pageCount: 1 }),
  rotatePages: () => Promise.resolve({ ok: true, pageCount: 1 }),
  setPageSize: () => Promise.resolve({ ok: true, pageCount: 1 }),
  cropPages: () => Promise.resolve({ ok: true, pageCount: 1 }),
  replacePages: () => Promise.resolve({ ok: true, pageCount: 1 }),
  extractPages: () => Promise.resolve({ ok: true, savedPath: '' }),
  splitPdf: () => Promise.resolve({ ok: true, savedDir: '', count: 0 }),
  splitPages: () => Promise.resolve({ ok: true, savedPath: '' }),
  mergePages: () => Promise.resolve({ ok: true, savedPath: '' }),
} as unknown as PdfAppDeps

const projectApi = {
  resolveChat: () => Promise.resolve({ projectId: 'p', chatId: 'c' }),
  appendChat: () => Promise.resolve(),
  loadChat: () => Promise.resolve([]),
  rebindChat: () => Promise.resolve(null),
}

function setWindowWidth(w: number): void {
  Object.defineProperty(window, 'innerWidth', { value: w, configurable: true, writable: true })
}

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  Object.defineProperty(window, 'projectApi', { value: projectApi, configurable: true })
  // jsdom implements neither chat scrolling nor pointer capture
  Element.prototype.scrollTo = () => {}
  const proto = HTMLElement.prototype as HTMLElement & {
    setPointerCapture?: (id: number) => void
  }
  if (typeof proto.setPointerCapture !== 'function') {
    proto.setPointerCapture = () => {}
  }
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
  localStorage.clear()
  setWindowWidth(1024)
})

async function renderPanel(): Promise<HTMLDivElement> {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(
      createElement(
        'div',
        { className: 'ai-dock' },
        createElement(AiPanel, {
          api,
          filePath: 'sample.pdf',
          onCollapse: () => {},
          preset: null,
        }),
      ),
    )
    await Promise.resolve()
  })
  return container.querySelector('.ai-dock')!
}

const dockVar = (dock: HTMLElement): string => dock.style.getPropertyValue('--ai-panel-width')

describe('AiPanel width rule', () => {
  it('shows the auto width until the user drags, re-derived on window resize', async () => {
    setWindowWidth(1920)
    const dock = await renderPanel()
    expect(dockVar(dock)).toBe(`${aiPanelAutoWidth(1920)}px`)
    expect(localStorage.getItem('pdf-ai-panel-width')).toBeNull()

    await act(async () => {
      setWindowWidth(1280)
      window.dispatchEvent(new Event('resize'))
      await Promise.resolve()
    })
    // 22% of 1280 sits below the 360 floor
    expect(dockVar(dock)).toBe(`${aiPanelAutoWidth(1280)}px`)
    expect(dockVar(dock)).toBe('360px')
  })

  it('migrates a legacy sub-360 saved width up to the floor', async () => {
    setWindowWidth(1920)
    localStorage.setItem('pdf-ai-panel-width', '200')
    const dock = await renderPanel()
    expect(dockVar(dock)).toBe('360px')
  })

  it('re-clamps a dragged width into the window cap and restores it when the window grows', async () => {
    setWindowWidth(800)
    localStorage.setItem('pdf-ai-panel-width', '500')
    const dock = await renderPanel()
    // cap = 60% of 800 = 480 < the saved 500
    expect(dockVar(dock)).toBe(`${aiPanelClampWidth(500, 800)}px`)
    expect(dockVar(dock)).toBe('480px')

    await act(async () => {
      setWindowWidth(1920)
      window.dispatchEvent(new Event('resize'))
      await Promise.resolve()
    })
    expect(dockVar(dock)).toBe('500px')
  })

  it('dragging the resizer clamps and persists the dragged width', async () => {
    setWindowWidth(1920)
    const dock = await renderPanel()
    const resizer = dock.querySelector<HTMLElement>('.ai-panel-resizer')!

    await act(async () => {
      resizer.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }))
      await Promise.resolve()
    })
    await act(async () => {
      window.dispatchEvent(new MouseEvent('pointermove', { clientX: 500 }))
      await Promise.resolve()
    })
    expect(dockVar(dock)).toBe('500px')

    await act(async () => {
      window.dispatchEvent(new MouseEvent('pointerup'))
      await Promise.resolve()
    })
    expect(localStorage.getItem('pdf-ai-panel-width')).toBe('500')
    expect(document.body.style.cursor).toBe('')
  })
})
