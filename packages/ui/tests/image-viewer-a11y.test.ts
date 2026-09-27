/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ImageViewer, type ImageViewerLabels } from '../src/image-viewer'
import { IMAGE_VIEWER_TITLES } from '../src/strings-image-viewer'
import { LANGS } from '@genoffice/i18n'

const LABELS: ImageViewerLabels = {
  zoomIn: 'Zoom in',
  zoomOut: 'Zoom out',
  actualSize: 'Actual size',
  fitToWindow: 'Fit to window',
  save: 'Save',
  close: 'Close',
}

let host: HTMLDivElement
let root: ReturnType<typeof createRoot>
let trigger: HTMLButtonElement
let onClose: ReturnType<typeof vi.fn>

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  onClose = vi.fn()
  host = document.createElement('div')
  document.body.append(host)
  trigger = document.createElement('button')
  document.body.append(trigger)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  trigger.remove()
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

function mountViewer(
  lang?: Parameters<typeof ImageViewer>[0]['lang'],
  extra: { onSave?: () => void } = {},
) {
  trigger.focus()
  act(() =>
    root.render(
      createElement(ImageViewer, {
        src: 'data:image/png;base64,',
        alt: 'a picture',
        labels: LABELS,
        onClose,
        lang,
        ...extra,
      }),
    ),
  )
  return host
}

const dialog = (scope: HTMLElement): HTMLElement =>
  scope.querySelector('[role="dialog"]') as HTMLElement

const control = (scope: HTMLElement, name: string): HTMLButtonElement =>
  scope.querySelector(`button[aria-label="${name}"]`) as HTMLButtonElement

describe('shared ImageViewer modal focus', () => {
  it('names the dialog', () => {
    mountViewer()
    expect(dialog(host).getAttribute('aria-label')).toBe('Image viewer')
    expect(dialog(host).getAttribute('aria-modal')).toBe('true')
  })

  it('localizes the dialog name', () => {
    mountViewer('zh')
    expect(dialog(host).getAttribute('aria-label')).toBe('\u56fe\u7247\u67e5\u770b\u5668')
    mountViewer('ja')
    expect(dialog(host).getAttribute('aria-label')).toBe(IMAGE_VIEWER_TITLES.ja)
  })

  it('names the dialog in every supported language', () => {
    for (const lang of LANGS) {
      expect(IMAGE_VIEWER_TITLES[lang], lang).toBeTruthy()
    }
    expect(new Set(Object.values(IMAGE_VIEWER_TITLES)).size).toBeGreaterThan(1)
  })

  it('moves focus to the close control on open', () => {
    mountViewer()
    expect(document.activeElement).toBe(control(host, 'Close'))
  })

  it('restores focus to the invoking control on close', () => {
    mountViewer()
    act(() => root.unmount())
    expect(document.activeElement).toBe(trigger)
  })

  it('keeps Tab inside the viewer', () => {
    mountViewer(undefined, { onSave: () => {} })
    const focusables = [...host.querySelectorAll<HTMLElement>('button')]
    const last = focusables[focusables.length - 1]!
    last.focus()
    const e = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true })
    act(() => {
      dialog(host).dispatchEvent(e)
    })
    expect(dialog(host).contains(document.activeElement)).toBe(true)
    expect(document.activeElement).toBe(focusables[0])
  })

  it('wraps backwards from the first control to the last', () => {
    mountViewer()
    const focusables = [...host.querySelectorAll<HTMLElement>('button')]
    const first = focusables[0]!
    first.focus()
    const e = new KeyboardEvent('keydown', {
      key: 'Tab',
      shiftKey: true,
      bubbles: true,
      cancelable: true,
    })
    act(() => {
      dialog(host).dispatchEvent(e)
    })
    expect(document.activeElement).toBe(focusables[focusables.length - 1])
  })

  it('closes on Escape without reaching the editor behind', () => {
    mountViewer()
    const e = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })
    act(() => {
      control(host, 'Close').dispatchEvent(e)
    })
    expect(onClose).toHaveBeenCalledTimes(1)
    expect(e.defaultPrevented).toBe(true)
  })

  it('pulls focus back onto the backdrop when it drops to body', () => {
    mountViewer()
    act(() => {
      ;(document.activeElement as HTMLElement).blur()
    })
    expect(document.activeElement).toBe(dialog(host))
    const e = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })
    act(() => {
      dialog(host).dispatchEvent(e)
    })
    expect(onClose).toHaveBeenCalledTimes(1)
  })

  it('closes on Escape even when focus has escaped to body', () => {
    mountViewer()
    // leave the viewer with a relatedTarget (no body-blur refocus), then drop focus
    trigger.focus()
    trigger.blur()
    expect(document.activeElement).toBe(document.body)
    const e = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })
    act(() => {
      document.body.dispatchEvent(e)
    })
    expect(onClose).toHaveBeenCalledTimes(1)
    expect(e.defaultPrevented).toBe(true)
  })

  it('still zooms from the keyboard', () => {
    mountViewer()
    const zoom = (): string => host.querySelector('.gs-imgview-zoom')!.textContent!
    expect(zoom()).toBe('100%')
    act(() => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: '+' }))
    })
    expect(zoom()).toBe('125%')
    act(() => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: '0' }))
    })
    expect(zoom()).toBe('100%')
  })

  it('closes when the backdrop itself is clicked', () => {
    mountViewer()
    act(() => {
      dialog(host).dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(onClose).toHaveBeenCalledTimes(1)
  })
})
