// The docs AI panel adopts the shared responsive width rule from @genoffice/ui:
// auto width while the user never dragged (grows with the window, floored at the
// default), a dragged preference clamped into [default, cap] with legacy
// sub-default saves migrating up, and re-derivation on window resize.
import { beforeAll, afterEach, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((p) => [p.id, { apiKey: '', model: p.defaultModel }]),
  ) as AiSettings['providers'],
}

const PANEL_WIDTH_KEY = 'docs-ai-panel-width'

function createEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'EVs market research' }],
        },
      ],
    },
  })
}

function setWindowWidth(width: number): void {
  Object.defineProperty(window, 'innerWidth', { value: width, configurable: true })
}

/** the panel writes its display width onto the closest .ai-dock wrapper */
function mountInDock(editor: Editor): {
  dock: HTMLElement
  root: Root
  cleanup: () => void
} {
  const dock = document.createElement('div')
  dock.className = 'ai-dock'
  document.body.appendChild(dock)
  const root = createRoot(dock)
  act(() =>
    root.render(
      createElement(
        'div',
        null,
        createElement(AiPanel, {
          editor,
          blocks: [],
          settings,
          open: true,
          onExpand: () => {},
          onCollapse: () => {},
        }),
      ),
    ),
  )
  return {
    dock,
    root,
    cleanup: () => {
      act(() => root.unmount())
      dock.remove()
    },
  }
}

const panelWidth = (dock: HTMLElement): number =>
  Number(dock.style.getPropertyValue('--ai-panel-width').replace(/px$/, ''))

beforeAll(() => {
  // jsdom has no scrollTo; the panel auto-scrolls its chat log
  Element.prototype.scrollTo ??= () => {}
})

afterEach(() => {
  localStorage.removeItem(PANEL_WIDTH_KEY)
})

describe('AI panel width rule', () => {
  it('uses the auto width while the user never dragged (default window: floored at 360)', () => {
    setWindowWidth(1024)
    const editor = createEditor()
    const mounted = mountInDock(editor)
    expect(panelWidth(mounted.dock)).toBe(360)
    mounted.cleanup()
    editor.destroy()
  })

  it('auto-grows with the window but never below the 360 floor', () => {
    setWindowWidth(2000)
    const editor = createEditor()
    const mounted = mountInDock(editor)
    // ~22% of 2000 = 440, within [360, cap 720]
    expect(panelWidth(mounted.dock)).toBe(440)

    act(() => {
      setWindowWidth(3000)
      window.dispatchEvent(new Event('resize'))
    })
    // ~22% of 3000 = 660, still under the cap
    expect(panelWidth(mounted.dock)).toBe(660)

    act(() => {
      setWindowWidth(1024)
      window.dispatchEvent(new Event('resize'))
    })
    // shrinking the window never drops below 360
    expect(panelWidth(mounted.dock)).toBe(360)
    mounted.cleanup()
    editor.destroy()
  })

  it('restores a dragged width clamped to the window cap', () => {
    // 60% of 1024 = 614: a wider preference caps there
    localStorage.setItem(PANEL_WIDTH_KEY, '900')
    setWindowWidth(1024)
    const editor = createEditor()
    const mounted = mountInDock(editor)
    expect(panelWidth(mounted.dock)).toBe(614)

    // the same preference widens up to the absolute ceiling on a bigger window
    act(() => {
      setWindowWidth(2000)
      window.dispatchEvent(new Event('resize'))
    })
    expect(panelWidth(mounted.dock)).toBe(720)
    mounted.cleanup()
    editor.destroy()
  })

  it('migrates a legacy sub-360 saved width up to the floor', () => {
    localStorage.setItem(PANEL_WIDTH_KEY, '280')
    setWindowWidth(1024)
    const editor = createEditor()
    const mounted = mountInDock(editor)
    expect(panelWidth(mounted.dock)).toBe(360)
    mounted.cleanup()
    editor.destroy()
  })

  it('keeps treating a cleared preference as never-dragged (auto width)', () => {
    localStorage.removeItem(PANEL_WIDTH_KEY)
    setWindowWidth(2000)
    const editor = createEditor()
    const mounted = mountInDock(editor)
    expect(panelWidth(mounted.dock)).toBe(440)
    mounted.cleanup()
    editor.destroy()
  })
})
