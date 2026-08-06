// Regression (feedback row 53): the home-tab style gallery clipped overflowing
// cards with no way to reach them. The gallery must show a "more styles"
// expander whenever cards are cut off, and the expander grid must list every style.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { Ribbon } from '../src/renderer/components/Ribbon'

class FakeResizeObserver {
  static instances: FakeResizeObserver[] = []
  targets = new Set<Element>()
  constructor(private cb: ResizeObserverCallback) {
    FakeResizeObserver.instances.push(this)
  }
  observe(el: Element) {
    this.targets.add(el)
  }
  unobserve(el: Element) {
    this.targets.delete(el)
  }
  disconnect() {
    this.targets.clear()
  }
  static fire(el: Element) {
    for (const inst of FakeResizeObserver.instances) {
      if (inst.targets.has(el)) {
        inst.cb([{ target: el } as ResizeObserverEntry], inst as unknown as ResizeObserver)
      }
    }
  }
}

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'hello world' }] }],
    },
  })
}

const noop = () => {}

function ribbonProps(editor: Editor) {
  return {
    editor,
    formatState: computeFormatState(editor),
    hasDoc: true,
    blocks: [],
    onOpen: noop,
    onSave: noop,
    onSaveAs: noop,
    showAi: false,
    onToggleAi: noop,
    section: null,
    onSection: noop,
    activeSection: null,
    onInsertSectionBreak: noop,
    pageColor: null,
    onPageColor: noop,
    watermark: null,
    onWatermark: noop,
    themeFonts: null,
    onThemeFonts: noop,
    themeColors: null,
    onThemeColors: noop,
    inkTool: 'select' as const,
    onInkTool: noop,
    inkPen: { color: 'C00000', width: 2 },
    onInkPen: noop,
    inkHighlighter: { color: 'FFFF00', width: 10 },
    onInkHighlighter: noop,
    inkCount: 0,
    onInkClearAll: noop,
    onInsertNote: noop,
    sources: [],
    onAddSource: noop,
    zoom: 100,
    onZoom: noop,
    onZoomFit: noop,
    darkCanvas: false,
    onDarkCanvas: noop,
    onAiPreset: noop,
    header: null,
    onHeader: noop,
    onPageNumFormat: noop,
    onInsertField: noop,
    footer: null,
    onFooter: noop,
    titlePg: false,
    onTitlePg: noop,
    evenOddHf: false,
    onEvenOddHf: noop,
    showMarks: false,
    onShowMarks: noop,
    showRuler: false,
    onShowRuler: noop,
    showNav: false,
    onShowNav: noop,
    commentCount: 0,
    onShowComments: noop,
    canComment: false,
    onNewComment: noop,
    trackChanges: false,
    onTrackChanges: noop,
    revisionDisplay: 'all' as const,
    onRevisionDisplay: noop,
    revisionCount: 0,
    onAcceptRevision: noop,
    onRejectRevision: noop,
    onGotoRevision: noop,
    isProtected: false,
    onToggleProtection: noop,
    onCompare: noop,
    filePath: null,
    viewMode: 'print' as const,
    onViewMode: noop,
    readMode: false,
    onReadMode: noop,
    showGrid: false,
    onShowGrid: noop,
    splitView: false,
    onSplitView: noop,
    onPagePreview: noop,
  }
}

function setWidths(el: Element, scrollWidth: number, clientWidth: number) {
  Object.defineProperty(el, 'scrollWidth', { configurable: true, value: scrollWidth })
  Object.defineProperty(el, 'clientWidth', { configurable: true, value: clientWidth })
}

describe('style gallery overflow', () => {
  let editor: Editor
  let root: Root
  let container: HTMLElement

  beforeEach(() => {
    vi.stubGlobal('ResizeObserver', FakeResizeObserver)
    editor = makeEditor()
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
    act(() => root.render(createElement(Ribbon, ribbonProps(editor))))
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
    editor.destroy()
    FakeResizeObserver.instances.length = 0
    vi.unstubAllGlobals()
  })

  it('shows no expander while every card fits', () => {
    expect(container.querySelector('.style-gallery')).not.toBeNull()
    expect(container.querySelectorAll('.style-gallery .style-card').length).toBe(6)
    expect(container.querySelector('.style-gallery-more')).toBeNull()
  })

  it('shows the expander when cards are clipped and hides it when they fit again', () => {
    const gallery = container.querySelector('.style-gallery')!
    setWidths(gallery, 460, 200)
    act(() => FakeResizeObserver.fire(gallery))
    expect(container.querySelector('.style-gallery-more')).not.toBeNull()

    setWidths(gallery, 460, 460)
    act(() => FakeResizeObserver.fire(gallery))
    expect(container.querySelector('.style-gallery-more')).toBeNull()
  })

  it('expander opens a grid with every style, and picking one applies it and closes', () => {
    const gallery = container.querySelector('.style-gallery')!
    setWidths(gallery, 460, 200)
    act(() => FakeResizeObserver.fire(gallery))

    const more = container.querySelector<HTMLButtonElement>('.style-gallery-more')!
    expect(more.getAttribute('aria-label')).toBeTruthy()
    act(() => more.click())

    const menu = container.querySelector('.style-gallery-menu')!
    const cards = menu.querySelectorAll<HTMLButtonElement>('.style-card')
    // 4 paragraph styles + 2 preset character styles: nothing is dropped
    expect(cards.length).toBe(6)

    editor.commands.setTextSelection(3)
    act(() => cards[1].click()) // Heading 1
    expect(editor.isActive('docHeading', { level: 1 })).toBe(true)
    expect(container.querySelector('.style-gallery-menu')).toBeNull()
  })
})
