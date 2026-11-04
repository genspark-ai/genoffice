// Home > Styles (Word for Mac): the Styles button drops down the quick style
// gallery; a card applies its style and closes the menu. No inline card row.
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { Ribbon } from '../src/renderer/components/Ribbon'

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
    onConvert: noop,
    onNameWithAi: noop,
    showAi: false,
    onToggleAi: noop,
    section: null,
    onSection: noop,
    activeSection: null,
    onInsertSectionBreak: noop,
    onPaperSizeAll: noop,
    onPageBorderAll: noop,
    mirrorMargins: false,
    onMirrorMargins: noop,
    pageColor: null,
    onPageColor: noop,
    watermark: null,
    watermarkStyle: null,
    watermarkPicture: null,
    watermarkDocImage: null,
    onWatermark: noop,
    themeFonts: null,
    onThemeFonts: noop,
    themeColors: null,
    onThemeColors: noop,
    inkTool: 'select' as const,
    onInkTool: noop,
    inkPresets: [{ id: 'p1', kind: 'pen' as const, color: 'C00000', width: 2 }],
    inkActiveId: 'p1',
    onInkPresets: noop,
    onInkActivate: noop,
    inkCount: 0,
    onInkClearAll: noop,
    onInsertNote: noop,
    sources: [],
    onAddSource: noop,
    zoom: 100,
    onZoom: noop,
    onZoomFit: noop,
    onZoomDialog: noop,
    onTrackChangesOptions: noop,
    darkPage: false,
    onDarkPage: noop,
    onAiPreset: noop,
    header: null,
    onHeader: noop,
    onPageNumFormat: noop,
    onInsertField: noop,
    footer: null,
    onFooter: noop,
    hfEditing: null,
    onHfAction: noop,
    onHfEdit: noop,
    showMarks: false,
    onShowMarks: noop,
    showRuler: false,
    onShowRuler: noop,
    showNav: false,
    onShowNav: noop,
    commentCount: 0,
    openCommentCount: 0,
    resolvedCommentCount: 0,
    onShowComments: noop,
    showComments: false,
    onToggleComments: noop,
    commentsMode: 'contextual' as const,
    onCommentsMode: noop,
    canComment: false,
    onNewComment: noop,
    commentAtCaret: false,
    onDeleteComment: noop,
    onDeleteAllComments: noop,
    onGotoComment: noop,
    trackChanges: false,
    onTrackChanges: noop,
    spellcheck: true,
    onSpellcheck: noop,
    revisionDisplay: 'all' as const,
    onRevisionDisplay: noop,
    balloonMode: 'comments' as const,
    onBalloonMode: noop,
    revisionCount: 0,
    onAcceptRevision: noop,
    onRejectRevision: noop,
    onGotoRevision: noop,
    revisionFilter: { insDel: true, format: true, hiddenAuthors: [] },
    onRevisionFilter: noop,
    revisionAuthors: [],
    showReviewPane: false,
    onToggleReviewPane: noop,
    isProtected: false,
    commentsAllowed: false,
    trackChangesForced: false,
    protectActive: false,
    onProtectDoc: noop,
    onCompare: noop,
    mergeSource: null,
    onPickMergeDataSource: noop,
    onMergeMailings: async () => ({ ok: true }),
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

describe('style gallery button', () => {
  let editor: Editor
  let root: Root
  let container: HTMLElement

  beforeEach(() => {
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
  })

  it('renders only the Styles and Styles Pane buttons, no inline cards', () => {
    expect(container.querySelector('.rb-styles-gallery')).not.toBeNull()
    expect(container.querySelector('.rb-styles-pane')).not.toBeNull()
    expect(container.querySelectorAll('.style-card').length).toBe(0)
    expect(container.querySelector('.style-gallery-menu')).toBeNull()
  })

  it('opens a grid with every style, and picking one applies it and closes', () => {
    const button = container.querySelector<HTMLButtonElement>('.rb-styles-gallery')!
    expect(button.getAttribute('aria-expanded')).toBe('false')
    act(() => button.click())
    expect(button.getAttribute('aria-expanded')).toBe('true')

    const menu = container.querySelector('.style-gallery-menu')!
    const cards = menu.querySelectorAll<HTMLButtonElement>('.style-card')
    // the four fallback paragraph styles: nothing is dropped
    expect(cards.length).toBe(4)
    expect(menu.querySelector('.style-gallery-clear')).not.toBeNull()

    editor.commands.setTextSelection(3)
    act(() => cards[1].click()) // Heading 1
    expect(editor.isActive('docHeading', { level: 1 })).toBe(true)
    expect(container.querySelector('.style-gallery-menu')).toBeNull()
  })

  it('clicking the button again closes the menu', () => {
    const button = container.querySelector<HTMLButtonElement>('.rb-styles-gallery')!
    act(() => button.click())
    expect(container.querySelector('.style-gallery-menu')).not.toBeNull()
    act(() => button.click())
    expect(container.querySelector('.style-gallery-menu')).toBeNull()
  })
})
