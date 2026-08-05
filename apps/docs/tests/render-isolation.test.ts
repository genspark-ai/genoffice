import { describe, expect, it, vi } from 'vitest'
import { act, createElement, useState } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { tableModelToPmNode } from '../src/renderer/editor/convert'
import { cachedByDoc } from '../src/renderer/doc-cache'
import { shallowEqual, useShallowStable, useStableCallbacks } from '../src/renderer/use-stable'
import {
  computeFormatState,
  EMPTY_FORMAT_STATE,
  type RibbonFormatState,
} from '../src/renderer/components/ribbon-format-state'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { CommentsPanel } from '../src/renderer/components/CommentsPanel'
// spy on a function the ribbon home tab calls on every render → precise re-render detector
vi.mock('../src/renderer/font-list', { spy: true })
import { fontFamiliesFor } from '../src/renderer/font-list'

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text: 'hello world plain' }] },
        { type: 'docHeading', attrs: { level: 2 }, content: [{ type: 'text', text: 'Title' }] },
        {
          type: 'docParagraph',
          content: [{ type: 'text', text: 'bold', marks: [{ type: 'bold' }] }],
        },
      ],
    },
  })
}

function mount(element: React.ReactElement): { root: Root; container: HTMLElement } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  return { root, container }
}

describe('cachedByDoc', () => {
  it('computes once per doc reference and recomputes on a new doc', () => {
    const editor = makeEditor()
    const compute = vi.fn((d: PmNode) => d.textContent.length)
    const cached = cachedByDoc(compute)
    const doc1 = editor.state.doc
    expect(cached(doc1)).toBe(doc1.textContent.length)
    expect(cached(doc1)).toBe(doc1.textContent.length)
    expect(compute).toHaveBeenCalledTimes(1)
    editor.commands.insertContentAt(2, 'x')
    const doc2 = editor.state.doc
    expect(doc2).not.toBe(doc1)
    cached(doc2)
    expect(compute).toHaveBeenCalledTimes(2)
    editor.destroy()
  })
})

describe('shallowEqual / useShallowStable', () => {
  it('compares own enumerable fields with Object.is', () => {
    expect(shallowEqual({ a: 1, b: null }, { a: 1, b: null })).toBe(true)
    expect(shallowEqual({ a: 1 }, { a: 2 })).toBe(false)
    expect(shallowEqual({ a: 1 }, { a: 1, b: 1 })).toBe(false)
    const o = {}
    expect(shallowEqual({ a: o }, { a: o })).toBe(true)
    expect(shallowEqual({ a: {} }, { a: {} })).toBe(false)
  })

  it('keeps the previous reference while values are equal', () => {
    const seen: Array<Record<string, unknown>> = []
    let setInput: (v: Record<string, unknown>) => void = () => {}
    function Probe() {
      const [input, set] = useState<Record<string, unknown>>({ a: 1 })
      setInput = set
      seen.push(useShallowStable(input))
      return null
    }
    const { root } = mount(createElement(Probe))
    act(() => setInput({ a: 1 }))
    act(() => setInput({ a: 2 }))
    expect(seen[1]).toBe(seen[0])
    expect(seen[2]).not.toBe(seen[0])
    expect(seen[2]).toEqual({ a: 2 })
    act(() => root.unmount())
  })
})

describe('useStableCallbacks', () => {
  it('returns stable identities that dispatch into the latest closure', () => {
    const bundles: Array<{ get: () => number }> = []
    let bump: () => void = () => {}
    function Probe() {
      const [n, setN] = useState(0)
      bump = () => setN((v) => v + 1)
      bundles.push(useStableCallbacks({ get: () => n }))
      return null
    }
    const { root } = mount(createElement(Probe))
    act(() => bump())
    expect(bundles[1]).toBe(bundles[0])
    expect(bundles[1].get).toBe(bundles[0].get)
    expect(bundles[1].get()).toBe(1)
    act(() => root.unmount())
  })
})

describe('computeFormatState', () => {
  it('returns the shared empty state without an editor', () => {
    expect(computeFormatState(null)).toBe(EMPTY_FORMAT_STATE)
  })

  it('is shallow-stable across caret moves inside uniform text', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection(3)
    const a = computeFormatState(editor)
    editor.commands.setTextSelection(8)
    const b = computeFormatState(editor)
    expect(a).not.toBe(b)
    expect(shallowEqual(a, b)).toBe(true)
    editor.destroy()
  })

  it('tracks mark, heading and doc changes', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection(3)
    const plain = computeFormatState(editor)
    expect(plain.bold).toBe(false)
    expect(plain.headingLevel).toBeNull()
    expect(plain.docEmpty).toBe(false)
    expect(plain.doc).toBe(editor.state.doc)

    editor.commands.setTextSelection(28) // inside the bold run
    const bold = computeFormatState(editor)
    expect(bold.bold).toBe(true)
    expect(shallowEqual(plain, bold)).toBe(false)

    editor.commands.setTextSelection(21) // inside the heading
    const heading = computeFormatState(editor)
    expect(heading.headingLevel).toBe(2)

    editor.commands.insertContentAt(3, 'x')
    const edited = computeFormatState(editor)
    expect(edited.doc).not.toBe(plain.doc)
    editor.destroy()
  })

  it('captures table context at the caret', () => {
    const editor = makeEditor()
    const table = {
      rows: [
        [{ paras: ['a'] }, { paras: ['b'] }],
        [{ paras: ['c'] }, { paras: ['d'] }],
      ],
      colWidthsPct: [50, 50],
    }
    editor.commands.insertContentAt(editor.state.doc.content.size, tableModelToPmNode(table))
    let cellPos = -1
    editor.state.doc.descendants((node, pos) => {
      if (cellPos === -1 && ['docTableCell', 'docTableHeader'].includes(node.type.name)) {
        cellPos = pos
      }
    })
    expect(cellPos).toBeGreaterThan(0)
    editor.commands.setTextSelection(cellPos + 2)
    const fs = computeFormatState(editor)
    expect(fs.inTable).toBe(true)
    expect(fs.cellKey).not.toBeNull()
    editor.destroy()
  })
})

/* ---- memoized ribbon: full-prop smoke render + skip on unchanged formatState ---- */

const noop = () => {}

function ribbonProps(editor: Editor, formatState: RibbonFormatState) {
  return {
    editor,
    formatState,
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

describe('Ribbon render isolation', () => {
  it('is wrapped in React.memo', () => {
    expect((Ribbon as unknown as { $$typeof: symbol }).$$typeof).toBe(Symbol.for('react.memo'))
  })

  it('skips re-render when formatState and callbacks are unchanged', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection(3)
    const fs1 = computeFormatState(editor)
    editor.commands.setTextSelection(28)
    const fsBold = computeFormatState(editor)

    const renders = () => vi.mocked(fontFamiliesFor).mock.calls.length
    let bump: () => void = () => {}
    let setFs: (fs: RibbonFormatState) => void = () => {}
    const base = ribbonProps(editor, fs1)
    function Host() {
      const [, setTick] = useState(0)
      const [fs, set] = useState(fs1)
      bump = () => setTick((t) => t + 1)
      setFs = set
      return createElement(Ribbon, { ...base, formatState: fs })
    }
    const { root, container } = mount(createElement(Host))
    expect(container.querySelector('.ribbon')).not.toBeNull()
    const afterMount = renders()
    expect(afterMount).toBeGreaterThan(0)

    act(() => bump()) // parent re-render, identical props: memo must bail out
    expect(renders()).toBe(afterMount)

    act(() => setFs(fsBold)) // format actually changed: ribbon re-renders
    expect(renders()).toBeGreaterThan(afterMount)
    expect(container.querySelector('.rb-icon.active')).not.toBeNull()

    act(() => root.unmount())
    editor.destroy()
  })
})

describe('CommentsPanel anchor scan', () => {
  it('shows anchor text collected in a single DOM pass', () => {
    const editor = makeEditor()
    const pm = document.createElement('div')
    pm.className = 'ProseMirror'
    const span = document.createElement('span')
    span.className = 'doc-comment'
    span.dataset.commentIds = 'c1'
    span.textContent = 'anchored text'
    pm.appendChild(span)
    document.body.appendChild(pm)

    const comments = [{ id: 'c1', author: 'A', text: 'note', done: false }]
    const { root, container } = mount(
      createElement(CommentsPanel, {
        comments: comments as never,
        docNode: editor.state.doc,
        composing: false,
        onSubmitNew: noop,
        onReply: noop,
        onResolve: noop,
        onCancelNew: noop,
        onDelete: noop,
        onClose: noop,
      }),
    )
    expect(container.querySelector('.comment-anchor')?.textContent).toContain('anchored text')
    act(() => root.unmount())
    pm.remove()
    editor.destroy()
  })
})
