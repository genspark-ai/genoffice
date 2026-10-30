/**
 * Review → Track Changes Options (Word's Advanced Options): the persisted
 * preference, the screen stylesheet it generates, the recorder switches
 * (moves / formatting) and the margin layer's geometry.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { syncMarginAnnotations } from '../src/renderer/editor/margin-annotations'
import { revisionDisplayState } from '../src/renderer/editor/marks'
import { collectRevisions, type TrackChangesStorage } from '../src/renderer/editor/revisions'
import { revisionAuthorCss } from '../src/renderer/editor/revision-filters'
import {
  DEFAULT_TRACK_CHANGES_OPTIONS,
  TRACK_CHANGES_OPTIONS_KEY,
  loadTrackChangesOptions,
  movesByAuthor,
  normalizeTrackChangesOptions,
  saveTrackChangesOptions,
  trackChangesOptionsCss,
  trackChangesOptionsState,
  type TrackChangesOptions,
} from '../src/renderer/editor/track-changes-options'

const withOpts = (patch: Partial<TrackChangesOptions>): TrackChangesOptions => ({
  ...DEFAULT_TRACK_CHANGES_OPTIONS,
  ...patch,
})

const para = (text: string) => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [{ type: 'text', text }],
})

function trackedEditor(storagePatch: Partial<TrackChangesStorage> = {}) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [para('one two three'), para('four')] },
  })
  Object.assign(editor.storage.trackChanges as TrackChangesStorage, {
    enabled: true,
    author: 'Alice',
    ...storagePatch,
  })
  return editor
}

afterEach(() => {
  localStorage.clear()
  trackChangesOptionsState.current = { ...DEFAULT_TRACK_CHANGES_OPTIONS }
})

describe('options store', () => {
  it('loads the defaults when nothing is stored and round-trips a save', () => {
    expect(loadTrackChangesOptions()).toEqual(DEFAULT_TRACK_CHANGES_OPTIONS)
    const o = withOpts({ delColor: 'red', balloonWidth: 260, trackMoves: false })
    saveTrackChangesOptions(o)
    expect(JSON.parse(localStorage.getItem(TRACK_CHANGES_OPTIONS_KEY)!)).toEqual(o)
    expect(loadTrackChangesOptions()).toEqual(o)
  })

  it('unknown or corrupt values fall back field by field', () => {
    localStorage.setItem(TRACK_CHANGES_OPTIONS_KEY, '{not json')
    expect(loadTrackChangesOptions()).toEqual(DEFAULT_TRACK_CHANGES_OPTIONS)
    const o = normalizeTrackChangesOptions({
      insMark: 'wavy',
      delColor: 'red',
      changedLines: 'top',
      cellIns: 'blue',
      balloonWidth: 9999,
      balloonMargin: 'left',
      trackFormatting: 'no',
    })
    expect(o.insMark).toBe('underline')
    expect(o.delColor).toBe('red')
    expect(o.changedLines).toBe('outside')
    expect(o.cellIns).toBe('blue')
    expect(o.balloonWidth).toBe(480)
    expect(o.balloonMargin).toBe('left')
    expect(o.trackFormatting).toBe(true)
  })
})

describe('generated stylesheet', () => {
  it('defaults generate nothing: styles.css already renders Word’s defaults', () => {
    expect(trackChangesOptionsCss(DEFAULT_TRACK_CHANGES_OPTIONS)).toBe('')
  })

  it('fixed colors become root custom properties, screen media only', () => {
    const css = trackChangesOptionsCss(
      withOpts({
        insColor: 'blue',
        delColor: 'red',
        changedLinesColor: 'black',
        commentColor: 'violet',
        cellIns: 'lightYellow',
      }),
    )
    expect(css.startsWith('@media screen{')).toBe(true)
    expect(css).toContain('--rev-ins-ink:var(--docs-paper-rev-color-blue)')
    expect(css).toContain('--rev-del-ink:var(--docs-paper-rev-color-red)')
    expect(css).toContain('--rev-bar-ink:var(--docs-paper-rev-color-black)')
    expect(css).toContain('--rev-comment-ink:var(--docs-paper-rev-color-violet)')
    expect(css).toContain('--rev-ins-bg:var(--docs-paper-rev-color-lightYellow)')
    expect(css).not.toContain('--rev-del-bg')
  })

  it('Color only deletions drop the strikethrough; bold insertions add weight and reset the line', () => {
    const css = trackChangesOptionsCss(withOpts({ delMark: 'color', insMark: 'bold' }))
    const del = /:where\(\.app\) \.doc-del:where\(:not\(\.doc-move-from\)\)\{([^}]*)\}/.exec(css)
    expect(del?.[1]).toContain('text-decoration:none')
    expect(del?.[1]).not.toContain('line-through')
    const ins = /:where\(\.app\) \.doc-ins:where\(:not\(\.doc-move-to\)\)\{([^}]*)\}/.exec(css)
    expect(ins?.[1]).toContain('font-weight:700')
    expect(ins?.[1]).toContain('text-decoration:none')
  })

  it('track moves off renders moves as deletion + insertion and lets the author ink apply', () => {
    const o = withOpts({ trackMoves: false })
    const css = trackChangesOptionsCss(o)
    expect(css).toContain(
      '.doc-move-from,.doc-move-to{--docs-paper-rev-ink:inherit;text-decoration-style:solid}',
    )
    expect(movesByAuthor(o)).toEqual({ from: true, to: true })
    const authorCss = revisionAuthorCss(['Ann', 'Bob'], [], movesByAuthor(o))
    expect(authorCss).toContain('.doc-ins[data-ins-author="Bob"],.doc-del[data-del-author="Bob"],')
    expect(revisionAuthorCss(['Ann', 'Bob'], [])).toContain(
      '.doc-ins[data-ins-author="Bob"]:not(.doc-move-to)',
    )
  })

  it('fixed move colors and marks, changed lines off', () => {
    const css = trackChangesOptionsCss(
      withOpts({ moveToColor: 'violet', moveFromMark: 'strike', changedLines: 'none' }),
    )
    expect(css).toContain(
      '.app .doc-move-to{--docs-paper-rev-ink:var(--docs-paper-rev-color-violet)}',
    )
    expect(css).toContain(
      ':where(.app) .doc-move-from{text-decoration:line-through;text-decoration-color:var(--docs-paper-rev-ink);text-decoration-style:solid}',
    )
    expect(css).toContain('.app .change-bar{display:none}')
  })
})

describe('recorder switches', () => {
  const cutThenPaste = (editor: Editor) => {
    const slice = editor.state.doc.slice(5, 9)
    editor.view.dispatch(editor.state.tr.delete(5, 9).setMeta('uiEvent', 'cut'))
    const end = editor.state.doc.content.size - 1
    editor.view.dispatch(editor.state.tr.insert(end, slice.content).setMeta('uiEvent', 'paste'))
  }

  it('track moves off: cut + paste records a deletion and an insertion', () => {
    const editor = trackedEditor({ trackMoves: false })
    cutThenPaste(editor)
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges.map((r) => r.kind)).toEqual(['del', 'ins'])
    expect(ranges.some((r) => r.moveId)).toBe(false)
    editor.destroy()
  })

  it('track moves on (default): the same edit pairs as a move', () => {
    const editor = trackedEditor()
    cutThenPaste(editor)
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toEqual(['moveFrom', 'moveTo'])
    editor.destroy()
  })

  it('track formatting off: a live bold leaves no rPrChange; on records one', () => {
    for (const trackFormatting of [false, true]) {
      const editor = trackedEditor({ trackFormatting })
      editor.commands.setTextSelection({ from: 1, to: 4 })
      editor.commands.setMark('bold')
      const kinds = collectRevisions(editor.state.doc).map((r) => r.kind)
      expect(kinds).toEqual(trackFormatting ? ['rPrChange'] : [])
      editor.destroy()
    }
  })
})

describe('margin layer geometry', () => {
  const rect = { top: 100, bottom: 116, left: 50, right: 200, width: 150, height: 16 }
  const wrapRect = { top: 0, bottom: 1000, left: 0, right: 800, width: 800, height: 1000 }
  const proto = Element.prototype
  const origRects = proto.getClientRects
  const origBox = proto.getBoundingClientRect
  beforeEach(() => {
    proto.getClientRects = function () {
      return [rect] as unknown as DOMRectList
    }
    proto.getBoundingClientRect = function () {
      return ((this as Element).classList.contains('page-wrap') ? wrapRect : rect) as DOMRect
    }
    revisionDisplayState.mode = 'all'
    revisionDisplayState.balloons = 'revisions'
    revisionDisplayState.balloonsActive = true
  })
  afterEach(() => {
    proto.getClientRects = origRects
    proto.getBoundingClientRect = origBox
    document.body.innerHTML = ''
    revisionDisplayState.balloons = 'comments'
    revisionDisplayState.balloonsActive = false
  })

  function mount() {
    const app = document.createElement('div')
    app.className = 'app rev-balloon'
    app.dataset.balloons = 'revisions'
    const wrap = document.createElement('div')
    wrap.className = 'page-wrap'
    const host = document.createElement('div')
    wrap.appendChild(host)
    app.appendChild(wrap)
    document.body.appendChild(app)
    const editor = new Editor({
      element: host,
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: null },
            content: [
              { type: 'text', text: 'keep ' },
              {
                type: 'text',
                text: 'gone',
                marks: [{ type: 'del', attrs: { author: 'Bob', date: '2026-07-02T10:00:00Z' } }],
              },
            ],
          },
        ],
      },
    })
    syncMarginAnnotations(wrap, editor.view.dom as HTMLElement, [], 1, [], editor.view, {})
    return { wrap, editor }
  }

  it('balloon width and left margin drive the column; connecting lines can be hidden', () => {
    trackChangesOptionsState.current = withOpts({
      balloonWidth: 300,
      balloonMargin: 'left',
      balloonLines: false,
    })
    const { wrap, editor } = mount()
    expect(wrap.classList.contains('has-markup-area')).toBe(true)
    expect(wrap.classList.contains('markup-left')).toBe(true)
    expect(wrap.classList.contains('markup-no-lines')).toBe(true)
    expect(wrap.style.getPropertyValue('--markup-w')).toBe('300px')
    const bubble = wrap.querySelector<HTMLElement>('.rev-bubble-del')!
    expect(bubble.style.width).toBe('268px')
    expect(parseFloat(bubble.style.left)).toBeLessThan(50)
    editor.destroy()
  })

  it('changed lines: none draws no bar, right border moves it past the paper edge', () => {
    trackChangesOptionsState.current = withOpts({ changedLines: 'none' })
    const a = mount()
    expect(a.wrap.querySelectorAll('.change-bar')).toHaveLength(0)
    a.editor.destroy()
    document.body.innerHTML = ''
    trackChangesOptionsState.current = withOpts({ changedLines: 'right' })
    const b = mount()
    const bar = b.wrap.querySelector<HTMLElement>('.change-bar')!
    expect(parseFloat(bar.style.left)).toBeGreaterThan(150)
    b.editor.destroy()
  })
})
