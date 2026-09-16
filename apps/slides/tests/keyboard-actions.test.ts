import { beforeEach, describe, expect, it, vi } from 'vitest'
import { handleGlobalKeydown } from '../src/renderer/keyboard-actions'
import * as clipboardActions from '../src/renderer/clipboard-actions'
import * as slideActions from '../src/renderer/slide-actions'
import * as showActions from '../src/renderer/show-actions'
import type { ActionCtx } from '../src/renderer/action-context'

vi.mock('../src/renderer/clipboard-actions', () => ({
  copySelected: vi.fn(),
  cutSelected: vi.fn(),
  copySlideAt: vi.fn(),
  copyFormat: vi.fn(),
  pasteFormat: vi.fn(),
  pasteClipboard: vi.fn(),
  duplicateSelected: vi.fn(),
  deleteSelected: vi.fn(),
}))
vi.mock('../src/renderer/slide-actions', () => ({ cutSlideAt: vi.fn(), deleteSlideAt: vi.fn() }))
vi.mock('../src/renderer/arrange-actions', () => ({}))
vi.mock('../src/renderer/show-actions', () => ({ startSlideShow: vi.fn() }))

function makeCtx(over: Record<string, unknown> = {}): ActionCtx {
  return {
    slideShow: false,
    presenter: false,
    editing: null,
    selectedIds: [],
    slide: { nodes: [] },
    slides: [{}],
    current: 0,
    masterItems: null,
    inkTool: 'select',
    viewMode: 'normal',
    brushMode: null,
    enteredGroupId: null,
    findNodeCtx: () => null,
    ...over,
  } as unknown as ActionCtx
}

function keydown(key: string, init: KeyboardEventInit = {}): KeyboardEvent {
  return new KeyboardEvent('keydown', { key, metaKey: true, cancelable: true, ...init })
}

function selectText(): void {
  const div = document.createElement('div')
  div.textContent = 'answer from the AI panel'
  document.body.appendChild(div)
  const range = document.createRange()
  range.selectNodeContents(div)
  const sel = window.getSelection()!
  sel.removeAllRanges()
  sel.addRange(range)
}

describe('slide show shortcuts', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
  })

  it('starts from the current slide with Command+Enter on macOS', () => {
    const ctx = makeCtx({ current: 2, slides: [{}, {}, {}] })
    const e = keydown('Enter')

    handleGlobalKeydown(ctx, e, 'MacIntel')

    expect(e.defaultPrevented).toBe(true)
    expect(showActions.startSlideShow).toHaveBeenCalledWith(ctx, false)
  })

  it('starts from the current slide with Shift+F5 on Windows', () => {
    const ctx = makeCtx({ current: 2, slides: [{}, {}, {}] })
    const e = keydown('F5', { metaKey: false, shiftKey: true })

    handleGlobalKeydown(ctx, e, 'Win32')

    expect(e.defaultPrevented).toBe(true)
    expect(showActions.startSlideShow).toHaveBeenCalledWith(ctx, false)
  })

  it('does not repurpose Ctrl+Enter on Windows', () => {
    const ctx = makeCtx()
    const e = keydown('Enter', { metaKey: false, ctrlKey: true })

    handleGlobalKeydown(ctx, e, 'Win32')

    expect(e.defaultPrevented).toBe(false)
    expect(showActions.startSlideShow).not.toHaveBeenCalled()
  })

  it('does not start a show from a text field on macOS', () => {
    const input = document.createElement('textarea')
    document.body.appendChild(input)
    input.focus()
    const ctx = makeCtx()
    const e = keydown('Enter')

    handleGlobalKeydown(ctx, e, 'MacIntel')

    expect(e.defaultPrevented).toBe(false)
    expect(showActions.startSlideShow).not.toHaveBeenCalled()
  })

  it('does not start a show when another handler consumed Enter', () => {
    const ctx = makeCtx()
    const e = keydown('Enter')
    e.preventDefault()

    handleGlobalKeydown(ctx, e, 'MacIntel')

    expect(showActions.startSlideShow).not.toHaveBeenCalled()
  })

  it('does not start a show while confirming a crop', () => {
    const ctx = makeCtx({ cropTarget: {} })
    const e = keydown('Enter')

    handleGlobalKeydown(ctx, e, 'MacIntel')

    expect(e.defaultPrevented).toBe(false)
    expect(showActions.startSlideShow).not.toHaveBeenCalled()
  })
})

describe('copy shortcuts with a DOM text selection', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
    window.getSelection()?.removeAllRanges()
  })

  it('does not hijack ⌘C for the slide clipboard when nothing is selected on canvas', () => {
    selectText()
    const e = keydown('c')
    handleGlobalKeydown(makeCtx(), e)
    expect(e.defaultPrevented).toBe(false)
    expect(clipboardActions.copySlideAt).not.toHaveBeenCalled()
  })

  it('does not hijack ⌘C/⌘X for the element clipboard when shapes are selected', () => {
    selectText()
    const ctx = makeCtx({ selectedIds: ['s1'] })
    const c = keydown('c')
    handleGlobalKeydown(ctx, c)
    const x = keydown('x')
    handleGlobalKeydown(ctx, x)
    expect(c.defaultPrevented).toBe(false)
    expect(x.defaultPrevented).toBe(false)
    expect(clipboardActions.copySelected).not.toHaveBeenCalled()
    expect(clipboardActions.cutSelected).not.toHaveBeenCalled()
  })

  it('still copies the current slide / selection when the selection is collapsed', () => {
    const noSel = keydown('c')
    handleGlobalKeydown(makeCtx(), noSel)
    expect(noSel.defaultPrevented).toBe(true)
    expect(clipboardActions.copySlideAt).toHaveBeenCalledTimes(1)

    const withSel = keydown('x')
    handleGlobalKeydown(makeCtx({ selectedIds: ['s1'] }), withSel)
    expect(withSel.defaultPrevented).toBe(true)
    expect(clipboardActions.cutSelected).toHaveBeenCalledTimes(1)
  })

  it('keeps non-copy shortcuts unaffected by a text selection', () => {
    selectText()
    const del = new KeyboardEvent('keydown', { key: 'Delete', cancelable: true })
    handleGlobalKeydown(makeCtx({ selectedIds: ['s1'] }), del)
    expect(del.defaultPrevented).toBe(true)
    expect(clipboardActions.deleteSelected).toHaveBeenCalledTimes(1)
  })

  it('does not hijack ⌘X for slide cut while text is selected', () => {
    selectText()
    const cut = keydown('x')
    handleGlobalKeydown(makeCtx(), cut)
    expect(cut.defaultPrevented).toBe(false)
    expect(slideActions.cutSlideAt).not.toHaveBeenCalled()
  })
})

describe('Delete/Backspace on the thumbnail pane', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
    window.getSelection()?.removeAllRanges()
  })

  const plain = (key: string) => new KeyboardEvent('keydown', { key, cancelable: true })

  it('deletes the current slide when nothing is selected on canvas', () => {
    const ctx = makeCtx({ current: 1, slides: [{}, {}] })
    const e = plain('Backspace')
    handleGlobalKeydown(ctx, e)
    expect(e.defaultPrevented).toBe(true)
    expect(slideActions.deleteSlideAt).toHaveBeenCalledWith(ctx, 1)
  })

  it('also deletes in the slide sorter view', () => {
    const e = plain('Delete')
    handleGlobalKeydown(makeCtx({ viewMode: 'sorter' }), e)
    expect(slideActions.deleteSlideAt).toHaveBeenCalledTimes(1)
  })

  it('does nothing while an ink tool is active', () => {
    const e = plain('Delete')
    handleGlobalKeydown(makeCtx({ inkTool: 'pen' }), e)
    expect(e.defaultPrevented).toBe(false)
    expect(slideActions.deleteSlideAt).not.toHaveBeenCalled()
  })

  it('does nothing in reading view', () => {
    const e = plain('Backspace')
    handleGlobalKeydown(makeCtx({ viewMode: 'reading' }), e)
    expect(e.defaultPrevented).toBe(false)
    expect(slideActions.deleteSlideAt).not.toHaveBeenCalled()
  })

  it('leaves the key to a plain-DOM text selection', () => {
    selectText()
    const e = plain('Delete')
    handleGlobalKeydown(makeCtx(), e)
    expect(e.defaultPrevented).toBe(false)
    expect(slideActions.deleteSlideAt).not.toHaveBeenCalled()
  })

  it('does not fire from master view or with modifiers', () => {
    const master = plain('Delete')
    handleGlobalKeydown(makeCtx({ masterItems: [] }), master)
    const alt = new KeyboardEvent('keydown', { key: 'Backspace', altKey: true, cancelable: true })
    handleGlobalKeydown(makeCtx(), alt)
    expect(slideActions.deleteSlideAt).not.toHaveBeenCalled()
  })
})
