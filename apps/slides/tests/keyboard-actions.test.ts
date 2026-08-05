import { beforeEach, describe, expect, it, vi } from 'vitest'
import { handleGlobalKeydown } from '../src/renderer/keyboard-actions'
import * as clipboardActions from '../src/renderer/clipboard-actions'
import * as slideActions from '../src/renderer/slide-actions'
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
vi.mock('../src/renderer/slide-actions', () => ({ cutSlideAt: vi.fn() }))
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
