import { beforeEach, describe, expect, it, vi } from 'vitest'
import { handleGlobalKeydown } from '../src/renderer/keyboard-actions'
import * as clipboardActions from '../src/renderer/clipboard-actions'
import * as slideActions from '../src/renderer/slide-actions'
import * as showActions from '../src/renderer/show-actions'
import * as fileActions from '../src/renderer/file-actions'
import * as styleActions from '../src/renderer/style-actions'
import type { ActionCtx } from '../src/renderer/action-context'

vi.mock('../src/renderer/clipboard-actions', () => ({
  copySelected: vi.fn(),
  cutSelected: vi.fn(),
  copySlides: vi.fn(),
  copyFormat: vi.fn(),
  pasteFormat: vi.fn(),
  pasteClipboard: vi.fn(),
  duplicateSelected: vi.fn(),
  deleteSelected: vi.fn(),
}))
vi.mock('../src/renderer/slide-actions', () => ({
  addSlide: vi.fn(),
  cutSlides: vi.fn(),
  deleteSlides: vi.fn(),
}))
vi.mock('../src/renderer/file-actions', () => ({ flushActiveEdit: vi.fn(async () => {}) }))
vi.mock('../src/renderer/arrange-actions', () => ({}))
vi.mock('../src/renderer/show-actions', () => ({ startSlideShow: vi.fn() }))
vi.mock('../src/renderer/style-actions', () => ({
  onAlign: vi.fn(),
  onFontSizeStep: vi.fn(),
  onTextToggle: vi.fn(),
}))

function makeCtx(over: Record<string, unknown> = {}): ActionCtx {
  return {
    slideShow: false,
    presenter: false,
    editing: null,
    selectedIds: [],
    slide: { nodes: [] },
    slides: [{}],
    current: 0,
    selectedSlides: [0],
    setSelectedSlides: vi.fn(),
    masterItems: null,
    inkTool: 'select',
    viewMode: 'normal',
    brushMode: null,
    enteredGroupId: null,
    overlayOpen: false,
    findNodeCtx: () => null,
    flushNotes: vi.fn(async () => {}),
    ...over,
  } as unknown as ActionCtx
}

function keydown(key: string, init: KeyboardEventInit = {}): KeyboardEvent {
  return new KeyboardEvent('keydown', { key, metaKey: true, cancelable: true, ...init })
}

/** Mounts a focusable thumbnail rail (or sorter grid) and puts keyboard focus on it */
function focusRail(cls: 'slide-list' | 'outline-pane' | 'sorter-view' = 'slide-list'): HTMLElement {
  const rail = document.createElement('div')
  rail.className = cls
  rail.tabIndex = 0
  document.body.appendChild(rail)
  rail.focus()
  return rail
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

describe('AI composer history shortcuts', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
  })

  it.each([
    ['Ctrl+Shift+Z', 'Z', true],
    ['Ctrl+Y', 'y', false],
  ])('routes %s to deck redo while the cleared composer is untouched', (_label, key, shiftKey) => {
    const textarea = document.createElement('textarea')
    textarea.dataset.slidesAiInput = 'true'
    textarea.dataset.deckUndoReady = 'true'
    document.body.appendChild(textarea)
    textarea.focus()
    const redo = vi.fn(async () => {})
    const ctx = makeCtx({ redo })
    textarea.addEventListener('keydown', (event) => handleGlobalKeydown(ctx, event, 'Win32'))
    const event = keydown(key, { metaKey: false, ctrlKey: true, shiftKey })

    textarea.dispatchEvent(event)

    expect(event.defaultPrevented).toBe(true)
    expect(redo).toHaveBeenCalledOnce()
  })

  it.each([
    ['Ctrl+Shift+Z', 'Z', true],
    ['Ctrl+Y', 'y', false],
  ])('preserves native %s after the composer is edited', (_label, key, shiftKey) => {
    const textarea = document.createElement('textarea')
    textarea.dataset.slidesAiInput = 'true'
    textarea.dataset.deckUndoReady = 'false'
    document.body.appendChild(textarea)
    textarea.focus()
    const redo = vi.fn(async () => {})
    const ctx = makeCtx({ redo })
    textarea.addEventListener('keydown', (event) => handleGlobalKeydown(ctx, event, 'Win32'))
    const event = keydown(key, { metaKey: false, ctrlKey: true, shiftKey })

    textarea.dispatchEvent(event)

    expect(event.defaultPrevented).toBe(false)
    expect(redo).not.toHaveBeenCalled()
  })
})

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
    expect(clipboardActions.copySlides).not.toHaveBeenCalled()
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
    focusRail()
    const noSel = keydown('c')
    handleGlobalKeydown(makeCtx(), noSel)
    expect(noSel.defaultPrevented).toBe(true)
    expect(clipboardActions.copySlides).toHaveBeenCalledTimes(1)

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
    focusRail()
    selectText()
    const cut = keydown('x')
    handleGlobalKeydown(makeCtx(), cut)
    expect(cut.defaultPrevented).toBe(false)
    expect(slideActions.cutSlides).not.toHaveBeenCalled()
  })
})

describe('Delete/Backspace on the thumbnail pane', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
    window.getSelection()?.removeAllRanges()
  })

  const plain = (key: string) => new KeyboardEvent('keydown', { key, cancelable: true })

  it('deletes the current slide when the rail has focus and nothing is selected', () => {
    focusRail()
    const ctx = makeCtx({ current: 1, slides: [{}, {}], selectedSlides: [1] })
    const e = plain('Backspace')
    handleGlobalKeydown(ctx, e)
    expect(e.defaultPrevented).toBe(true)
    expect(slideActions.deleteSlides).toHaveBeenCalledWith(ctx, [1])
  })

  it('deletes every selected slide of a multi-selection', () => {
    focusRail()
    const ctx = makeCtx({ current: 1, slides: [{}, {}, {}, {}], selectedSlides: [1, 3] })
    const e = plain('Delete')
    handleGlobalKeydown(ctx, e)
    expect(e.defaultPrevented).toBe(true)
    expect(slideActions.deleteSlides).toHaveBeenCalledWith(ctx, [1, 3])
  })

  it('also deletes in the slide sorter view', () => {
    focusRail('sorter-view')
    const e = plain('Delete')
    handleGlobalKeydown(makeCtx({ viewMode: 'sorter' }), e)
    expect(slideActions.deleteSlides).toHaveBeenCalledTimes(1)
  })

  it('also deletes from the outline pane, but not while typing in a field inside it', () => {
    const pane = focusRail('outline-pane')
    const e = plain('Delete')
    handleGlobalKeydown(makeCtx({ viewMode: 'outline' }), e)
    expect(e.defaultPrevented).toBe(true)
    expect(slideActions.deleteSlides).toHaveBeenCalledTimes(1)

    const input = document.createElement('input')
    pane.appendChild(input)
    input.focus()
    const typed = plain('Delete')
    handleGlobalKeydown(makeCtx({ viewMode: 'outline' }), typed)
    expect(typed.defaultPrevented).toBe(false)
    expect(slideActions.deleteSlides).toHaveBeenCalledTimes(1)
  })

  it('does nothing with canvas focus (body) and nothing selected', () => {
    const del = plain('Delete')
    handleGlobalKeydown(makeCtx({ current: 1, slides: [{}, {}] }), del)
    const back = plain('Backspace')
    handleGlobalKeydown(makeCtx({ current: 1, slides: [{}, {}] }), back)
    expect(del.defaultPrevented).toBe(false)
    expect(back.defaultPrevented).toBe(false)
    expect(slideActions.deleteSlides).not.toHaveBeenCalled()
  })

  it('cuts/copies the slide only from the rail', () => {
    const cutFromCanvas = keydown('x')
    handleGlobalKeydown(makeCtx(), cutFromCanvas)
    const copyFromCanvas = keydown('c')
    handleGlobalKeydown(makeCtx(), copyFromCanvas)
    expect(cutFromCanvas.defaultPrevented).toBe(false)
    expect(copyFromCanvas.defaultPrevented).toBe(false)
    expect(slideActions.cutSlides).not.toHaveBeenCalled()
    expect(clipboardActions.copySlides).not.toHaveBeenCalled()

    focusRail()
    const ctx = makeCtx({ current: 1, slides: [{}, {}, {}], selectedSlides: [1, 2] })
    const cut = keydown('x')
    handleGlobalKeydown(ctx, cut)
    expect(cut.defaultPrevented).toBe(true)
    expect(slideActions.cutSlides).toHaveBeenCalledWith(ctx, [1, 2])
    const copy = keydown('c')
    handleGlobalKeydown(ctx, copy)
    expect(clipboardActions.copySlides).toHaveBeenCalledWith(ctx, [1, 2])
  })

  it('still switches slides with the arrow keys from the canvas and from the rail', () => {
    const setCurrent = vi.fn()
    const setSelectedSlides = vi.fn()
    const ctx = makeCtx({
      current: 1,
      slides: [{}, {}, {}],
      selectedSlides: [0, 1],
      setCurrent,
      setSelectedSlides,
    })
    const down = plain('ArrowDown')
    handleGlobalKeydown(ctx, down)
    expect(down.defaultPrevented).toBe(true)
    expect(setCurrent).toHaveBeenCalledWith(2)
    expect(setSelectedSlides).toHaveBeenCalledWith([2])

    focusRail()
    const up = plain('ArrowUp')
    handleGlobalKeydown(ctx, up)
    expect(up.defaultPrevented).toBe(true)
    expect(setCurrent).toHaveBeenLastCalledWith(0)
    expect(setSelectedSlides).toHaveBeenLastCalledWith([0])
  })

  it('does not switch slides from a ribbon control', () => {
    const btn = document.createElement('button')
    document.body.appendChild(btn)
    btn.focus()
    const setCurrent = vi.fn()
    const e = plain('ArrowDown')
    handleGlobalKeydown(makeCtx({ slides: [{}, {}], setCurrent }), e)
    expect(e.defaultPrevented).toBe(false)
    expect(setCurrent).not.toHaveBeenCalled()
  })

  it('does nothing while an ink tool is active', () => {
    focusRail()
    const e = plain('Delete')
    handleGlobalKeydown(makeCtx({ inkTool: 'pen' }), e)
    expect(e.defaultPrevented).toBe(false)
    expect(slideActions.deleteSlides).not.toHaveBeenCalled()
  })

  it('does nothing in reading view', () => {
    focusRail()
    const e = plain('Backspace')
    handleGlobalKeydown(makeCtx({ viewMode: 'reading' }), e)
    expect(e.defaultPrevented).toBe(false)
    expect(slideActions.deleteSlides).not.toHaveBeenCalled()
  })

  it('leaves the key to a plain-DOM text selection', () => {
    focusRail()
    selectText()
    const e = plain('Delete')
    handleGlobalKeydown(makeCtx(), e)
    expect(e.defaultPrevented).toBe(false)
    expect(slideActions.deleteSlides).not.toHaveBeenCalled()
  })

  it('does not fire from master view or with modifiers', () => {
    focusRail()
    const master = plain('Delete')
    handleGlobalKeydown(makeCtx({ masterItems: [] }), master)
    const alt = new KeyboardEvent('keydown', { key: 'Backspace', altKey: true, cancelable: true })
    handleGlobalKeydown(makeCtx(), alt)
    expect(slideActions.deleteSlides).not.toHaveBeenCalled()
  })
})

describe('Home/End and select-all on the thumbnail pane', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
  })

  const plain = (key: string, init: KeyboardEventInit = {}) =>
    new KeyboardEvent('keydown', { key, cancelable: true, ...init })
  const railCtx = (over: Record<string, unknown> = {}) =>
    makeCtx({
      current: 2,
      slides: [{}, {}, {}, {}, {}],
      selectedSlides: [2],
      setCurrent: vi.fn(),
      setSelectedSlides: vi.fn(),
      setSelectedIds: vi.fn(),
      ...over,
    })

  it('Home / End jump to the first / last slide as a single selection', () => {
    focusRail()
    const ctx = railCtx()
    const home = plain('Home')
    handleGlobalKeydown(ctx, home)
    expect(home.defaultPrevented).toBe(true)
    expect(ctx.setCurrent).toHaveBeenCalledWith(0)
    expect(ctx.setSelectedSlides).toHaveBeenCalledWith([0])

    const end = plain('End')
    handleGlobalKeydown(ctx, end)
    expect(ctx.setCurrent).toHaveBeenLastCalledWith(4)
    expect(ctx.setSelectedSlides).toHaveBeenLastCalledWith([4])
  })

  it('Shift+End / Shift+Home extend the range from the anchor without moving it', () => {
    focusRail('sorter-view')
    const ctx = railCtx({ viewMode: 'sorter' })
    const end = plain('End', { shiftKey: true })
    handleGlobalKeydown(ctx, end)
    expect(end.defaultPrevented).toBe(true)
    expect(ctx.setSelectedSlides).toHaveBeenCalledWith([2, 3, 4])
    handleGlobalKeydown(ctx, plain('Home', { shiftKey: true }))
    expect(ctx.setSelectedSlides).toHaveBeenLastCalledWith([0, 1, 2])
    expect(ctx.setCurrent).not.toHaveBeenCalled()
  })

  it('ignores Home/End from the canvas and from a ribbon control', () => {
    const fromCanvas = plain('End')
    const ctx = railCtx()
    handleGlobalKeydown(ctx, fromCanvas)
    const btn = document.createElement('button')
    document.body.appendChild(btn)
    btn.focus()
    const fromRibbon = plain('Home')
    handleGlobalKeydown(ctx, fromRibbon)
    expect(fromCanvas.defaultPrevented).toBe(false)
    expect(fromRibbon.defaultPrevented).toBe(false)
    expect(ctx.setSelectedSlides).not.toHaveBeenCalled()
  })

  it('Command+A selects every slide with the rail focused, every element otherwise', () => {
    focusRail()
    const ctx = railCtx({ slide: { nodes: [{ sourceId: 'a' }, { sourceId: 'b' }] } })
    const onRail = keydown('a')
    handleGlobalKeydown(ctx, onRail)
    expect(onRail.defaultPrevented).toBe(true)
    expect(ctx.setSelectedSlides).toHaveBeenCalledWith([0, 1, 2, 3, 4])
    expect(ctx.setSelectedIds).not.toHaveBeenCalled()

    document.body.innerHTML = ''
    const onCanvas = keydown('a')
    handleGlobalKeydown(ctx, onCanvas)
    expect(ctx.setSelectedIds).toHaveBeenCalledWith(['a', 'b'])
    expect(ctx.setSelectedSlides).toHaveBeenCalledTimes(1)
  })
})

describe('arrow-key nudge', () => {
  const boxes: Record<string, { x: number; y: number; w: number; h: number; rotationDeg: number }> =
    {
      a: { x: 10, y: 20, w: 100, h: 50, rotationDeg: 0 },
      b: { x: 200, y: 40, w: 80, h: 80, rotationDeg: 15 },
      c: { x: 300, y: 300, w: 30, h: 30, rotationDeg: 0 },
    }
  const api = { editTransformMulti: vi.fn() }
  const arrow = (key: string, init: KeyboardEventInit = {}) =>
    new KeyboardEvent('keydown', { key, cancelable: true, ...init })
  const nudgeCtx = (selectedIds: string[], groupOf: Record<string, string> = {}) =>
    makeCtx({
      selectedIds,
      findNodeCtx: (id: string) =>
        boxes[id] ? { node: { box: boxes[id] }, groupId: groupOf[id] } : null,
      onTransform: vi.fn(async () => {}),
      applySlide: vi.fn(),
    })

  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
    api.editTransformMulti.mockResolvedValue({ nodes: [] })
    ;(window as unknown as { slidesApi: typeof api }).slidesApi = api
  })

  it('sends a multi-selection as one transform batch and applies the slide once', async () => {
    const ctx = nudgeCtx(['a', 'b', 'c'])
    const e = arrow('ArrowRight')
    handleGlobalKeydown(ctx, e)
    expect(e.defaultPrevented).toBe(true)
    await vi.waitFor(() => expect(ctx.applySlide).toHaveBeenCalledTimes(1))
    expect(api.editTransformMulti).toHaveBeenCalledTimes(1)
    const op = api.editTransformMulti.mock.calls[0]![0]
    expect(op.slideIndex).toBe(0)
    expect(op.items).toEqual([
      { sourceId: 'a', xPx: 19.6, yPx: 20, wPx: 100, hPx: 50, rotationDeg: 0 },
      { sourceId: 'b', xPx: 209.6, yPx: 40, wPx: 80, hPx: 80, rotationDeg: 15 },
      { sourceId: 'c', xPx: 309.6, yPx: 300, wPx: 30, hPx: 30, rotationDeg: 0 },
    ])
    expect(ctx.onTransform).not.toHaveBeenCalled()
    expect(ctx.applySlide).toHaveBeenCalledWith(0, { nodes: [] })
  })

  it('batches a Shift resize too and carries the group id of in-group children', async () => {
    const ctx = nudgeCtx(['a', 'b'], { b: 'grp' })
    handleGlobalKeydown(ctx, arrow('ArrowUp', { shiftKey: true }))
    await vi.waitFor(() => expect(api.editTransformMulti).toHaveBeenCalledTimes(1))
    const items = api.editTransformMulti.mock.calls[0]![0].items
    expect(items.map((it: { hPx: number }) => it.hPx)).toEqual([59.6, 89.6])
    expect(items.map((it: { yPx: number }) => it.yPx)).toEqual([15.2, 35.2])
    expect(items[0]).not.toHaveProperty('groupId')
    expect(items[1].groupId).toBe('grp')
  })

  it('keeps a single element on the per-element transform path', async () => {
    const ctx = nudgeCtx(['b'], { b: 'grp' })
    handleGlobalKeydown(ctx, arrow('ArrowDown'))
    await vi.waitFor(() => expect(ctx.onTransform).toHaveBeenCalledTimes(1))
    expect(ctx.onTransform).toHaveBeenCalledWith(
      'b',
      { x: 200, y: 49.6, w: 80, h: 80, rotationDeg: 15 },
      undefined,
      'grp',
    )
    expect(api.editTransformMulti).not.toHaveBeenCalled()
  })

  it('ignores ids that are no longer on the slide', async () => {
    const ctx = nudgeCtx(['a', 'gone', 'c'])
    handleGlobalKeydown(ctx, arrow('ArrowLeft'))
    await vi.waitFor(() => expect(api.editTransformMulti).toHaveBeenCalledTimes(1))
    const items = api.editTransformMulti.mock.calls[0]![0].items
    expect(items.map((it: { sourceId: string }) => it.sourceId)).toEqual(['a', 'c'])
  })
})

describe('new slide shortcut', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
  })

  const flush = () => new Promise((r) => setTimeout(r, 0))

  it('adds a slide once with Shift+Command+N on macOS', async () => {
    const ctx = makeCtx()
    const e = keydown('N', { shiftKey: true })
    handleGlobalKeydown(ctx, e, 'MacIntel')
    await flush()
    expect(e.defaultPrevented).toBe(true)
    expect(slideActions.addSlide).toHaveBeenCalledTimes(1)
    expect(slideActions.addSlide).toHaveBeenCalledWith(ctx)
  })

  it('leaves Command+M to the Minimize menu accelerator on macOS', async () => {
    const e = keydown('m')
    handleGlobalKeydown(makeCtx(), e, 'MacIntel')
    await flush()
    expect(e.defaultPrevented).toBe(false)
    expect(slideActions.addSlide).not.toHaveBeenCalled()
  })

  it('adds a slide with Ctrl+M on Windows', async () => {
    const e = keydown('m', { metaKey: false, ctrlKey: true })
    handleGlobalKeydown(makeCtx(), e, 'Win32')
    await flush()
    expect(e.defaultPrevented).toBe(true)
    expect(slideActions.addSlide).toHaveBeenCalledTimes(1)
  })

  it('also takes Shift+Command+N on macOS but not on Windows', async () => {
    const mac = keydown('N', { shiftKey: true })
    handleGlobalKeydown(makeCtx(), mac, 'MacIntel')
    const win = keydown('N', { metaKey: false, ctrlKey: true, shiftKey: true })
    handleGlobalKeydown(makeCtx(), win, 'Win32')
    await flush()
    expect(mac.defaultPrevented).toBe(true)
    expect(win.defaultPrevented).toBe(false)
    expect(slideActions.addSlide).toHaveBeenCalledTimes(1)
  })

  it('commits the open text edit before adding the slide', async () => {
    const box = document.createElement('div')
    box.contentEditable = 'true'
    document.body.appendChild(box)
    box.focus()
    const ctx = makeCtx({ editing: { sourceId: 't1' } })
    const e = keydown('N', { shiftKey: true })
    handleGlobalKeydown(ctx, e, 'MacIntel')
    await flush()
    expect(e.defaultPrevented).toBe(true)
    expect(fileActions.flushActiveEdit).toHaveBeenCalledWith(ctx)
    expect(slideActions.addSlide).toHaveBeenCalledTimes(1)
  })

  it('flushes the notes draft after the text edit and before the new slide', async () => {
    const notes = document.createElement('textarea')
    document.body.appendChild(notes)
    notes.focus()
    const order: string[] = []
    vi.mocked(fileActions.flushActiveEdit).mockImplementationOnce(async () => {
      order.push('edit')
    })
    const ctx = makeCtx({
      flushNotes: vi.fn(async () => {
        order.push('notes')
      }),
    })
    vi.mocked(slideActions.addSlide).mockImplementationOnce(async () => {
      order.push('slide')
    })
    const e = keydown('N', { shiftKey: true })
    handleGlobalKeydown(ctx, e, 'MacIntel')
    await flush()
    expect(e.defaultPrevented).toBe(true)
    expect(order).toEqual(['edit', 'notes', 'slide'])
  })

  it('is a no-op in slide show, presenter and reading view', async () => {
    for (const over of [{ slideShow: true }, { presenter: true }, { viewMode: 'reading' }]) {
      const e = keydown('N', { shiftKey: true })
      handleGlobalKeydown(makeCtx(over), e, 'MacIntel')
      expect(e.defaultPrevented).toBe(false)
    }
    await flush()
    expect(slideActions.addSlide).not.toHaveBeenCalled()
  })

  it('ignores Shift+M, Alt+M and master view', async () => {
    const ctrl = { metaKey: false, ctrlKey: true }
    handleGlobalKeydown(makeCtx(), keydown('M', { ...ctrl, shiftKey: true }), 'Win32')
    handleGlobalKeydown(makeCtx(), keydown('m', { ...ctrl, altKey: true }), 'Win32')
    handleGlobalKeydown(makeCtx({ masterItems: [] }), keydown('m', ctrl), 'Win32')
    await flush()
    expect(slideActions.addSlide).not.toHaveBeenCalled()
  })
})

describe('arrow-key nudging with a selection', () => {
  const STEP = 9.6

  function makeNudgeCtx(zoom: number, nodes: Array<Record<string, unknown>>) {
    const onTransform = vi.fn()
    const ctx = makeCtx({
      zoom,
      onTransform,
      selectedIds: nodes.map((n) => n.sourceId as string),
      findNodeCtx: (id: string) => {
        const node = nodes.find((n) => n.sourceId === id)
        return node ? { node } : null
      },
    })
    return { ctx, onTransform }
  }

  function square(sourceId = 'a'): Record<string, unknown> {
    return { sourceId, type: 'shape', box: { x: 553.68, y: 200, w: 96, h: 96, rotationDeg: 0 } }
  }

  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
  })

  it('moves by one 0.1 in grid step (9.6 px) on a plain arrow', () => {
    const { ctx, onTransform } = makeNudgeCtx(1, [square()])
    const e = keydown('ArrowRight', { metaKey: false })

    handleGlobalKeydown(ctx, e)

    expect(e.defaultPrevented).toBe(true)
    expect(onTransform).toHaveBeenCalledTimes(1)
    const box = onTransform.mock.calls[0]![1]
    expect(box.x).toBeCloseTo(553.68 + STEP)
    expect(box).toMatchObject({ y: 200, w: 96, h: 96 })
  })

  it('moves by one screen pixel (1 / zoom) with the command/control key', () => {
    const { ctx, onTransform } = makeNudgeCtx(0.8, [square()])
    handleGlobalKeydown(ctx, keydown('ArrowUp', { metaKey: true }))
    expect(onTransform.mock.calls[0]![1].y).toBeCloseTo(200 - 1 / 0.8)

    onTransform.mockClear()
    handleGlobalKeydown(ctx, keydown('ArrowLeft', { ctrlKey: true, metaKey: false }))
    expect(onTransform.mock.calls[0]![1].x).toBeCloseTo(553.68 - 1 / 0.8)
  })

  it('resizes about the center with Shift (72 pt square + Shift+Right = 79.2 pt wide)', () => {
    const { ctx, onTransform } = makeNudgeCtx(1, [square()])

    handleGlobalKeydown(ctx, keydown('ArrowRight', { metaKey: false, shiftKey: true }))

    const box = onTransform.mock.calls[0]![1]
    expect(box.w).toBeCloseTo(105.6)
    expect(box.x).toBeCloseTo(553.68 - 4.8)
    expect(box).toMatchObject({ y: 200, h: 96 })
  })

  it('shrinks with Shift+Left / Shift+Down and grows height with Shift+Up', () => {
    const { ctx, onTransform } = makeNudgeCtx(1, [square()])

    handleGlobalKeydown(ctx, keydown('ArrowLeft', { metaKey: false, shiftKey: true }))
    expect(onTransform.mock.calls[0]![1]).toMatchObject({ x: 553.68 + 4.8, w: 96 - STEP, h: 96 })

    handleGlobalKeydown(ctx, keydown('ArrowUp', { metaKey: false, shiftKey: true }))
    expect(onTransform.mock.calls[1]![1].h).toBeCloseTo(96 + STEP)
    expect(onTransform.mock.calls[1]![1].y).toBeCloseTo(200 - 4.8)

    handleGlobalKeydown(ctx, keydown('ArrowDown', { metaKey: false, shiftKey: true }))
    expect(onTransform.mock.calls[2]![1].h).toBeCloseTo(96 - STEP)
    expect(onTransform.mock.calls[2]![1].y).toBeCloseTo(200 + 4.8)
  })

  it('never shrinks below one step', () => {
    const tiny = { ...square(), box: { x: 100, y: 100, w: 12, h: 12, rotationDeg: 0 } }
    const { ctx, onTransform } = makeNudgeCtx(1, [tiny])

    handleGlobalKeydown(ctx, keydown('ArrowLeft', { metaKey: false, shiftKey: true }))

    const box = onTransform.mock.calls[0]![1]
    expect(box.w).toBeCloseTo(STEP)
    expect(box.x).toBeCloseTo(100 + (12 - STEP) / 2)
  })

  it('leaves the axis the arrow does not change untouched, even when thinner than a step', () => {
    const bar = { ...square(), box: { x: 100, y: 100, w: 96, h: 2, rotationDeg: 0 } }
    const { ctx, onTransform } = makeNudgeCtx(1, [bar])

    handleGlobalKeydown(ctx, keydown('ArrowRight', { metaKey: false, shiftKey: true }))
    expect(onTransform.mock.calls[0]![1]).toMatchObject({ y: 100, h: 2, w: 96 + STEP })

    handleGlobalKeydown(ctx, keydown('ArrowDown', { metaKey: false, shiftKey: true }))
    expect(onTransform.mock.calls[1]![1]).toMatchObject({ x: 100, y: 100, w: 96, h: 2 })

    handleGlobalKeydown(ctx, keydown('ArrowUp', { metaKey: false, shiftKey: true }))
    expect(onTransform.mock.calls[2]![1].h).toBeCloseTo(2 + STEP)
    expect(onTransform.mock.calls[2]![1].y).toBeCloseTo(100 - STEP / 2)
    expect(onTransform.mock.calls[2]![1]).toMatchObject({ x: 100, w: 96 })
  })

  it('resizes every selected element about its own center, but only moves connectors', () => {
    const line = {
      sourceId: 'b',
      type: 'shape',
      line: {},
      box: { x: 0, y: 0, w: 200, h: 0, rotationDeg: 0 },
    }
    const { ctx, onTransform } = makeNudgeCtx(1, [square('a'), line])

    handleGlobalKeydown(ctx, keydown('ArrowRight', { metaKey: false, shiftKey: true }))
    expect(onTransform).toHaveBeenCalledTimes(1)
    expect(onTransform.mock.calls[0]![0]).toBe('a')

    onTransform.mockClear()
    const editTransformMulti = vi.fn((_op: { items: Array<Record<string, unknown>> }) =>
      Promise.resolve(null),
    )
    ;(
      window as unknown as { slidesApi: { editTransformMulti: typeof editTransformMulti } }
    ).slidesApi = { editTransformMulti }
    handleGlobalKeydown(ctx, keydown('ArrowRight', { metaKey: false }))
    expect(onTransform).not.toHaveBeenCalled()
    expect(editTransformMulti).toHaveBeenCalledTimes(1)
    const items = editTransformMulti.mock.calls[0]![0].items
    expect(items).toHaveLength(2)
    expect(items[1]).toMatchObject({ sourceId: 'b', xPx: STEP, wPx: 200 })
  })

  it('leaves the caret alone while editing text', () => {
    const { ctx, onTransform } = makeNudgeCtx(1, [square()])
    ctx.editing = { sourceId: 'a' } as ActionCtx['editing']
    const e = keydown('ArrowRight', { metaKey: false })

    handleGlobalKeydown(ctx, e)

    expect(e.defaultPrevented).toBe(false)
    expect(onTransform).not.toHaveBeenCalled()
  })
})

describe('Escape on the canvas', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
  })

  it('clears a plain selection', () => {
    const setSelectedIds = vi.fn()
    const ctx = makeCtx({ selectedIds: ['a', 'b'], setSelectedIds })
    const e = keydown('Escape', { metaKey: false })

    handleGlobalKeydown(ctx, e)

    expect(e.defaultPrevented).toBe(true)
    expect(setSelectedIds).toHaveBeenCalledWith([])
  })

  it('selects the group when leaving in-group editing instead of clearing', () => {
    const setSelectedIds = vi.fn()
    const setEnteredGroupId = vi.fn()
    const ctx = makeCtx({
      selectedIds: ['child'],
      enteredGroupId: 'grp',
      setSelectedIds,
      setEnteredGroupId,
    })

    handleGlobalKeydown(ctx, keydown('Escape', { metaKey: false }))

    expect(setSelectedIds).toHaveBeenCalledTimes(1)
    expect(setSelectedIds).toHaveBeenCalledWith(['grp'])
    expect(setEnteredGroupId).toHaveBeenCalledWith(null)
  })

  it('is a no-op with an empty selection', () => {
    const setSelectedIds = vi.fn()
    const ctx = makeCtx({ setSelectedIds })
    const e = keydown('Escape', { metaKey: false })

    handleGlobalKeydown(ctx, e)

    expect(e.defaultPrevented).toBe(false)
    expect(setSelectedIds).not.toHaveBeenCalled()
  })

  it('leaves text editing to the overlay', () => {
    const setSelectedIds = vi.fn()
    const ctx = makeCtx({ selectedIds: ['a'], editing: { sourceId: 'a' }, setSelectedIds })
    const e = keydown('Escape', { metaKey: false })

    handleGlobalKeydown(ctx, e)

    expect(e.defaultPrevented).toBe(false)
    expect(setSelectedIds).not.toHaveBeenCalled()
  })

  it.each(['context menu', 'media playback overlay', 'crop', 'color picker'])(
    'keeps the selection while an Escape-owning overlay is open (%s)',
    () => {
      const setSelectedIds = vi.fn()
      const ctx = makeCtx({ selectedIds: ['a'], overlayOpen: true, setSelectedIds })
      const e = keydown('Escape', { metaKey: false })

      handleGlobalKeydown(ctx, e)

      expect(e.defaultPrevented).toBe(false)
      expect(setSelectedIds).not.toHaveBeenCalled()
    },
  )

  it('leaves in-group editing untouched while an overlay is open', () => {
    const setSelectedIds = vi.fn()
    const setEnteredGroupId = vi.fn()
    const ctx = makeCtx({
      selectedIds: ['child'],
      enteredGroupId: 'grp',
      overlayOpen: true,
      setSelectedIds,
      setEnteredGroupId,
    })
    const e = keydown('Escape', { metaKey: false })

    handleGlobalKeydown(ctx, e)

    expect(e.defaultPrevented).toBe(false)
    expect(setSelectedIds).not.toHaveBeenCalled()
    expect(setEnteredGroupId).not.toHaveBeenCalled()
  })

  it('leaves the ink tool and format painter alone while an overlay is open', () => {
    const setInkTool = vi.fn()
    const setBrushMode = vi.fn()
    const ctx = makeCtx({
      inkTool: 'pen',
      brushMode: 'continuous',
      overlayOpen: true,
      setInkTool,
      setBrushMode,
    })

    handleGlobalKeydown(ctx, keydown('Escape', { metaKey: false }))

    expect(setInkTool).not.toHaveBeenCalled()
    expect(setBrushMode).not.toHaveBeenCalled()
  })
})

describe('text formatting shortcuts', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
  })

  /** Mounts the text edit overlay's contentEditable and focuses it */
  function focusEditor(): HTMLElement {
    const box = document.createElement('div')
    box.className = 'slide-text-editor'
    box.tabIndex = 0
    box.contentEditable = 'true'
    Object.defineProperty(box, 'isContentEditable', { value: true })
    document.body.appendChild(box)
    box.focus()
    return box
  }

  const noFormatCalls = () => {
    expect(styleActions.onAlign).not.toHaveBeenCalled()
    expect(styleActions.onFontSizeStep).not.toHaveBeenCalled()
    expect(styleActions.onTextToggle).not.toHaveBeenCalled()
  }

  it.each([
    ['l', 'left'],
    ['e', 'center'],
    ['r', 'right'],
    ['j', 'justify'],
  ] as const)('aligns the selected shapes with Command+%s', (key, align) => {
    const ctx = makeCtx({ selectedIds: ['a', 'b'] })
    const e = keydown(key)
    handleGlobalKeydown(ctx, e)
    expect(e.defaultPrevented).toBe(true)
    expect(styleActions.onAlign).toHaveBeenCalledWith(ctx, align)
  })

  it('aligns the caret paragraph while editing and takes Ctrl on Windows', () => {
    focusEditor()
    const ctx = makeCtx({ editing: { sourceId: 'a' }, selectedIds: [] })
    const e = keydown('e', { metaKey: false, ctrlKey: true })
    handleGlobalKeydown(ctx, e, 'Win32')
    expect(e.defaultPrevented).toBe(true)
    expect(styleActions.onAlign).toHaveBeenCalledWith(ctx, 'center')
  })

  it('also works from a table cell edit', () => {
    focusEditor()
    const ctx = makeCtx({ editingCell: { sourceId: 't', row: 0, col: 0 } })
    handleGlobalKeydown(ctx, keydown('r'))
    expect(styleActions.onAlign).toHaveBeenCalledWith(ctx, 'right')
  })

  it('steps the size by one point with Command+] and Command+[', () => {
    const ctx = makeCtx({ selectedIds: ['a'] })
    handleGlobalKeydown(ctx, keydown(']'))
    handleGlobalKeydown(ctx, keydown('['))
    expect(styleActions.onFontSizeStep).toHaveBeenNthCalledWith(1, ctx, { dir: 1, mode: 'point' })
    expect(styleActions.onFontSizeStep).toHaveBeenNthCalledWith(2, ctx, { dir: -1, mode: 'point' })
  })

  it('grows and shrinks along the ladder with Shift+Command+> and <', () => {
    const ctx = makeCtx({ selectedIds: ['a'] })
    const grow = keydown('>', { shiftKey: true })
    handleGlobalKeydown(ctx, grow)
    const shrink = keydown('<', { shiftKey: true })
    handleGlobalKeydown(ctx, shrink)
    expect(grow.defaultPrevented).toBe(true)
    expect(shrink.defaultPrevented).toBe(true)
    expect(styleActions.onFontSizeStep).toHaveBeenNthCalledWith(1, ctx, { dir: 1, mode: 'ladder' })
    expect(styleActions.onFontSizeStep).toHaveBeenNthCalledWith(2, ctx, {
      dir: -1,
      mode: 'ladder',
    })
  })

  it('accepts the unshifted period/comma keys some layouts report with Shift held', () => {
    const ctx = makeCtx({ selectedIds: ['a'] })
    handleGlobalKeydown(ctx, keydown('.', { shiftKey: true }))
    handleGlobalKeydown(ctx, keydown(',', { shiftKey: true }))
    expect(styleActions.onFontSizeStep).toHaveBeenNthCalledWith(1, ctx, { dir: 1, mode: 'ladder' })
    expect(styleActions.onFontSizeStep).toHaveBeenNthCalledWith(2, ctx, {
      dir: -1,
      mode: 'ladder',
    })
  })

  it('does not treat a plain period or comma as a size step', () => {
    handleGlobalKeydown(makeCtx({ selectedIds: ['a'] }), keydown('.'))
    handleGlobalKeydown(makeCtx({ selectedIds: ['a'] }), keydown(','))
    expect(styleActions.onFontSizeStep).not.toHaveBeenCalled()
  })

  it('resizes the selection while editing', () => {
    focusEditor()
    const ctx = makeCtx({ editing: { sourceId: 'a' } })
    handleGlobalKeydown(ctx, keydown(']'))
    expect(styleActions.onFontSizeStep).toHaveBeenCalledWith(ctx, { dir: 1, mode: 'point' })
  })

  it.each([
    ['b', 'bold'],
    ['i', 'italic'],
    ['u', 'underline'],
  ] as const)('toggles %s on selected shapes with Command+%s', (key, kind) => {
    const ctx = makeCtx({ selectedIds: ['a'] })
    const e = keydown(key)
    handleGlobalKeydown(ctx, e)
    expect(e.defaultPrevented).toBe(true)
    expect(styleActions.onTextToggle).toHaveBeenCalledWith(ctx, kind)
  })

  it('leaves Command+B/I/U to the browser while editing', () => {
    focusEditor()
    const e = keydown('b')
    handleGlobalKeydown(makeCtx({ editing: { sourceId: 'a' } }), e)
    expect(e.defaultPrevented).toBe(false)
    expect(styleActions.onTextToggle).not.toHaveBeenCalled()
  })

  it('stays out of ordinary inputs, even with shapes selected', () => {
    const input = document.createElement('input')
    document.body.appendChild(input)
    input.focus()
    const ctx = makeCtx({ selectedIds: ['a'] })
    for (const e of [keydown('l'), keydown(']'), keydown('>', { shiftKey: true }), keydown('b')]) {
      handleGlobalKeydown(ctx, e)
      expect(e.defaultPrevented).toBe(false)
    }
    noFormatCalls()
  })

  it('does nothing while editing with focus in a keep-edit ribbon field', () => {
    const input = document.createElement('input')
    document.body.appendChild(input)
    input.focus()
    const e = keydown('e')
    handleGlobalKeydown(makeCtx({ editing: { sourceId: 'a' }, selectedIds: ['a'] }), e)
    expect(e.defaultPrevented).toBe(false)
    noFormatCalls()
  })

  it('does nothing with nothing selected, with Shift, or with Alt', () => {
    handleGlobalKeydown(makeCtx(), keydown('l'))
    handleGlobalKeydown(makeCtx({ selectedIds: ['a'] }), keydown('L', { shiftKey: true }))
    handleGlobalKeydown(makeCtx({ selectedIds: ['a'] }), keydown('B', { shiftKey: true }))
    handleGlobalKeydown(makeCtx({ selectedIds: ['a'] }), keydown(']', { altKey: true }))
    noFormatCalls()
  })

  it('is a no-op in slide show, presenter, reading and master views', () => {
    for (const over of [
      { slideShow: { startAt: 0 } },
      { presenter: { startAt: 0 } },
      { viewMode: 'reading' },
      { masterItems: [] },
    ]) {
      for (const e of [
        keydown('l'),
        keydown(']'),
        keydown('>', { shiftKey: true }),
        keydown('b'),
      ]) {
        handleGlobalKeydown(makeCtx({ selectedIds: ['a'], ...over }), e)
        expect(e.defaultPrevented).toBe(false)
      }
    }
    noFormatCalls()
  })
})

describe('Enter / F2 on a selected text shape', () => {
  const nodes: Record<string, Record<string, unknown>> = {
    txt: { sourceId: 'txt', type: 'shape' },
    pic: { sourceId: 'pic', type: 'picture' },
    line: { sourceId: 'line', type: 'shape', line: {} },
  }
  const editCtx = (selectedIds: string[], over: Record<string, unknown> = {}) =>
    makeCtx({
      selectedIds,
      setEditing: vi.fn(),
      findNodeCtx: (id: string) => (nodes[id] ? { node: nodes[id], groupId: over.groupId } : null),
      ...over,
    })
  const key = (k: string, init: KeyboardEventInit = {}) =>
    new KeyboardEvent('keydown', { key: k, cancelable: true, ...init })

  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
  })

  it('Enter enters editing with the caret after the text', () => {
    const ctx = editCtx(['txt'])
    const e = key('Enter')
    handleGlobalKeydown(ctx, e)
    expect(e.defaultPrevented).toBe(true)
    expect(ctx.setEditing).toHaveBeenCalledWith({ sourceId: 'txt' })
  })

  it('F2 enters editing with the whole text selected', () => {
    const ctx = editCtx(['txt'], { groupId: 'g1' })
    const e = key('F2')
    handleGlobalKeydown(ctx, e)
    expect(e.defaultPrevented).toBe(true)
    expect(ctx.setEditing).toHaveBeenCalledWith({ sourceId: 'txt', selectAll: true, groupId: 'g1' })
  })

  it('a printable key still replaces the whole text', () => {
    const ctx = editCtx(['txt'])
    handleGlobalKeydown(ctx, key('a'))
    expect(ctx.setEditing).toHaveBeenCalledWith({ sourceId: 'txt', replaceWith: 'a' })
  })

  it('does nothing on a picture or a connector', () => {
    for (const id of ['pic', 'line']) {
      const ctx = editCtx([id])
      const e = key('Enter')
      handleGlobalKeydown(ctx, e)
      handleGlobalKeydown(ctx, key('F2'))
      expect(e.defaultPrevented).toBe(false)
      expect(ctx.setEditing).not.toHaveBeenCalled()
    }
  })

  it('ignores Shift+Enter, Alt+Enter and an already consumed Enter', () => {
    const ctx = editCtx(['txt'])
    handleGlobalKeydown(ctx, key('Enter', { shiftKey: true }))
    handleGlobalKeydown(ctx, key('Enter', { altKey: true }))
    const consumed = key('Enter')
    consumed.preventDefault()
    handleGlobalKeydown(ctx, consumed)
    expect(ctx.setEditing).not.toHaveBeenCalled()
  })

  it('does nothing while editing, in a text field, with several shapes, or in master/reading view', () => {
    for (const over of [
      { editing: { sourceId: 'txt' } },
      { masterItems: [] },
      { viewMode: 'reading' },
    ]) {
      const ctx = editCtx(['txt'], over)
      handleGlobalKeydown(ctx, key('Enter'))
      handleGlobalKeydown(ctx, key('F2'))
      expect(ctx.setEditing).not.toHaveBeenCalled()
    }
    const multi = editCtx(['txt', 'pic'])
    handleGlobalKeydown(multi, key('Enter'))
    expect(multi.setEditing).not.toHaveBeenCalled()

    const input = document.createElement('input')
    document.body.appendChild(input)
    input.focus()
    const field = editCtx(['txt'])
    handleGlobalKeydown(field, key('Enter'))
    handleGlobalKeydown(field, key('F2'))
    expect(field.setEditing).not.toHaveBeenCalled()
  })
})
