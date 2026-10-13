import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { HeuristicMetrics, layoutText, makeViewport } from '@genoffice/pptx-render'
import type { Paragraph } from '@genoffice/pptx-engine'
import type { RenderNode, RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'
import { applyEditParagraphs } from '@genoffice/pptx-ops'
import { TextEditOverlay } from '../src/renderer/TextEditOverlay'
import { buildCtxItems } from '../src/renderer/context-menu-items'
import { RedactDialog } from '../src/renderer/components/RedactDialog'
import * as redactActions from '../src/renderer/redact-actions'
import { t } from '../src/renderer/i18n/locale'
import type { CtxItem } from '../src/renderer/components/ContextMenu'
import type { ActionCtx, CtxMenuState, RedactDialogState } from '../src/renderer/action-context'
import type { EditParagraph } from '../src/shared/ipc'

// The menu builder reaches konva through these three (konva's node build wants the
// native canvas package). Nothing here draws, so only the eager predicate is stubbed;
// the item callbacks that would draw or move things stay inert.
vi.mock('../src/renderer/picture-edit-actions', () => ({
  canSaveAsPicture: () => true,
  saveSelectionAsPicture: vi.fn(),
  replacePicture: vi.fn(),
  startCrop: vi.fn(),
  startCutout: vi.fn(),
}))
vi.mock('../src/renderer/clipboard-actions', () => ({}))
vi.mock('../src/renderer/slide-actions', () => ({}))

/**
 * The reader's surface for withholding content from the model.
 *
 * The centre of gravity is the text hand-off. Text is not marked by an op — a text
 * element has no runs for one to address — it is marked in the editor DOM and left
 * to the overlay's ordinary commit. So the property that matters is not "the dialog
 * rendered": it is that the payload the existing commit handler sends really carries
 * the label *and* the underline, and that a model rebuilt from that payload hides the
 * words from a model while the reader keeps them.
 */

const SECRET = '13800138000'
const LABEL = 'client phone'
const vp = makeViewport({ cx: 12192000, cy: 6858000 }, 1280) // scale = 1
const NO_INSETS = { l: 0, t: 0, r: 0, b: 0 }
const roots: Array<{ root: Root; container: HTMLElement }> = []

function mount(element: React.ReactElement): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  roots.push({ root, container })
  return container
}

/** A text box whose text really is laid out, so the editor DOM has run containers in it. */
function textNode(paragraphs: Paragraph[]): ShapeRenderNode {
  const { lines } = layoutText({
    body: { paragraphs, insets: NO_INSETS },
    boxWidthPx: 900,
    boxHeightPx: 400,
    metrics: new HeuristicMetrics(),
    vp,
  })
  return {
    type: 'text',
    id: 'r1',
    sourceId: 's1',
    box: { x: 0, y: 0, w: 600, h: 60 },
    text: { lines },
  } as unknown as ShapeRenderNode
}

const BODY = [{ runs: [{ text: `Call ${SECRET} now` }] }]

/** The mark on a run of its own, so a neighbouring run can be checked for staying unmarked. */
const MARKED_BODY = [
  { runs: [{ text: 'Call ' }, { text: SECRET, redact: LABEL, underline: true }, { text: ' now' }] },
]

/**
 * jsdom does not derive isContentEditable from the attribute, and will not focus a
 * div that is not focusable — both are what the selection hand-off reads. The
 * overlay is mounted for real, so the two properties are pinned onto the element
 * React rendered (the same accommodation edit-selection.test.ts makes).
 */
function makeSelectable(editor: HTMLElement): HTMLElement {
  editor.tabIndex = 0
  Object.defineProperty(editor, 'isContentEditable', { value: true })
  return editor
}

/** Select the characters the reader dragged over, the way a drag would. */
function selectText(editor: HTMLElement, needle: string): void {
  editor.focus()
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT)
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const text = n as Text
    const at = text.data.indexOf(needle)
    if (at === -1) continue
    const range = document.createRange()
    range.setStart(text, at)
    range.setEnd(text, at + needle.length)
    const sel = window.getSelection()!
    sel.removeAllRanges()
    sel.addRange(range)
    return
  }
  throw new Error(`no editor text node holds ${needle}`)
}

/**
 * An edit the reader made somewhere in the box. jsdom does not run contenteditable
 * input, and the extraction reads the DOM, so the text node is rewritten directly —
 * which is exactly the state a keystroke would leave behind.
 */
function typeInto(editor: HTMLElement, needle: string, replacement: string): void {
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT)
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const text = n as Text
    const at = text.data.indexOf(needle)
    if (at === -1) continue
    text.data = text.data.slice(0, at) + replacement + text.data.slice(at + needle.length)
    return
  }
  throw new Error(`no editor text node holds ${needle}`)
}

/** The overlay's own right-click handler: it saves the range before the menu takes the click. */
function rightClick(editor: HTMLElement): void {
  act(() => {
    editor.dispatchEvent(
      new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 20, clientY: 20 }),
    )
  })
}

afterEach(() => {
  for (const { root, container } of roots.splice(0)) {
    act(() => root.unmount())
    container.remove()
  }
  window.getSelection()?.removeAllRanges()
})

/** The dialog state and setters `redactActions` reads and writes. */
function actionCtx(over: Partial<ActionCtx> = {}): ActionCtx {
  const state: { redactDialog: ActionCtx['redactDialog'] } = { redactDialog: null }
  return {
    editing: { sourceId: 's1' },
    setStatus: vi.fn(),
    get redactDialog() {
      return state.redactDialog
    },
    // the action always sets a value here, never an updater
    setRedactDialog: (next: RedactDialogState | null) => {
      state.redactDialog = next
    },
    ...over,
  } as unknown as ActionCtx
}

describe('withholding a text selection: the DOM → EditRun.redact hand-off', () => {
  /** The real gesture, end to end: right-click → dialog → confirm → the overlay's own commit. */
  function markSelection(label = LABEL): {
    onCommit: ReturnType<typeof vi.fn>
    ctx: ActionCtx
    editor: HTMLElement
  } {
    const onCommit = vi.fn()
    const onCancel = vi.fn()
    const container = mount(
      createElement(TextEditOverlay, {
        node: textNode(BODY),
        scale: 1,
        onCommit,
        onCancel,
        onContextMenu: vi.fn(),
      }),
    )
    const editor = makeSelectable(container.querySelector('.slide-text-editor') as HTMLElement)
    selectText(editor, SECRET)

    const ctx = actionCtx()
    // the overlay's own right-click handler saves the range before the menu opens
    act(() => {
      editor.dispatchEvent(
        new MouseEvent('contextmenu', {
          bubbles: true,
          cancelable: true,
          clientX: 20,
          clientY: 20,
        }),
      )
    })
    act(() => redactActions.openTextRedaction(ctx))
    act(() => void redactActions.applyRedaction(ctx, label))
    // leaving the editor is what has always committed an edit
    act(() => {
      editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })
    return { onCommit, ctx, editor }
  }

  it('commits the label and the underline through the existing commit handler', () => {
    const { onCommit } = markSelection()
    expect(onCommit).toHaveBeenCalledTimes(1)
    const payload = onCommit.mock.calls[0]![0] as EditParagraph[]
    const runs = payload.flatMap((p) => p.runs)
    const marked = runs.filter((r) => r.redact != null)

    expect(marked).toHaveLength(1)
    expect(marked[0]!.redact).toBe(LABEL)
    // not cosmetic: the engine rewrites u from the model, so a mark without an
    // underline reopens with nothing showing
    expect(marked[0]!.underline).toBe(true)
    // and the reader's own words are still there — a mark, never a deletion
    expect(runs.map((r) => r.text).join('')).toBe(`Call ${SECRET} now`)
  })

  it('rebuilds a model that hides the words from a model and keeps them for the reader', () => {
    const { onCommit } = markSelection()
    const payload = onCommit.mock.calls[0]![0] as EditParagraph[]
    const old: Paragraph[] = [{ runs: [{ text: `Call ${SECRET} now` }] }]

    const rebuilt = applyEditParagraphs(old, payload)
    const held = rebuilt[0]!.runs.find((r) => r.text === SECRET)
    expect(held?.redact).toBe(LABEL)
    expect(held?.underline).toBe(true)
    expect(rebuilt[0]!.runs.map((r) => r.text).join('')).toBe(`Call ${SECRET} now`)
  })

  it('offers to stop withholding an already-marked selection, on the same gesture', () => {
    const onCommit = vi.fn()
    const container = mount(
      createElement(TextEditOverlay, {
        node: textNode(MARKED_BODY),
        scale: 1,
        onCommit,
        onCancel: vi.fn(),
        onContextMenu: vi.fn(),
      }),
    )
    const editor = makeSelectable(container.querySelector('.slide-text-editor') as HTMLElement)
    const ctx = actionCtx()
    selectText(editor, SECRET)
    rightClick(editor)
    // the mark came back with the run, so the same gesture now clears it
    act(() => redactActions.openTextRedaction(ctx))
    // an edit elsewhere in the box, so the session is a change and commits
    typeInto(editor, 'now', 'today')
    act(() => {
      editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })

    const payload = onCommit.mock.calls.at(-1)![0] as EditParagraph[]
    const runs = payload.flatMap((p) => p.runs)
    expect(runs.map((r) => r.text).join('')).toBe(`Call ${SECRET} today`)
    expect(runs.every((r) => r.redact == null)).toBe(true)

    // The payload alone does not prove it: a run that omits `redact` entirely reads as
    // "keep the model's mark" downstream, so the *model* is what has to lose it.
    const rebuilt = applyEditParagraphs(
      [
        {
          runs: [
            { text: 'Call ' },
            { text: SECRET, redact: LABEL, underline: true },
            { text: ' today' },
          ],
        },
      ],
      payload,
    )
    expect(rebuilt[0]!.runs.every((r) => r.redact == null)).toBe(true)
    expect(rebuilt[0]!.runs.map((r) => r.text).join('')).toBe(`Call ${SECRET} today`)
  })

  it('a marked run survives being reopened and edited: the model keeps the mark', () => {
    // The engine rewrites u from the model, and the editor DOM is authoritative on
    // commit, so an editor that did not restore the mark would drop it the moment
    // the reader typed anything else in the box.
    const onCommit = vi.fn()
    const container = mount(
      createElement(TextEditOverlay, {
        node: textNode(MARKED_BODY),
        scale: 1,
        onCommit,
        onCancel: vi.fn(),
        onContextMenu: vi.fn(),
      }),
    )
    const editor = makeSelectable(container.querySelector('.slide-text-editor') as HTMLElement)
    // the reader types at the end of the box, nowhere near the withheld span
    typeInto(editor, 'now', 'today')
    act(() => {
      editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })

    const runs = (onCommit.mock.calls[0]![0] as EditParagraph[]).flatMap((p) => p.runs)
    const held = runs.find((r) => r.text === SECRET)
    expect(held?.redact).toBe(LABEL)
    expect(held?.underline).toBe(true)
    expect(runs.map((r) => r.text).join('')).toBe(`Call ${SECRET} today`)
    // the words either side of it were never part of the mark and stay readable
    expect(runs.filter((r) => r.text === 'Call ').every((r) => r.redact == null)).toBe(true)
  })

  it('marks only the selected part of a run the model already underlines', () => {
    // Nothing but the mark separates the two halves here: same source run, same
    // underline, same colour. If the mark were not part of what makes two runs
    // different, the halves would merge and the words would quietly come back.
    const onCommit = vi.fn()
    const ctx = actionCtx()
    const container = mount(
      createElement(TextEditOverlay, {
        node: textNode([{ runs: [{ text: `Call ${SECRET} now`, underline: true }] }]),
        scale: 1,
        onCommit,
        onCancel: vi.fn(),
        onContextMenu: vi.fn(),
      }),
    )
    const editor = makeSelectable(container.querySelector('.slide-text-editor') as HTMLElement)
    selectText(editor, SECRET)
    rightClick(editor)
    act(() => redactActions.openTextRedaction(ctx))
    act(() => void redactActions.applyRedaction(ctx, LABEL))
    act(() => {
      editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })

    const runs = (onCommit.mock.calls[0]![0] as EditParagraph[]).flatMap((p) => p.runs)
    const held = runs.find((r) => r.redact != null)
    expect(held?.text).toBe(SECRET)
    expect(held?.underline).toBe(true)
    // the rest of the run is still the model's to read
    expect(
      runs
        .filter((r) => r.redact == null)
        .map((r) => r.text)
        .join(''),
    ).toBe('Call  now')
  })

  it('leaves a run the reader never marked unmarked, so the model keeps seeing it', () => {
    const onCommit = vi.fn()
    const container = mount(
      createElement(TextEditOverlay, {
        node: textNode(BODY),
        scale: 1,
        onCommit,
        onCancel: vi.fn(),
        onContextMenu: vi.fn(),
      }),
    )
    const editor = makeSelectable(container.querySelector('.slide-text-editor') as HTMLElement)
    typeInto(editor, 'now', 'today')
    act(() => {
      editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })
    const runs = (onCommit.mock.calls[0]![0] as EditParagraph[]).flatMap((p) => p.runs)
    expect(runs.map((r) => r.text).join('')).toBe(`Call ${SECRET} today`)
    expect(runs.every((r) => r.redact == null)).toBe(true)
  })
})

describe('the two context-menu items', () => {
  const picture = (redact?: string): RenderNode =>
    ({
      type: 'picture',
      id: 'r2',
      sourceId: 'p1',
      box: { x: 0, y: 0, w: 200, h: 150 },
      ...(redact ? { redact } : {}),
    }) as unknown as RenderNode

  const slideOf = (nodes: RenderNode[]): RenderSlide =>
    ({ nodes, index: 0 }) as unknown as RenderSlide

  function ctxFor(ctxMenu: CtxMenuState, nodes: RenderNode[], selectedIds: string[]): ActionCtx {
    return {
      ctxMenu,
      slides: [slideOf(nodes)],
      slide: slideOf(nodes),
      current: 0,
      selectedIds,
      sections: [],
      hasClipboard: false,
      canPasteSlide: false,
      setRedactDialog: vi.fn(),
    } as unknown as ActionCtx
  }

  const itemOf = (items: Array<CtxItem | null>, label: string) =>
    items.find((i) => i?.label === label)

  it('offers the command over a text selection and disables it on a bare caret', () => {
    const nodes = [textNode(BODY)]
    const over = itemOf(
      buildCtxItems(ctxFor({ kind: 'text', x: 0, y: 0, collapsed: false }, nodes, [])),
      t('redactMenuLabel'),
    )
    const caret = itemOf(
      buildCtxItems(ctxFor({ kind: 'text', x: 0, y: 0, collapsed: true }, nodes, [])),
      t('redactMenuLabel'),
    )
    expect(over).toBeDefined()
    // disabled, not hidden: the menu must not reflow as the selection collapses
    expect(over?.disabled).toBeFalsy()
    expect(caret?.disabled).toBe(true)
  })

  it('offers it on a picture, disabled when the op could only mark one of several', () => {
    const nodes = [picture(), picture()]
    const single = itemOf(
      buildCtxItems(ctxFor({ kind: 'element', x: 0, y: 0, targetId: 'p1' }, nodes, ['p1'])),
      t('redactMenuLabel'),
    )
    const many = itemOf(
      buildCtxItems(ctxFor({ kind: 'element', x: 0, y: 0, targetId: 'p1' }, nodes, ['p1', 'p2'])),
      t('redactMenuLabel'),
    )
    expect(single?.disabled).toBeFalsy()
    expect(many?.disabled).toBe(true)
  })

  it('leaves a shape with no picture in it alone', () => {
    const text = { type: 'text', id: 'r3', sourceId: 'sh1', box: { x: 0, y: 0, w: 10, h: 10 } }
    const items = buildCtxItems(
      ctxFor(
        { kind: 'element', x: 0, y: 0, targetId: 'sh1' },
        [text as unknown as RenderNode],
        ['sh1'],
      ),
    )
    expect(itemOf(items, t('redactMenuLabel'))).toBeUndefined()
  })
})

describe('the dialog', () => {
  const type = (input: HTMLInputElement, value: string) =>
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(
        window.HTMLInputElement.prototype,
        'value',
      )!.set!
      setter.call(input, value)
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })

  const openDialog = () => {
    const onSubmit = vi.fn()
    const onCancel = vi.fn()
    mount(createElement(RedactDialog, { seed: '', onSubmit, onCancel }))
    return {
      onSubmit,
      onCancel,
      input: document.querySelector('.modal input') as HTMLInputElement,
      preview: document.querySelector('.redact-preview') as HTMLElement,
      // not named `confirm`: jsdom defines a global confirm() that would shadow it
      confirmBtn: document.querySelector('.modal-actions button.primary') as HTMLButtonElement,
    }
  }

  it('shows the marker it will commit as the label is typed', () => {
    const { input, preview, onSubmit, confirmBtn } = openDialog()
    type(input, LABEL)
    expect(preview.textContent).toBe(`{{${LABEL}}}`)
    act(() => confirmBtn.click())
    expect(onSubmit).toHaveBeenCalledWith(LABEL)
  })

  it('cannot confirm a label that sanitizes away', () => {
    const { input, confirmBtn, onSubmit } = openDialog()
    type(input, '{}<>')
    expect(confirmBtn.disabled).toBe(true)
    act(() => confirmBtn.click())
    expect(onSubmit).not.toHaveBeenCalled()
  })

  it('caps the label at the length the placeholder can carry', () => {
    const { input } = openDialog()
    expect(input.maxLength).toBe(40)
  })
})
