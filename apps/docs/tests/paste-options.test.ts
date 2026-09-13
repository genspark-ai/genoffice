/**
 * Paste options for foreign HTML: the default-mode setting, the per-paste
 * handshake shared by the paste entry points, and the merge-formatting
 * transform (keep emphasis and structure, format text and paragraphs like
 * typing at the insertion point).
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import type { Mark, Node as PmNode } from '@tiptap/pm/model'
import { DOMParser as PmDOMParser, Fragment, Slice as PmSlice } from '@tiptap/pm/model'
import type { EditorState, Transaction } from '@tiptap/pm/state'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  beginForeignPaste,
  capturePasteRestore,
  caretParaFormat,
  caretStyleMark,
  consumeForeignPaste,
  defaultPasteMode,
  forceNextPasteMode,
  lastForeignPasteMode,
  mergeFormattingFragment,
  revertPaste,
  setDefaultPasteMode,
  stashPastePayload,
  takePastePayload,
  takeTextReroute,
  type PasteMode,
  type PasteRestore,
} from '../src/renderer/editor/paste-options'
import { pasteTextSlice } from '../src/renderer/editor/paste-text'
import { isForeignPasteHtml } from '../src/renderer/editor/paste-web-html'

class FakeClipboardEvent extends Event {
  clipboardData = null
}
;(globalThis as Record<string, unknown>).ClipboardEvent = FakeClipboardEvent

const FOREIGN = `<meta charset='utf-8'><p style="font-family: Arial;">web text</p>`
const INTERNAL = `<div data-pm-slice="1 1 []"><p>ours</p></div>`

function makeEditor() {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { align: 'both', lineSpacing: 2, indentLeft: 240 },
          content: [
            {
              type: 'text',
              text: 'host',
              marks: [{ type: 'docTextStyle', attrs: { font: 'Calibri', fontAscii: 'Calibri' } }],
            },
          ],
        },
      ],
    },
  })
}

beforeEach(() => {
  localStorage.clear()
  beginForeignPaste('') // reset the handshake between tests
})

describe('default paste mode setting', () => {
  it('defaults to keep-source and persists changes', () => {
    expect(defaultPasteMode()).toBe('source')
    setDefaultPasteMode('merge')
    expect(defaultPasteMode()).toBe('merge')
    setDefaultPasteMode('source')
    expect(defaultPasteMode()).toBe('source')
  })

  it('ignores garbage stored values', () => {
    localStorage.setItem('aidocs.pasteFromOtherApps', 'banana')
    expect(defaultPasteMode()).toBe('source')
  })
})

describe('per-paste handshake', () => {
  it('arms only for foreign HTML', () => {
    expect(beginForeignPaste(FOREIGN)).toBe(true)
    expect(consumeForeignPaste()).toBe('source')
    expect(beginForeignPaste(INTERNAL)).toBe(false)
    expect(consumeForeignPaste()).toBeNull()
    expect(beginForeignPaste('')).toBe(false)
  })

  it('the armed mode follows the setting and a one-shot force wins once', () => {
    setDefaultPasteMode('merge')
    beginForeignPaste(FOREIGN)
    expect(consumeForeignPaste()).toBe('merge')
    forceNextPasteMode('text')
    beginForeignPaste(FOREIGN)
    expect(consumeForeignPaste()).toBe('text')
    beginForeignPaste(FOREIGN)
    expect(consumeForeignPaste()).toBe('merge')
  })

  it('text mode raises the reroute and records the consumed mode', () => {
    forceNextPasteMode('text')
    beginForeignPaste(FOREIGN)
    expect(takeTextReroute()).toBe(false)
    consumeForeignPaste()
    expect(lastForeignPasteMode()).toBe('text')
    expect(takeTextReroute()).toBe(true)
    expect(takeTextReroute()).toBe(false)
  })

  it('a new paste clears the previous paste state', () => {
    forceNextPasteMode('text')
    beginForeignPaste(FOREIGN)
    consumeForeignPaste()
    beginForeignPaste(FOREIGN)
    expect(lastForeignPasteMode()).toBeNull()
    expect(takeTextReroute()).toBe(false)
  })
})

describe('chip payload stash', () => {
  it('stores prose payloads and consumes once', () => {
    stashPastePayload({ html: FOREIGN, text: 'web text', mode: 'source' })
    expect(takePastePayload()?.text).toBe('web text')
    expect(takePastePayload()).toBeNull()
  })

  it('never stores image-only payloads', () => {
    stashPastePayload({ html: '<img src="https://x/y.png">', text: '', mode: 'source' })
    expect(takePastePayload()).toBeNull()
  })
})

describe('mergeFormattingFragment', () => {
  it('replaces run styles with the caret style, keeps emphasis, clones paragraph format', () => {
    const editor = makeEditor()
    const schema = editor.state.schema
    const $host = editor.state.doc.resolve(2)
    const style = caretStyleMark($host.marks(), schema)
    const paraAttrs = caretParaFormat($host, schema)
    expect(style?.attrs.font).toBe('Calibri')
    expect(paraAttrs).toMatchObject({ align: 'both', lineSpacing: 2, indentLeft: 240 })

    const arial = schema.marks.docTextStyle!.create({
      font: 'Arial',
      fontAscii: 'Arial',
      sizeHalfPoints: 24,
      color: '222222',
    })
    const bold = schema.marks.bold!.create()
    const pasted = Fragment.from([
      schema.nodes.docParagraph!.create(null, [
        schema.text('plain ', [arial]),
        schema.text('strong', [arial, bold]),
      ]),
      schema.nodes.docParagraph!.create(null, [schema.text('second', [arial])]),
    ])

    const merged = mergeFormattingFragment(pasted, schema, style, paraAttrs)
    const runs: Array<{ text: string; marks: readonly Mark[] }> = []
    const paras: PmNode[] = []
    merged.forEach((node) => {
      paras.push(node)
      node.descendants((child) => {
        if (child.isText) runs.push({ text: child.text ?? '', marks: child.marks })
        return true
      })
    })
    for (const para of paras) {
      expect(para.attrs).toMatchObject({ align: 'both', lineSpacing: 2, indentLeft: 240 })
    }
    for (const run of runs) {
      const docStyle = run.marks.find((mark) => mark.type.name === 'docTextStyle')
      expect(docStyle?.attrs).toMatchObject({ font: 'Calibri', fontAscii: 'Calibri' })
      expect(docStyle?.attrs.color).toBeNull()
    }
    expect(
      runs.find((run) => run.text === 'strong')?.marks.some((mark) => mark.type.name === 'bold'),
    ).toBe(true)
    editor.destroy()
  })

  it('with no caret style the pasted text follows the document default', () => {
    const editor = makeEditor()
    const schema = editor.state.schema
    const arial = schema.marks.docTextStyle!.create({ font: 'Arial', fontAscii: 'Arial' })
    const pasted = Fragment.from(
      schema.nodes.docParagraph!.create(null, [schema.text('web', [arial])]),
    )
    const merged = mergeFormattingFragment(pasted, schema, null, null)
    let styleSeen = false
    merged.forEach((node) =>
      node.descendants((child) => {
        if (child.isText && child.marks.some((mark) => mark.type.name === 'docTextStyle')) {
          styleSeen = true
        }
        return true
      }),
    )
    expect(styleSeen).toBe(false)
    editor.destroy()
  })
})

// ── chip re-apply: undo the paste, restore the insertion point, paste again ─

const ARIAL_P = `<meta charset='utf-8'><p><span style="font-family: Arial;">web text</span></p>`
const H1 = `<meta charset='utf-8'><h1><span style="font-family: Arial;">Title</span></h1>`

/** An editor with the two App paste lanes a mode change depends on: the
 *  transformPasted merge branch and the empty-paragraph wholesale lane
 *  (plus the text-mode reroute), all reading the current insertion point. */
function makePasteEditor() {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            {
              type: 'text',
              text: 'host',
              marks: [{ type: 'docTextStyle', attrs: { font: 'Calibri', fontAscii: 'Calibri' } }],
            },
          ],
        },
        { type: 'docParagraph' },
      ],
    },
    editorProps: {
      transformPasted: (slice, view) => {
        const mode = consumeForeignPaste()
        if (mode !== 'merge') return slice
        const schema = view.state.schema
        const marks = view.state.storedMarks ?? view.state.selection.$from.marks()
        return new PmSlice(
          mergeFormattingFragment(
            slice.content,
            schema,
            caretStyleMark(marks, schema),
            caretParaFormat(view.state.selection.$from, schema),
          ),
          slice.openStart,
          slice.openEnd,
        )
      },
      handlePaste: (view, event) => {
        const data = event.clipboardData!
        const html = data.getData('text/html')
        if (takeTextReroute()) {
          const slice = pasteTextSlice(data.getData('text/plain'), view.state.selection.$from, view)
          view.dispatch(view.state.tr.replaceSelection(slice))
          return true
        }
        const { $from, empty } = view.state.selection
        if (!(html && empty && $from.parent.isTextblock && $from.parent.content.size === 0)) {
          return false
        }
        const dom = new window.DOMParser().parseFromString(html, 'text/html')
        let content = PmDOMParser.fromSchema(view.state.schema).parse(dom.body).content
        if (isForeignPasteHtml(html) && lastForeignPasteMode() === 'merge') {
          const schema = view.state.schema
          const marks = view.state.storedMarks ?? $from.marks()
          content = mergeFormattingFragment(
            content,
            schema,
            caretStyleMark(marks, schema),
            caretParaFormat($from, schema),
          )
        }
        view.dispatch(view.state.tr.replaceWith($from.before(1), $from.after(1), content))
        return true
      },
    },
  })
}

function fakePasteEvent(html: string, text: string): ClipboardEvent {
  return {
    clipboardData: { getData: (type: string) => (type === 'text/html' ? html : text) },
  } as unknown as ClipboardEvent
}

/** The chip's handshake around a paste: remember the state the paste starts
 *  from, snapshot the restore when it lands. */
function pasteWithRestore(editor: Editor, html: string, text: string, mode: PasteMode) {
  let before: EditorState | null = null
  let restore: PasteRestore | null = null
  const onBefore = ({ transaction }: { transaction: Transaction }) => {
    if (transaction.docChanged) before = editor.state
  }
  const onTransaction = ({
    transaction,
    appendedTransactions,
  }: {
    transaction: Transaction
    appendedTransactions: Transaction[]
  }) => {
    if (transaction.docChanged && before) {
      restore = capturePasteRestore(before, [transaction, ...appendedTransactions])
    }
  }
  editor.on('beforeTransaction', onBefore)
  editor.on('transaction', onTransaction)
  forceNextPasteMode(mode)
  beginForeignPaste(html)
  editor.view.pasteHTML(html, fakePasteEvent(html, text))
  editor.off('beforeTransaction', onBefore)
  editor.off('transaction', onTransaction)
  return restore!
}

function reapply(
  editor: Editor,
  restore: PasteRestore,
  html: string,
  text: string,
  mode: PasteMode,
) {
  const revert = revertPaste(editor.state, restore)
  expect(revert).not.toBeNull()
  editor.view.dispatch(revert!)
  return pasteWithRestore(editor, html, text, mode)
}

function caretIntoEmptyParagraph(editor: Editor) {
  const calibri = editor.state.schema.marks.docTextStyle!.create({
    font: 'Calibri',
    fontAscii: 'Calibri',
  })
  const pos = editor.state.doc.content.size - 1
  editor.view.dispatch(
    editor.state.tr
      .setSelection(TextSelection.create(editor.state.doc, pos))
      .setStoredMarks([calibri]),
  )
}

function runFonts(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.isText) {
      const style = node.marks.find((mark) => mark.type.name === 'docTextStyle')
      out.push(`${node.text}:${style?.attrs.font ?? '-'}`)
    }
    return true
  })
  return out
}

const blockTypes = (editor: Editor) =>
  editor.state.doc.content.content.map((node) => node.type.name)

describe('chip re-apply restores the insertion point', () => {
  it('merge and text after a keep-source paste into an empty paragraph adopt the caret style', () => {
    const editor = makePasteEditor()
    caretIntoEmptyParagraph(editor)
    const pre = editor.state.doc
    let restore = pasteWithRestore(editor, ARIAL_P, 'web text', 'source')
    expect(runFonts(editor)).toEqual(['host:Calibri', 'web text:Arial'])

    restore = reapply(editor, restore, ARIAL_P, 'web text', 'merge')
    expect(runFonts(editor)).toEqual(['host:Calibri', 'web text:Calibri'])

    restore = reapply(editor, restore, ARIAL_P, 'web text', 'text')
    expect(runFonts(editor)).toEqual(['host:Calibri', 'web text:Calibri'])

    restore = reapply(editor, restore, ARIAL_P, 'web text', 'source')
    expect(runFonts(editor)).toEqual(['host:Calibri', 'web text:Arial'])

    const revert = revertPaste(editor.state, restore)!
    expect(revert.doc.eq(pre)).toBe(true)
    expect(revert.selection.empty).toBe(true)
    expect(revert.selection.$from.parent.content.size).toBe(0)
    editor.destroy()
  })

  it('a heading pasted onto an empty paragraph keeps the wholesale lane on every mode change', () => {
    const editor = makePasteEditor()
    caretIntoEmptyParagraph(editor)
    let restore = pasteWithRestore(editor, H1, 'Title', 'text')
    expect(blockTypes(editor)).toEqual(['docParagraph', 'docParagraph'])

    restore = reapply(editor, restore, H1, 'Title', 'source')
    expect(blockTypes(editor)).toEqual(['docParagraph', 'docHeading'])
    expect(runFonts(editor)).toEqual(['host:Calibri', 'Title:Arial'])

    restore = reapply(editor, restore, H1, 'Title', 'text')
    expect(blockTypes(editor)).toEqual(['docParagraph', 'docParagraph'])
    expect(runFonts(editor)).toEqual(['host:Calibri', 'Title:Calibri'])

    reapply(editor, restore, H1, 'Title', 'source')
    expect(blockTypes(editor)).toEqual(['docParagraph', 'docHeading'])
    editor.destroy()
  })

  it('a mode change is one undo step back to the original paste and is not tracked as a deletion', () => {
    vi.useFakeTimers()
    try {
      const editor = makePasteEditor()
      caretIntoEmptyParagraph(editor)
      const restore = pasteWithRestore(editor, ARIAL_P, 'web text', 'source')
      const afterPaste = editor.state.doc
      vi.advanceTimersByTime(2000)
      const revert = revertPaste(editor.state, restore)!
      expect(revert.getMeta('trackIgnore')).toBe(true)
      reapply(editor, restore, ARIAL_P, 'web text', 'merge')
      expect(runFonts(editor)).toEqual(['host:Calibri', 'web text:Calibri'])
      expect(editor.commands.undo()).toBe(true)
      expect(editor.state.doc.eq(afterPaste)).toBe(true)
      expect(runFonts(editor)).toEqual(['host:Calibri', 'web text:Arial'])
      editor.destroy()
    } finally {
      vi.useRealTimers()
    }
  })

  it('revertPaste refuses a document the paste no longer fits', () => {
    const editor = makePasteEditor()
    caretIntoEmptyParagraph(editor)
    const restore = pasteWithRestore(editor, ARIAL_P, 'web text', 'source')
    editor.commands.setContent('<p>replaced</p>')
    expect(revertPaste(editor.state, restore)).toBeNull()
    editor.destroy()
  })
})
