import { afterEach, describe, expect, it } from 'vitest'
import { Editor, type JSONContent } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blockHasManualOverride,
  storyMarkerFor,
  storyViewKey,
} from '../src/renderer/editor/story-view'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

const para = (text: string): JSONContent => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: text ? [{ type: 'text', text }] : [],
})

const heading = (level: number, text: string): JSONContent => ({
  type: 'docHeading',
  attrs: { level },
  content: [{ type: 'text', text }],
})

function createEditor(content: JSONContent[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.push(editor)
  return editor
}

/** what the view does on a keypress: insert text carrying the stored marks
    (PM falls back to $from.marks() when state.storedMarks is null) */
function typeText(editor: Editor, text: string): void {
  const stored = editor.view.state.storedMarks ?? editor.state.selection.$from.marks()
  editor.view.dispatch(
    editor.state.tr.insert(
      editor.state.selection.head,
      editor.schema.text(text, Array.from(stored)),
    ),
  )
}

describe('storyMarkerFor', () => {
  it('headings get one # per level (capped at six)', () => {
    const editor = createEditor([heading(1, 'A'), heading(3, 'B'), heading(6, 'C')])
    const doc = editor.state.doc
    expect(storyMarkerFor(doc.child(0))).toBe('#')
    expect(storyMarkerFor(doc.child(1))).toBe('###')
    expect(storyMarkerFor(doc.child(2))).toBe('######')
  })

  it('list paragraphs get a dash by their numId, plain paragraphs get nothing', () => {
    const editor = createEditor([para('plain')])
    const doc = editor.state.doc
    expect(storyMarkerFor(doc.child(0))).toBeNull()
    const listed = editor.schema.nodes.docListItem.create({ numId: '3', ilvl: 0 }, [])
    expect(storyMarkerFor(listed)).toBe('-')
  })

  it('paragraphs in a quote style get a chevron', () => {
    const editor = createEditor([para('x')])
    const quoted = editor.schema.nodes.docParagraph.create({ styleId: 'Quote' }, [])
    expect(storyMarkerFor(quoted)).toBe('>')
    const intense = editor.schema.nodes.docParagraph.create({ styleId: 'IntenseQuote' }, [])
    expect(storyMarkerFor(intense)).toBe('>')
    const plain = editor.schema.nodes.docParagraph.create({ styleId: 'Normal' }, [])
    expect(storyMarkerFor(plain)).toBeNull()
  })
})

describe('blockHasManualOverride', () => {
  it('fires on a run with hidden formatting, not on emphasis-only runs', () => {
    const editor = createEditor([para('x')])
    const schema = editor.schema
    const style = schema.marks.docTextStyle
    const red = schema.text('red', [style.create({ color: 'FF0000' })])
    expect(blockHasManualOverride(schema.nodes.docParagraph.create(null, red))).toBe(true)

    const bold = schema.text('b', [schema.marks.bold.create()])
    expect(blockHasManualOverride(schema.nodes.docParagraph.create(null, bold))).toBe(false)
  })
})

describe('the typing semantic: inserted text drops the hidden run formatting', () => {
  it('typing after a red run takes the style formatting, not red', () => {
    const editor = createEditor([para('')])
    const schema = editor.schema
    const style = schema.marks.docTextStyle
    const pos = 1 // inside the empty paragraph
    editor.view.dispatch(
      editor.state.tr.insert(
        pos,
        schema.text('red', [style.create({ color: 'FF0000', sizeHalfPoints: 28 })]),
      ),
    )
    // caret right after the red run: stored marks pick up its docTextStyle
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 3)),
    )
    setActive(editor, true)
    typeText(editor, 'x')

    const paraNode = editor.state.doc.child(0)
    expect(paraNode.textContent).toBe('redx')
    paraNode.descendants((n) => {
      if (n.isText) {
        expect(n.marks.some((m) => m.type.name === 'docTextStyle')).toBe(false)
      }
      return true
    })
  })

  it('leaves the emphasis marks on inserted text alone', () => {
    const editor = createEditor([para('word')])
    const schema = editor.schema
    setActive(editor, true)
    const caret = 1 // before "word"
    editor.view.dispatch(
      editor.state.tr.insert(caret, schema.text('X', [schema.marks.bold.create()])),
    )
    editor.state.doc.child(0).descendants((n) => {
      if (n.isText && n.textContent === 'X') {
        expect(n.marks.some((m) => m.type.name === 'bold')).toBe(true)
      }
      return true
    })
  })

  it('does not strip while the view is off', () => {
    const editor = createEditor([para('plain')])
    const schema = editor.schema
    const caret = editor.state.doc.child(0).content.size + 1
    editor.view.dispatch(
      editor.state.tr.insert(
        caret,
        schema.text('y', [schema.marks.docTextStyle.create({ color: 'FF0000' })]),
      ),
    )
    let sawRed = false
    editor.state.doc.child(0).descendants((n) => {
      if (n.isText && n.marks.some((m) => m.type.name === 'docTextStyle')) sawRed = true
      return true
    })
    expect(sawRed).toBe(true)
    void schema
  })

  it('clears the decorations when the view switches off', () => {
    const editor = createEditor([heading(1, 'H'), para('b')])
    setActive(editor, true)
    expect(storyViewKey.getState(editor.state)!.decorations.find().length).toBeGreaterThan(0)
    setActive(editor, false)
    expect(storyViewKey.getState(editor.state)!.decorations.find()).toHaveLength(0)
  })
})

describe('decorations', () => {
  it('one gutter widget per top-level block, plus markers', () => {
    const editor = createEditor([heading(2, 'H'), para('body')])
    setActive(editor, true)
    const decos = storyViewKey.getState(editor.state)!.decorations
    // 2 gutters + 1 heading marker
    expect(decos.find(0, editor.state.doc.content.size)).toHaveLength(3)
  })
})

function setActive(editor: Editor, active: boolean): void {
  editor.commands.setStoryViewActive(active)
}
