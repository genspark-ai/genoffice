import { describe, afterEach, beforeEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import StarterKit from '@tiptap/starter-kit'
import { Redaction } from '../src/renderer/editor/Redaction'
import { SelectiveEscapeMarkdown } from '../src/renderer/editor/markdownEscape'
import {
  checkPlaceholders,
  redactJson,
  collectPlaceholders,
  placeholderInstruction,
  placeholderSource,
  sanitizeLabel,
} from '../src/renderer/editor/redact'

/**
 * Every editor this file builds, torn down after each case.
 *
 * An EditorView left alive keeps ProseMirror's DOMObserver on a timer. Once the
 * file finishes, that timer fires against a jsdom that is already gone, and
 * the `document is not defined` it throws lands as an unhandled error — which
 * fails the whole run even when every assertion passed.
 */
const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

function makeEditor(md = '') {
  const editor = new Editor({
    extensions: [StarterKit, SelectiveEscapeMarkdown, Redaction],
  })
  if (md) editor.commands.setContent(md, { contentType: 'markdown' })
  editors.add(editor)
  return editor
}

const toMd = (editor: Editor) => editor.markdown!.serialize(editor.getJSON())
/** what the model is actually shown: the same document through redactJson */
const toModelView = (editor: Editor) => editor.markdown!.serialize(redactJson(editor.getJSON()))

/** select by character offset inside the paragraph's text */
function selectText(editor: Editor, needle: string) {
  let start = -1
  editor.state.doc.descendants((node, pos) => {
    if (start >= 0) return
    const text = node.textContent
    const at = text.indexOf(needle)
    if (at >= 0) start = pos + 1 + at
  })
  expect(start).toBeGreaterThan(0)
  editor.view.dispatch(
    editor.state.tr.setSelection(
      TextSelection.create(editor.state.doc, start, start + needle.length),
    ),
  )
}

describe('marking a span keeps the reader’s text', () => {
  let editor: Editor
  beforeEach(() => {
    editor = makeEditor('call 13800138000 to confirm the order')
  })

  it('leaves the real number in the document', () => {
    selectText(editor, '13800138000')
    editor.commands.setRedaction('client phone')
    // the whole point: nothing of the reader's own data was thrown away
    expect(editor.state.doc.textContent).toContain('13800138000')
  })

  it('shows the model a marker instead of the number', () => {
    selectText(editor, '13800138000')
    editor.commands.setRedaction('client phone')
    expect(toModelView(editor)).toContain('{{client phone}}')
    expect(toModelView(editor)).not.toContain('13800138000')
  })

  it('keeps the surrounding sentence intact', () => {
    selectText(editor, '13800138000')
    editor.commands.setRedaction('client phone')
    expect(toMd(editor)).toContain('call ')
    expect(toMd(editor)).toContain('confirm the order')
  })

  it('marks exactly the selection, not a character more or less', () => {
    selectText(editor, '13800138000')
    editor.commands.setRedaction('client phone')
    // a leaked or clipped digit would show up as a stray number here
    expect(toMd(editor)).not.toMatch(/1380013800[^0]/)
    expect(editor.state.doc.textContent).toBe('call 13800138000 to confirm the order')
  })

  it('unmarking brings the text back and leaves it alone', () => {
    selectText(editor, '13800138000')
    editor.commands.setRedaction('client phone')
    editor.commands.unsetRedaction()
    expect(editor.state.doc.textContent).toBe('call 13800138000 to confirm the order')
    expect(toMd(editor)).not.toContain('{{')
  })

  it('refuses an empty label', () => {
    selectText(editor, '13800138000')
    expect(editor.commands.setRedaction('  ')).toBe(false)
  })
})

describe('two spans in one document', () => {
  it('each becomes its own marker', () => {
    const editor = makeEditor('call 13800138000 or ship to Pudong New Area, Shanghai')
    selectText(editor, '13800138000')
    editor.commands.setRedaction('client phone')
    selectText(editor, 'Pudong New Area, Shanghai')
    editor.commands.setRedaction('shipping address')
    const md = toModelView(editor)
    expect(md).toContain('{{client phone}}')
    expect(md).toContain('{{shipping address}}')
    expect(md).not.toContain('13800138000')
    expect(editor.state.doc.textContent).toContain('Pudong New Area, Shanghai')
  })
})

describe('what the model is told', () => {
  it('lists the markers this document actually has', () => {
    const text = placeholderInstruction(['client phone', 'shipping address', 'client phone'])
    expect(text).toContain('- {{client phone}}')
    expect(text).toContain('- {{shipping address}}')
    // a repeated label is one entry, not two
    expect(text.split('\n').filter((l) => l === '- {{client phone}}')).toHaveLength(1)
  })

  it('says not to split, merge, rename or drop one', () => {
    const text = placeholderInstruction(['a'])
    expect(text).toMatch(/never split/i)
    expect(text).toMatch(/never merge/i)
    expect(text).toMatch(/never rename/i)
    expect(text).toMatch(/never drop/i)
  })
})

describe('the write guard still applies to a span', () => {
  const before = 'call {{client phone}} to confirm the order'

  it('accepts a reply that keeps the marker', () => {
    expect(
      checkPlaceholders(before, 'please call {{client phone}} before the end of today.'),
    ).toEqual([])
  })

  it('rejects a reply that splits the marker across a line break', () => {
    expect(
      checkPlaceholders(before, 'call {{client\nphone}} to confirm the order').length,
    ).toBeGreaterThan(0)
  })

  it('rejects a reply that interleaves two markers', () => {
    // a span is several characters wide, so this is the failure mode an atom
    // could never have — the guard is what stops it
    expect(checkPlaceholders('{{a}}{{b}}', '{{{{a}}b}}').length).toBeGreaterThan(0)
  })

  it('rejects a reply that drops one', () => {
    expect(checkPlaceholders('{{a}} and {{b}}', '{{a}}').length).toBeGreaterThan(0)
  })

  it('rejects a half-typed marker', () => {
    expect(checkPlaceholders('{{a}}', '{{a}').some((i) => i.reason === 'split')).toBe(true)
  })
})

describe('label cleaning', () => {
  it('strips characters that would make the marker ambiguous', () => {
    expect(sanitizeLabel('a{b}<c>"d"')).toBe('abcd')
  })

  it('falls back to a neutral word', () => {
    expect(placeholderSource('')).toBe('{{private}}')
  })

  it('caps the length', () => {
    expect(sanitizeLabel('x'.repeat(100)).length).toBeLessThanOrEqual(40)
  })
})

describe('collectPlaceholders', () => {
  it('lists markers, keeping duplicates', () => {
    expect(collectPlaceholders('{{a}} then {{b}} then {{a}}')).toEqual(['{{a}}', '{{b}}', '{{a}}'])
  })
})

describe('the two views of the document never cross', () => {
  it('the file keeps the real text; only the model view is redacted', () => {
    const editor = makeEditor('call 13800138000 to confirm the order')
    selectText(editor, '13800138000')
    editor.commands.setRedaction('client phone')

    // what a save writes
    const onDisk = editor.getMarkdown()
    // what a model request carries
    const toModel = toModelView(editor)

    expect(onDisk).toContain('13800138000')
    expect(onDisk).not.toContain('{{client phone}}')
    expect(toModel).not.toContain('13800138000')
    expect(toModel).toContain('{{client phone}}')
  })

  it('redactJson is not reachable from the save path', () => {
    // a guard against someone "simplifying" the save to reuse the model view:
    // the one thing that must never happen is the secret reaching the file
    const editor = makeEditor('show 310101199001011234 to enter')
    selectText(editor, '310101199001011234')
    editor.commands.setRedaction('ID number')
    expect(editor.getMarkdown()).toContain('310101199001011234')
  })
})
