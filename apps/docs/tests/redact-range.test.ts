import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { redactTextBetween, type RedactableNode } from '../src/renderer/ai/redact-view'
import { buildDocContext, serializeRangeToHtml } from '../src/renderer/ai/protocol'
import { listRevisionEntries } from '../src/renderer/ai/revision-ops'
import { resolveQueueItem } from '../src/renderer/ai/edit-queue'
import { addQueueAnchor } from '../src/renderer/editor/ai-queue-anchors'
import { REDACT_MARK } from '../src/renderer/ai/redact'
import { executeTool } from '../src/renderer/ai/tools'

/**
 * The model-facing text of a *range*, which is what a selection preview quotes.
 *
 * A selection is where a leak hides: the context sent alongside a partial
 * selection repeats the span back, and the span is what the reader withheld.
 * `doc.textBetween` reads the words, so every call site on the model's side of
 * the app has to go through here instead.
 *
 * The first group is the load-bearing one. With nothing withheld this has to be
 * `doc.textBetween`, character for character — otherwise it would change what
 * every unredacted document sends, and the offsets would drift from
 * ProseMirror's without anything noticing until a selection came out garbled.
 */

const SECRET = '13800138000'
// editorExtensions already carries DocRedaction; adding it again registers a
// second mark of the same name and the editor stops being the one users get
const extensions = editorExtensions
const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

/** a two-paragraph document, optionally with the secret marked */
async function open(withMark: boolean): Promise<Editor> {
  const bodyXml =
    '<w:p><w:r><w:t xml:space="preserve">Call </w:t></w:r>' +
    `<w:r><w:t>${SECRET}</w:t></w:r>` +
    '<w:r><w:t xml:space="preserve"> about the invoice</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t xml:space="preserve">Second block here</w:t></w:r></w:p>'
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions,
    content: { type: 'doc', content: blocksToPmDoc(parsed.blocks).content as never[] },
  })
  editors.add(editor)
  if (withMark) {
    let start = -1
    editor.state.doc.descendants((node, pos) => {
      if (start >= 0) return
      const at = node.textContent.indexOf(SECRET)
      if (at >= 0) start = pos + 1 + at
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, start, start + SECRET.length),
      ),
    )
    editor.commands.setRedaction('客户电话')
  }
  return editor
}

const json = (editor: Editor): RedactableNode => editor.getJSON() as RedactableNode

/** words, so a comparison is about the text and not about separators */
const words = (s: string): string[] => s.trim().split(/\s+/).filter(Boolean)

describe('a range with nothing withheld reads as ProseMirror reads it', () => {
  /**
   * Positions a real selection can be at: a block's first character, a run
   * boundary, and the document's end. An exhaustive sweep of every integer is
   * not a thing ProseMirror's own `textBetween` accepts — some integers fall
   * inside a node's opening token — and a selection never lands on one.
   */
  function positions(editor: Editor): number[] {
    const out = new Set<number>([1])
    const end = editor.state.doc.content.size
    editor.state.doc.descendants((node, pos) => {
      if (node.isText) {
        out.add(pos + 1)
        out.add(pos + node.nodeSize)
      }
    })
    out.add(end)
    return [...out].filter((p) => p >= 1 && p <= end).sort((a, b) => a - b)
  }

  it('carries the same words, whichever two positions it spans', async () => {
    const editor = await open(false)
    const doc = json(editor)
    const at = positions(editor)
    for (const from of at) {
      for (const to of at) {
        if (to <= from) continue
        expect(words(redactTextBetween(doc, from, to)), `range ${from}..${to}`).toEqual(
          words(editor.state.doc.textBetween(from, to, '\n', ' ')),
        )
      }
    }
  })

  it('reads across several runs inside one block', async () => {
    // The docx round trip merges adjacent runs of equal formatting into one
    // text node, so a plain paragraph has exactly one child and never exercises
    // the walk past its first step. Differently-formatted runs stay separate,
    // and this is the shape that catches an off-by-one in the offsets.
    const editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: null },
            content: [
              { type: 'text', text: 'Base ', marks: [{ type: 'bold' }] },
              { type: 'text', text: 'added ', marks: [{ type: 'ins' }] },
              { type: 'text', text: 'and ' },
              { type: 'text', text: 'gone ', marks: [{ type: 'del' }] },
              { type: 'text', text: 'words.' },
            ],
          },
        ],
      },
    })
    editors.add(editor)
    const doc = json(editor)
    const end = editor.state.doc.content.size
    for (let from = 1; from <= end; from++) {
      expect(redactTextBetween(doc, from, end).replace(/\s+/g, ''), `from ${from}`).toBe(
        editor.state.doc.textContent.slice(from - 1).replace(/\s+/g, ''),
      )
    }
  })

  it('starts on the character the range starts on', async () => {
    // The slicing arithmetic is the part that can go quietly wrong, and an
    // off-by-one drops the first character — which a word-for-word comparison
    // of short words would not notice.
    const editor = await open(false)
    const doc = json(editor)
    const end = editor.state.doc.content.size
    const full = editor.state.doc.textContent
    for (let from = 1; from <= 20; from++) {
      expect(redactTextBetween(doc, from, end).replace(/\s+/g, ''), `from ${from}`).toBe(
        full.slice(from - 1).replace(/\s+/g, ''),
      )
    }
  })
})

describe('a range with a withheld span in it', () => {
  it('never returns the words the reader hid', async () => {
    const editor = await open(true)
    const doc = json(editor)
    const end = editor.state.doc.content.size
    for (let from = 1; from <= end; from++) {
      for (const to of [from + 1, from + 5, end]) {
        if (to <= from) continue
        expect(redactTextBetween(doc, from, to), `range ${from}..${to}`).not.toContain(SECRET)
      }
    }
  })

  it('stands the marker in for the whole span, even for a range that clips it', async () => {
    const editor = await open(true)
    const doc = json(editor)
    const end = editor.state.doc.content.size
    // a range that starts inside the secret and ends well before it finishes
    for (let from = 1; from < end; from++) {
      const text = redactTextBetween(doc, from, Math.min(from + 4, end))
      if (!text.includes('客户')) continue
      // a fragment of the marker would read as a typo the model should repair
      expect(text).toContain('{{客户电话}}')
      // a marker that opens and never closes, or closes without opening: a
      // fragment the model would read as a typo to repair
      expect(text).not.toMatch(/\{\{[^}]*$/)
      expect(text).not.toMatch(/^[^{]*\}\}/)
    }
  })

  it('leaves the text beside the span alone', async () => {
    const editor = await open(true)
    const text = redactTextBetween(json(editor), 1, editor.state.doc.content.size)
    expect(text).toContain('Call ')
    expect(text).toContain('about the invoice')
    expect(text).toContain('Second block here')
  })
})

/**
 * The call sites, not just the helper.
 *
 * A projection nothing routes through is a projection that does nothing, and
 * the context a model gets for a partial selection quotes that selection back
 * to it — which is the exact string a withheld span would have escaped through.
 */
describe('what a selection sends to the model', () => {
  /** select the secret, the way a reader dragging over a number would */
  function selectSecret(editor: Editor): void {
    let start = -1
    editor.state.doc.descendants((node, pos) => {
      if (start >= 0) return
      const at = node.textContent.indexOf(SECRET)
      if (at >= 0) start = pos + 1 + at
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, start, start + SECRET.length),
      ),
    )
  }

  it('quotes the selection with the marker, not the words', async () => {
    const editor = await open(true)
    selectSecret(editor)
    const context = buildDocContext(editor)
    expect(context).not.toContain(SECRET)
    expect(context).toContain('{{客户电话}}')
  })

  it('serialises the range with the marker too', async () => {
    const editor = await open(true)
    const html = serializeRangeToHtml(editor, 0, 1)
    expect(html).not.toContain(SECRET)
    expect(html).toContain('{{客户电话}}')
  })

  it('leaves an unmarked document exactly as it was', async () => {
    const editor = await open(false)
    selectSecret(editor)
    // nothing is withheld, so the context still carries the real text: the
    // filter has to be a no-op here, not a blanking
    expect(buildDocContext(editor)).toContain(SECRET)
  })
})

/**
 * The other three paths the guard has to cover, each a `textBetween` that used
 * to read the words.
 *
 * They are the same mistake in four places, which is why they are pinned
 * together: a projection nothing routes through is a projection that does
 * nothing, and the next one added would read the words again.
 */
describe('the paths that repeat the document to the model', () => {
  it('a revision entry carries the marker, not the words', async () => {
    // one run that is both an insertion and withheld, which is the case that
    // leaks: `read_revisions` sends each entry's text to the model so it can
    // target them by id
    const author = { author: 'Carol', date: '2026-03-01T10:00:00Z' }
    const editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: null },
            content: [
              {
                type: 'text',
                text: SECRET,
                marks: [
                  { type: 'ins', attrs: author },
                  { type: REDACT_MARK, attrs: { label: '客户电话' } },
                ],
              },
            ],
          },
        ],
      },
    })
    editors.add(editor)

    const entries = listRevisionEntries(editor.state.doc as never)
    // nothing to assert against if no entry came back
    expect(entries.length).toBeGreaterThan(0)
    for (const entry of entries) {
      expect(String(entry.text ?? ''), entry.id).not.toContain(SECRET)
    }
    expect(entries.some((e) => String(e.text ?? '').includes('{{客户电话}}'))).toBe(true)
  })

  it('a queued edit labels its anchor with the marker', async () => {
    const editor = await open(true)
    // an anchor over the withheld run, which is what selecting it to annotate
    // produces
    let from = -1
    let to = -1
    editor.state.doc.descendants((node, pos) => {
      if (from >= 0) return
      if (node.marks.some((m) => m.type.name === REDACT_MARK)) {
        from = pos + 1
        to = pos + 1 + (node.text?.length ?? 0)
      }
    })
    expect(from).toBeGreaterThan(0)
    addQueueAnchor(editor, 'q1', from, to)

    const resolved = resolveQueueItem(editor, {
      qid: 'q1',
      instruction: 'tighten this',
      capturedText: 'captured',
    })
    // the target has to exist, or the assertion below would pass on nothing
    expect(resolved.target).not.toBeNull()
    expect(String(resolved.target?.excerpt ?? '')).not.toContain(SECRET)
    expect(String(resolved.target?.excerpt ?? '')).toContain('{{客户电话}}')
  })
})

describe('the echo a replacement sends back', () => {
  const NUM_IDS = { bullet: null, ordered: null }

  function marked(): Editor {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: null },
            content: [
              { type: 'text', text: 'Call ' },
              {
                type: 'text',
                text: SECRET,
                marks: [{ type: REDACT_MARK, attrs: { label: '客户电话' } }],
              },
              { type: 'text', text: ' now' },
            ],
          },
        ],
      },
    })
    editors.add(editor)
    return editor
  }

  function selectSpan(editor: Editor, length: number): void {
    let start = -1
    editor.state.doc.descendants((node, pos) => {
      if (start >= 0) return
      const at = node.textContent.indexOf(SECRET)
      if (at >= 0) start = pos + 1 + at
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, start, start + length)),
    )
  }

  /** the tool is synchronous for this input; a promise would be a silent no-op here */
  const run = (editor: Editor): string => {
    const out = executeTool(
      editor,
      { id: 't', name: 'replace_selection', input: { html: 'redacted' } },
      NUM_IDS,
    )
    if (out instanceof Promise) throw new Error('replace_selection should be sync')
    return String(out.output ?? '')
  }

  it('quotes the marker when the selection covers the whole span', () => {
    const editor = marked()
    selectSpan(editor, SECRET.length)
    const out = run(editor)
    expect(out).toContain('{{客户电话}}')
    expect(out).not.toContain(SECRET)
  })

  it('quotes the marker when the selection only clips into it', () => {
    // the reachable case that matters: `replace_selection` runs on a partial
    // overlap where the write guard's range check does not stop it, and its
    // echo is what goes back to the model
    const editor = marked()
    selectSpan(editor, 3)
    const out = run(editor)
    expect(out).toContain('{{客户电话}}')
    expect(out).not.toContain(SECRET)
  })

  it('still echoes the real text when nothing is withheld there', () => {
    // the filter has to be a no-op on ordinary text, not a blanking
    const editor = marked()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 6)))
    const plain = executeTool(
      editor,
      { id: 't', name: 'replace_selection', input: { html: 'Ring' } },
      NUM_IDS,
    )
    if (plain instanceof Promise) throw new Error('replace_selection should be sync')
    const out = String(plain.output ?? '')
    expect(out).toContain('Call ')
  })
})
