import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { DocRedaction } from '../src/renderer/editor/redaction'
import { executeOps } from '../src/renderer/ai/ops'

/**
 * Every model edit arrives here, so this is the one place a mangled marker can
 * be stopped. It runs before any op is validated, so a rejected batch leaves
 * the document exactly as it was.
 */

const SECRET = '13800138000'
const extensions = [...editorExtensions, DocRedaction]
const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

async function open(text: string): Promise<Editor> {
  const parsed = await parseDocx(
    await buildDocx({ bodyXml: `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>` }),
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions,
    content: { type: 'doc', content: blocksToPmDoc(parsed.blocks).content as JsonNode[] },
  })
  editors.add(editor)
  return editor
}

function selectText(editor: Editor, needle: string) {
  let start = -1
  editor.state.doc.descendants((node, pos) => {
    if (start >= 0) return
    const at = node.textContent.indexOf(needle)
    if (at >= 0) start = pos + 1 + at
  })
  expect(start).toBeGreaterThan(0)
  editor.view.dispatch(
    editor.state.tr.setSelection(
      TextSelection.create(editor.state.doc, start, start + needle.length),
    ),
  )
}

async function marked(): Promise<Editor> {
  const editor = await open(`Call ${SECRET} now`)
  selectText(editor, SECRET)
  editor.commands.setRedaction('client phone')
  return editor
}

/** the plain-text op every model edit funnels through */
const replaceOps = (find: string, replace: string) => [{ op: 'findReplace', find, replace }]

describe('the write guard on model output', () => {
  it('lets a batch that edits text beside a marker through', async () => {
    // the marker is not in the replaced text and is not in the replacement:
    // nothing about it changes, so the batch must not be refused
    const editor = await marked()
    const out = executeOps(editor, replaceOps('now', 'today, and again'))
    expect(out.ok, out.error).toBe(true)
    // and the marker is still there afterwards
    let still = 0
    editor.state.doc.descendants((n) => {
      if (n.marks.some((m) => m.type.name === 'redaction')) still++
    })
    expect(still).toBe(1)
  })

  it('leaves the secret itself alone when the edit lands beside it', async () => {
    // the model never saw the number, so it can only ever rewrite the words
    // around it — and doing so must not cost the reader a single character
    const editor = await marked()
    const out = executeOps(editor, replaceOps('now', 'Call them today'))
    expect(out.ok, out.error).toBe(true)
    expect(editor.state.doc.textContent).toBe(`Call ${SECRET} Call them today`)
    let withheld = 0
    editor.state.doc.descendants((n) => {
      if (n.isText && n.marks.some((m) => m.type.name === 'redaction')) {
        withheld += n.text?.length ?? 0
      }
    })
    expect(withheld).toBe(SECRET.length)
  })

  it('refuses a batch that overwrites the span itself', async () => {
    // the one way a span can actually be lost: the words themselves are the
    // find text. `find` matches real text, so this needs the model to have
    // echoed the secret from somewhere — but it has to be refused
    const editor = await marked()
    const out = executeOps(editor, replaceOps(SECRET, 'the number'))
    expect(out.ok).toBe(false)
    expect(out.error).toMatch(/placeholder/i)
  })

  it('refuses a batch that swaps the words back for the marker they stand in for', async () => {
    // The model's view is byte-identical before and after, so comparing markers
    // sees nothing wrong — and the reader's phone number is gone, replaced by
    // literal braces that mean nothing to anyone. Only a range check sees it.
    const editor = await marked()
    const before = editor.state.doc.textContent
    const out = executeOps(editor, replaceOps(SECRET, '{{client phone}}'))
    expect(out.ok, out.error).toBe(false)
    expect(editor.state.doc.textContent).toBe(before)
  })

  it('refuses a batch that eats part of the span and leaves the rest marked', async () => {
    // Half the secret is gone, the other half still wears the mark, so the
    // model's view still reads `{{client phone}}` and the count matches. The
    // document is silently destroyed and no marker comparison can tell.
    const editor = await marked()
    const before = editor.state.doc.textContent
    const out = executeOps(editor, replaceOps(SECRET.slice(0, -1), 'x'))
    expect(out.ok, out.error).toBe(false)
    expect(editor.state.doc.textContent).toBe(before)
  })

  it('refuses a batch that mangles the marker with a space', async () => {
    const editor = await marked()
    const out = executeOps(editor, replaceOps('now', '{{client phone}} today'))
    expect(out.ok).toBe(false)
    expect(out.error).toMatch(/placeholder/i)
  })

  it('leaves the document untouched after a refusal', async () => {
    const editor = await marked()
    const before = editor.state.doc.textContent
    const out = executeOps(editor, replaceOps('now', '{{client phone}} today'))
    expect(out.ok).toBe(false)
    expect(editor.state.doc.textContent).toBe(before)
  })

  it('does not guard a document with no spans', async () => {
    const editor = await open('nothing secret here')
    const out = executeOps(editor, replaceOps('now', '{{client phone}} today'))
    expect(out.ok, out.error).toBe(true)
  })

  it('lets a mail-merge template be filled', async () => {
    // a `{{token}}` the document already contained is content, not one of our
    // markers. A model asked to fill a template in is doing what it was asked,
    // and reading the filled token as a damaged marker refused the whole batch
    // — which is what stopped `genoffice merge` from filling a docx at all.
    const editor = await open('Invoice for {{name}}, due {{due.date}}')
    const out = executeOps(editor, replaceOps('{{name}}', 'Ada'))
    expect(out.ok, out.error).toBe(true)
    expect(editor.state.doc.textContent).toBe('Invoice for Ada, due {{due.date}}')
  })

  it('refuses a template fill in a block that also holds a span', async () => {
    // Deliberate, and the reason the two must not be confused: once a block
    // holds a withheld span it is policed whole, so a `{{token}}` in it is
    // treated as part of the marker bag and may not be filled in either. A
    // document that mixes a redaction with a mail-merge template has to be
    // filled before the span is marked, not after. Filling is still allowed
    // everywhere the model is shown the whole block.
    const editor = await open(`Call {{name}} about ${SECRET}`)
    selectText(editor, SECRET)
    editor.commands.setRedaction('client phone')
    const before = editor.state.doc.textContent
    const out = executeOps(editor, replaceOps('{{name}}', 'Ada'))
    expect(out.ok).toBe(false)
    expect(editor.state.doc.textContent).toBe(before)
  })

  it('still refuses a batch that renames a marker into a template token', async () => {
    // the token spelling is not a free pass: a marker that became ordinary
    // content is the one damage this guard exists to stop
    const editor = await marked()
    const out = executeOps(editor, replaceOps('now', 'and {{name}}'))
    expect(out.ok).toBe(false)
    expect(out.error).toMatch(/private placeholder/i)
  })

  it('reports the same reason whatever the damage', async () => {
    // three different ways to wreck a marker, one refusal. a marker that was
    // never there, a span overwritten, a marker broken in half
    const damages: Array<[string, string]> = [
      ['now', 'Call {{client phone}}'],
      [SECRET, 'the number'],
      ['now', '{{cli}ent{{pho}ne}'],
    ]
    for (const [find, replace] of damages) {
      const editor = await marked()
      const out = executeOps(editor, replaceOps(find, replace))
      expect(out.ok, `${find} -> ${replace}`).toBe(false)
      expect(out.error, `${find} -> ${replace}`).toMatch(/private placeholder/i)
    }
  })
})
