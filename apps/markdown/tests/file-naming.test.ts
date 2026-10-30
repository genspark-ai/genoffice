import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import StarterKit from '@tiptap/starter-kit'
import { Redaction } from '../src/renderer/editor/Redaction'
import { SelectiveEscapeMarkdown } from '../src/renderer/editor/markdownEscape'
import { modelSaveAsName, proposedSaveName } from '../src/renderer/file-naming'

/**
 * These go through the naming functions the save path and the ribbon actually
 * call, not through `suggestFileName` directly. A test of the shared client would
 * still pass if this app stopped using the projection — which is the failure
 * that leaks a secret, so the app's own call site is what needs covering.
 */

const SECRET = '13800138000'
const LABEL = '\u5ba2\u6237\u7535\u8bdd'

/** Calls the naming bridge received, in order. */
let sent: Array<{ content: string }> = []

/** What the model answers with; '' is a model that returned nothing usable. */
let answer = 'Quarterly notes'

/**
 * Every editor this file builds, torn down after each case.
 *
 * An EditorView left alive keeps ProseMirror's DOMObserver on a timer. Once the
 * file finishes, that timer fires against a jsdom that is already gone, and the
 * `document is not defined` it throws lands as an unhandled error — which fails
 * the whole run even when every assertion passed.
 */
const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function makeEditor(md: string): Editor {
  const editor = new Editor({ extensions: [StarterKit, SelectiveEscapeMarkdown, Redaction] })
  editor.commands.setContent(md, { contentType: 'markdown' })
  editors.add(editor)
  return editor
}

function selectText(editor: Editor, needle: string): void {
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

/** A document whose phone number the reader has hidden from the model. */
function editorWithRedaction(): Editor {
  const editor = makeEditor(
    '# \u5ba2\u6237\u56de\u8bbf\u8bb0\u5f55\n\n\u8bf7\u62e8\u6253 13800138000 \u786e\u8ba4\u8ba2\u5355',
  )
  selectText(editor, SECRET)
  editor.commands.setRedaction(LABEL)
  return editor
}

/**
 * Install the naming bridge the preload exposes. Assigned onto the real jsdom
 * window rather than replacing it: TipTap needs the DOM that lives there.
 */
function installBridge(api: {
  suggestFileName: (input: { content: string }) => Promise<unknown>
}): void {
  ;(window as unknown as { aiNamingApi: unknown }).aiNamingApi = api
}

beforeEach(() => {
  sent = []
  answer = 'Quarterly notes'
  installBridge({
    suggestFileName: async (input) => {
      sent.push({ content: input.content })
      return answer ? { ok: true, name: answer } : { ok: false, reason: 'empty-name' }
    },
  })
})

afterEach(() => {
  delete (window as unknown as { aiNamingApi?: unknown }).aiNamingApi
  vi.restoreAllMocks()
})

describe('the name the reader asks for', () => {
  it('never hands the model a value the reader withheld', async () => {
    const editor = editorWithRedaction()
    await modelSaveAsName(editor)
    // The one assertion that matters. The file name is written to disk, synced
    // and emailed, so a secret that reached it would outlive the redaction
    // meant to contain it.
    expect(sent).toHaveLength(1)
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(JSON.stringify(sent[0]!.content)).not.toContain('13800138')
  })

  it('sends the marker the reader chose, so the model can still read the line', async () => {
    const editor = editorWithRedaction()
    await modelSaveAsName(editor)
    expect(sent[0]!.content).toContain(LABEL)
    expect(sent[0]!.content).toContain('\u5ba2\u6237\u56de\u8bbf\u8bb0\u5f55')
  })

  it('sends the whole document for one with nothing withheld', async () => {
    const editor = makeEditor('# Release notes\n\nthe build is green')
    await modelSaveAsName(editor)
    expect(sent[0]!.content).toContain('Release notes')
    expect(sent[0]!.content).toContain('the build is green')
  })

  it('returns the name the model proposed', async () => {
    const editor = makeEditor('# Release notes')
    expect(await modelSaveAsName(editor)).toBe('Quarterly notes')
    expect(sent).toHaveLength(1)
  })

  it('returns nothing when the model declines, leaving the dialog on the current name', async () => {
    answer = ''
    const editor = makeEditor('# Release notes')
    expect(await modelSaveAsName(editor)).toBe('')
  })

  it('returns nothing when the model errors', async () => {
    installBridge({
      suggestFileName: async () => {
        throw new Error('provider unreachable')
      },
    })
    const editor = makeEditor('# Release notes')
    expect(await modelSaveAsName(editor)).toBe('')
  })

  it('returns nothing when there is no naming bridge at all', async () => {
    delete (window as unknown as { aiNamingApi?: unknown }).aiNamingApi
    const editor = makeEditor('# Release notes')
    expect(await modelSaveAsName(editor)).toBe('')
  })

  it('does not ask the model about a document with nothing to name', async () => {
    const editor = makeEditor('')
    expect(await modelSaveAsName(editor)).toBe('')
    expect(sent).toHaveLength(0)
  })

  it('is repeatable: a second request asks again', async () => {
    // The reader changed their mind and clicked again: a fresh answer, not a
    // cached one.
    const editor = makeEditor('# Release notes')
    await modelSaveAsName(editor)
    await modelSaveAsName(editor)
    expect(sent).toHaveLength(2)
  })
})

/**
 * The save path's own rule: which name, if any, a save offers.
 *
 * `proposedSaveName` is what `App.tsx` asks before every save, so these are
 * the cases a reader feels rather than the cases the naming module feels.
 */
describe('the name a save offers', () => {
  it('names a document that has never been on disk from its own content', async () => {
    // No model call: a save must not wait on a model round-trip to learn what
    // the document already says about itself.
    const editor = makeEditor('# \u5ba2\u6237\u56de\u8bbf\u8bb0\u5f55\n\n\u6b63\u6587')
    expect(await proposedSaveName(editor, { filePath: null })).toBe(
      '\u5ba2\u6237\u56de\u8bbf\u8bb0\u5f55',
    )
    expect(sent).toHaveLength(0)
  })

  it('never renames a document that already has a name', async () => {
    // The rule the whole feature rests on: a file that is open elsewhere,
    // referenced from a note, or synced keeps the name it has.
    const editor = makeEditor('# \u5ba2\u6237\u56de\u8bbf\u8bb0\u5f55\n\n\u6b63\u6587')
    expect(await proposedSaveName(editor, { filePath: '/tmp/notes.md' })).toBeUndefined()
    expect(sent).toHaveLength(0)
  })

  it('lets a name the reader asked for win over the document own name', async () => {
    // The ribbon's "Name with AI" proposes a name for a document that already
    // has one. That is the reader asking, so nothing else gets a say — and
    // the model is not asked a second time.
    const editor = makeEditor('# \u5ba2\u6237\u56de\u8bbf\u8bb0\u5f55\n\n\u6b63\u6587')
    expect(
      await proposedSaveName(editor, {
        filePath: '/tmp/notes.md',
        proposedName: 'Quarterly notes',
      }),
    ).toBe('Quarterly notes')
    expect(sent).toHaveLength(0)
  })

  it('does not ask again when the save already carries a name of its own', async () => {
    // The AI auto-naming path writes an untitled document under a name it
    // derived itself; there is nothing left to ask about.
    const editor = makeEditor('# \u5ba2\u6237\u56de\u8bbf\u8bb0\u5f55\n\n\u6b63\u6587')
    expect(
      await proposedSaveName(editor, {
        filePath: null,
        suggestedName: '\u5ba2\u6237\u56de\u8bbf\u8bb0\u5f55',
      }),
    ).toBeUndefined()
    expect(sent).toHaveLength(0)
  })

  it('offers nothing for a document with no content to name', async () => {
    const editor = makeEditor('')
    expect(await proposedSaveName(editor, { filePath: null })).toBeUndefined()
    expect(sent).toHaveLength(0)
  })
})
