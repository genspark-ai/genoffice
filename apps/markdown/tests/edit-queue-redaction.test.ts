import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { Redaction } from '../src/renderer/editor/Redaction'
import { addQueueAnchor } from '../src/renderer/editor/aiQueueAnchors'
import { resolveQueueItem } from '../src/renderer/ai/edit-queue'
import { REDACT_MARK } from '../src/renderer/editor/redact'

/**
 * A queued edit's anchor excerpt is shown to the model, so an anchor that
 * covered a withheld span would carry the words along — the same leak docs
 * closed, in the file that reads the document a fifth time.
 */

const SECRET = '13800138000'
const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

function markedEditor(): Editor {
  const editor = new Editor({
    extensions: [
      ...buildExtensions({
        slashController: {
          onOpen: () => {},
          onUpdate: () => {},
          onKeyDown: () => false,
          onClose: () => {},
        },
        slashItems: () => [],
      }),
      Redaction,
    ],
    content: '',
  })
  editor.commands.setContent(`Call ${SECRET} about the invoice`, { contentType: 'markdown' })
  editors.add(editor)
  return editor
}

/** select the secret and withhold it, the way a reader would */
function withhold(editor: Editor): void {
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

function anchorOverTheMarkedRun(editor: Editor): void {
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
}

describe('a queued edit over a withheld span', () => {
  it('labels its anchor with the marker, not the words', () => {
    const editor = markedEditor()
    withhold(editor)
    anchorOverTheMarkedRun(editor)
    const r = resolveQueueItem(editor, {
      qid: 'q1',
      instruction: 'tighten this',
      capturedText: '',
    })
    // the target has to exist, or the assertion below would pass on nothing
    expect(r.target).not.toBeNull()
    expect(r.target!.excerpt).not.toContain(SECRET)
    expect(r.target!.excerpt).toContain('{{客户电话}}')
  })

  it('leaves an ordinary anchor reading the real text', () => {
    // the projection has to be a no-op here, not a blanking
    const editor = markedEditor()
    addQueueAnchor(editor, 'q2', 1, 6)
    const r = resolveQueueItem(editor, {
      qid: 'q2',
      instruction: 'x',
      capturedText: '',
    })
    expect(r.target!.excerpt).toBe('Call')
  })
})
