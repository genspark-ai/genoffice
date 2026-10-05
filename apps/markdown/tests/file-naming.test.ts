import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import StarterKit from '@tiptap/starter-kit'
import { Redaction } from '../src/renderer/editor/Redaction'
import { SelectiveEscapeMarkdown } from '../src/renderer/editor/markdownEscape'
import {
  firstSaveName,
  modelSaveAsName,
  noteDocumentSwapped,
} from '../src/renderer/file-naming'

/**
 * These go through the naming functions the save path and the ribbon actually
 * call, not through `nameForSave` directly. A test of the shared client would
 * still pass if this app stopped using the projection — which is the failure
 * that leaks a secret, so the app's own call site is what needs covering.
 */

const SECRET = '13800138000'
const LABEL = '客户电话'

/** Calls the naming bridge received, in order. */
let sent: Array<{ content: string; trigger: string; filePath: string | null }> = []

/** What the model answers with; '' is a model that returned nothing usable. */
let answer = 'Quarterly notes'
/** What the preference reads as. First-save naming consults it; manual does not. */
let namingEnabled = true

function makeEditor(md: string): Editor {
  const editor = new Editor({ extensions: [StarterKit, SelectiveEscapeMarkdown, Redaction] })
  editor.commands.setContent(md, { contentType: 'markdown' })
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
  const editor = makeEditor('# 客户回访记录\n\n请拨打 13800138000 确认订单')
  selectText(editor, SECRET)
  editor.commands.setRedaction(LABEL)
  return editor
}

/**
 * Install the naming bridge the preload exposes. Assigned onto the real jsdom
 * window rather than replacing it: TipTap needs the DOM that lives there.
 */
function installBridge(api: {
  suggestFileName: (input: { content: string; trigger: string; filePath?: string | null }) => Promise<unknown>
  getFileNamingEnabled: () => Promise<boolean>
}): void {
  ;(window as unknown as { aiOffice: unknown }).aiOffice = api
}

beforeEach(() => {
  sent = []
  answer = 'Quarterly notes'
  namingEnabled = true
  // The latch is module state, deliberately: one attempt per document. Every
  // test starts from a freshly swapped document so the cases stay independent.
  noteDocumentSwapped()
  installBridge({
    suggestFileName: async (input) => {
      sent.push({ content: input.content, trigger: input.trigger, filePath: input.filePath ?? null })
      return answer ? { ok: true, name: answer } : { ok: false, reason: 'empty-name' }
    },
    getFileNamingEnabled: async () => namingEnabled,
  })
})

afterEach(() => {
  delete (window as unknown as { aiOffice?: unknown }).aiOffice
  vi.restoreAllMocks()
})

describe('the first-save name is derived from the projected text', () => {
  it('never hands the model a value the reader withheld', async () => {
    const editor = editorWithRedaction()
    await firstSaveName(editor)
    // The one assertion that matters. The file name is written to disk, synced
    // and emailed, so a secret that reached it would outlive the redaction
    // meant to contain it.
    expect(sent).toHaveLength(1)
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(JSON.stringify(sent[0]!.content)).not.toContain('13800138')
  })

  it('sends the marker the reader chose, so the model can still read the line', async () => {
    const editor = editorWithRedaction()
    await firstSaveName(editor)
    expect(sent[0]!.content).toContain(LABEL)
    expect(sent[0]!.content).toContain('客户回访记录')
  })

  it('sends the whole document for one with nothing withheld', async () => {
    const editor = makeEditor('# Release notes\n\nthe build is green')
    await firstSaveName(editor)
    expect(sent[0]!.content).toContain('Release notes')
    expect(sent[0]!.content).toContain('the build is green')
  })

  it('asks as a first-save attempt, which is the one the preference governs', async () => {
    const editor = makeEditor('# Anything')
    await firstSaveName(editor)
    expect(sent[0]!.trigger).toBe('first-save')
  })
})

describe('the first-save name falls back to the document own heading', () => {
  it('keeps the heading when naming is declined', async () => {
    namingEnabled = false
    const editor = makeEditor('# 客户回访记录\n\n正文')
    expect(await firstSaveName(editor)).toBe('客户回访记录')
  })

  it('keeps the heading when the model errors', async () => {
    installBridge({
      suggestFileName: async () => {
        throw new Error('provider unreachable')
      },
      getFileNamingEnabled: async () => true,
    })
    const editor = makeEditor('# 客户回访记录\n\n正文')
    expect(await firstSaveName(editor)).toBe('客户回访记录')
  })

  it('keeps the heading when the model times out with nothing to show', async () => {
    // The main process aborts a naming turn at 30s and answers ok:false, so a
    // timeout reaches the app as a decline rather than as a throw.
    installBridge({
      suggestFileName: async () => ({ ok: false, reason: 'timeout' }),
      getFileNamingEnabled: async () => true,
    })
    const editor = makeEditor('# 客户回访记录\n\n正文')
    expect(await firstSaveName(editor)).toBe('客户回访记录')
  })

  it('keeps the heading when the model returns an empty name', async () => {
    answer = ''
    const editor = makeEditor('# 客户回访记录\n\n正文')
    expect(await firstSaveName(editor)).toBe('客户回访记录')
  })

  it('keeps the heading when there is no naming bridge at all', async () => {
    delete (window as unknown as { aiOffice?: unknown }).aiOffice
    const editor = makeEditor('# 客户回访记录\n\n正文')
    expect(await firstSaveName(editor)).toBe('客户回访记录')
  })

  it('prefers the model name when one does arrive', async () => {
    const editor = makeEditor('# 客户回访记录\n\n正文')
    expect(await firstSaveName(editor)).toBe('Quarterly notes')
  })
})

describe('one naming turn per document', () => {
  it('does not ask again on a second save', async () => {
    const editor = makeEditor('# 客户回访记录\n\n正文')
    expect(await firstSaveName(editor)).toBe('Quarterly notes')
    // A save that re-asked would put a model call behind every keystroke the
    // reader made after the first one.
    expect(await firstSaveName(editor)).toBe('客户回访记录')
    expect(sent).toHaveLength(1)
  })

  it('does not start a second call while the first is still in flight', async () => {
    const editor = makeEditor('# 客户回访记录\n\n正文')
    // Two saves overlapping is what a second ⌘S during a slow turn looks like.
    const [first, second] = await Promise.all([firstSaveName(editor), firstSaveName(editor)])
    expect(sent).toHaveLength(1)
    expect(first).toBe('Quarterly notes')
    expect(second).toBe('客户回访记录')
  })

  it('asks again once a different document is loaded', async () => {
    const first = makeEditor('# 客户回访记录')
    expect(await firstSaveName(first)).toBe('Quarterly notes')
    noteDocumentSwapped()
    const second = makeEditor('# 另一个文档')
    expect(await firstSaveName(second)).toBe('Quarterly notes')
    expect(sent).toHaveLength(2)
  })

  it('still falls back to the heading on the second save of the same document', async () => {
    const editor = makeEditor('# 客户回访记录')
    await firstSaveName(editor)
    expect(await firstSaveName(editor)).toBe('客户回访记录')
  })
})

describe('the manual name goes to the Save As dialog, projected', () => {
  it('never hands the model a value the reader withheld', async () => {
    const editor = editorWithRedaction()
    await modelSaveAsName(editor, '/tmp/notes.md')
    expect(sent).toHaveLength(1)
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(sent[0]!.content).toContain(LABEL)
  })

  it('tells the main process which file is being renamed, so it can refuse a real one', async () => {
    const editor = makeEditor('# Release notes')
    await modelSaveAsName(editor, '/tmp/notes.md')
    expect(sent[0]!.filePath).toBe('/tmp/notes.md')
    expect(sent[0]!.trigger).toBe('manual')
  })

  it('goes through even with first-save naming switched off', async () => {
    // The reader asked for it by clicking, so the preference does not apply.
    namingEnabled = false
    const editor = makeEditor('# Release notes')
    expect(await modelSaveAsName(editor, '/tmp/notes.md')).toBe('Quarterly notes')
    expect(sent).toHaveLength(1)
  })

  it('returns nothing when the model declines, leaving the dialog on the current name', async () => {
    answer = ''
    const editor = makeEditor('# Release notes')
    expect(await modelSaveAsName(editor, '/tmp/notes.md')).toBe('')
  })

  it('is repeatable: a second request asks again', async () => {
    // The latch is a first-save promise, not a rate limit — a reader who
    // changed their mind and clicked again gets a fresh answer.
    const editor = makeEditor('# Release notes')
    await modelSaveAsName(editor, '/tmp/notes.md')
    await modelSaveAsName(editor, '/tmp/notes.md')
    expect(sent).toHaveLength(2)
  })
})
