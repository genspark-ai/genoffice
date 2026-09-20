import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeOps } from '../src/renderer/ai/ops'
import { setModuleLang } from '../src/renderer/i18n/locale'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const boldText = (t: string): JsonNode => ({ type: 'text', text: t, marks: [{ type: 'bold' }] })

const para = (content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.add(editor)
  return editor
}

setModuleLang('en')

describe('no-op command summary', () => {
  it('reports blocks that matched but already had the requested state', () => {
    const editor = createEditor([para([boldText('Already bold')]), para([boldText('Also bold')])])
    const outcome = executeOps(editor, [
      { op: 'setFont', target: { nodeType: 'docParagraph' }, bold: true },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(outcome.results[0]).toMatchObject({ matched: 2, changed: 0, skippedProtected: 0 })
    expect(outcome.summary).toContain('2 matching block(s) already had the requested state')
    expect(outcome.summary).not.toContain('No matching')
  })

  it('still reports no matching blocks when the selector hits nothing', () => {
    const editor = createEditor([para([boldText('Already bold')])])
    const outcome = executeOps(editor, [
      { op: 'setFont', target: { containsText: 'nowhere in this document' }, bold: true },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(outcome.results[0]).toMatchObject({ matched: 0, changed: 0 })
    expect(outcome.summary).toBe('No matching blocks; the document was not changed.')
  })
})
