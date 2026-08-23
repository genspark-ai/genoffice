import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildDocContext, buildDocumentContext, type AiHfState } from '../src/renderer/ai/protocol'
import {
  executeTool,
  type AiHeaderFooterAccess,
  type ToolExecution,
} from '../src/renderer/ai/tools'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const para = (t: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [{ type: 'text', text: t }],
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function createEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [para('Body paragraph.')] },
  })
  editors.add(editor)
  return editor
}

const NUM_IDS = { bullet: null, ordered: null }

const emptyState = (over: Partial<AiHfState> = {}): AiHfState => ({
  header: '',
  footer: '',
  headerFirst: null,
  footerFirst: null,
  headerEven: null,
  footerEven: null,
  titlePg: false,
  evenOddHf: false,
  multiSection: false,
  ...over,
})

function makeAccess(state: AiHfState): {
  access: AiHeaderFooterAccess
  writes: { kind: string; view: string; text: string }[]
} {
  const writes: { kind: string; view: string; text: string }[] = []
  return {
    access: {
      read: () => state,
      set: (kind, view, text) => {
        writes.push({ kind, view, text })
        return null
      },
    },
    writes,
  }
}

function run(editor: Editor, input: Record<string, unknown>, access?: AiHeaderFooterAccess) {
  return executeTool(
    editor,
    { id: 't1', name: 'set_header_footer', input },
    NUM_IDS,
    undefined,
    undefined,
    undefined,
    undefined,
    access,
  ) as ToolExecution
}

describe('set_header_footer tool', () => {
  it('writes through the access and reports the variant', () => {
    const editor = createEditor()
    const { access, writes } = makeAccess(emptyState())
    const exec = run(editor, { kind: 'footer', text: '{PAGE} / {NUMPAGES}' }, access)
    expect(exec.isError).toBeFalsy()
    expect(exec.mutated).toBe(false)
    expect(writes).toEqual([{ kind: 'footer', view: 'default', text: '{PAGE} / {NUMPAGES}' }])
    const first = run(editor, { kind: 'header', text: 'Cover', view: 'first' }, access)
    expect(first.output).toContain('first-page variant')
    expect(writes[1]).toEqual({ kind: 'header', view: 'first', text: 'Cover' })
  })

  it('validates input and surfaces access errors', () => {
    const editor = createEditor()
    const { access } = makeAccess(emptyState())
    expect(run(editor, { kind: 'margin', text: 'x' }, access).isError).toBe(true)
    expect(run(editor, { kind: 'header', text: 'x', view: 'odd' }, access).isError).toBe(true)
    expect(run(editor, { kind: 'header', text: 42 }, access).isError).toBe(true)
    expect(run(editor, { kind: 'header', text: 'x' }).isError).toBe(true) // no access wired
    const locked: AiHeaderFooterAccess = {
      read: () => emptyState(),
      set: () => 'the document is read-only; headers/footers cannot be edited',
    }
    const exec = run(editor, { kind: 'header', text: 'x' }, locked)
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('read-only')
  })
})

describe('header/footer context', () => {
  it('lists current texts and variants in the document context', () => {
    const editor = createEditor()
    const ctx = buildDocumentContext(
      editor,
      undefined,
      emptyState({
        header: 'Confidential',
        footer: '{PAGE} / {NUMPAGES}',
        titlePg: true,
        headerFirst: '',
        footerFirst: 'Cover footer',
        multiSection: true,
      }),
    )
    expect(ctx).toContain('- header: "Confidential" | footer: "{PAGE} / {NUMPAGES}"')
    expect(ctx).toContain('- first-page variant: header (empty) | footer "Cover footer"')
    expect(ctx).not.toContain('even-page variant')
    expect(ctx).toContain('multiple sections')
  })

  it('omits the section when no hf state is provided', () => {
    const editor = createEditor()
    expect(buildDocumentContext(editor)).not.toContain('Headers & footers')
  })

  it('keeps headers/footers visible on a blank document', () => {
    const blank = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph', attrs: { docxIndex: null } }] },
    })
    editors.add(blank)
    const ctx = buildDocContext(blank, undefined, undefined, emptyState({ header: 'Kept header' }))
    expect(ctx).toContain('The document is currently blank.')
    expect(ctx).toContain('- header: "Kept header" | footer: (empty)')
  })
})
