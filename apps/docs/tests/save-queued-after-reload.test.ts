/**
 * A save queued behind another one. The first save reloads the editor from the
 * bytes it wrote, so every docxIndex anchor then points into the re-parsed
 * document, while the app's context still holds the old one until the next
 * render. The queued save has to wait for that context: built from the old
 * blocks and the new anchors it writes the wrong paragraphs.
 */
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import { parseDocx, readSections } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocOptions } from '../src/renderer/editor/convert'
import {
  buildDocBytes,
  noteDocumentSwapped,
  save,
  type FileActionContext,
} from '../src/renderer/file-actions'
import { EMPTY_HF_VARIANTS } from '../src/renderer/doc-state'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const dirs: string[] = []
const editors: Editor[] = []
afterEach(async () => {
  for (const editor of editors.splice(0)) editor.destroy()
  for (const dir of dirs.splice(0)) await rm(dir, { recursive: true, force: true })
})

const LINES = ['line 01', 'line 02', 'line 03', 'line 04', 'line 05']
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>'

/**
 * The app's side of a save: state the save's setters change at once, and a
 * context that is a snapshot of it, renewed only by `render` — as the app's is
 * by a React render, some time after the setters ran.
 */
async function fixture() {
  noteDocumentSwapped()
  const bytes = await buildDocx({
    bodyXml: LINES.map((text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`).join(''),
  })
  const parsed = await parseDocx(bytes)
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-save-queued-'))
  dirs.push(dir)
  const path = join(dir, 'fixture.docx')
  await writeFile(path, bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    editorProps: { handleScrollToSelection: () => true },
    content: blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed)),
  })
  editors.push(editor)
  const dirtyRef = { current: false }
  let state = {
    editor,
    doc: { parsed, filePath: path, fileName: 'fixture.docx', hash: 'fixture' },
    dirtyRef,
    saveInFlightRef: { current: false },
    saveIncompleteRef: { current: false },
    inkAnnotations: [],
    sections: readSections(parsed),
    sectionsDirty: [],
    pgNumDirtySections: [],
    hfVariants: EMPTY_HF_VARIANTS,
    hfVariantsDirty: [],
    sectionHfEdits: {},
    hfLinks: {},
    styleUpserts: {},
    pendingNumbering: { newDefs: [], restartNums: [] },
  }
  type State = typeof state
  let rendered: FileActionContext
  const noop = () => {}
  const snapshot = (of: State): FileActionContext =>
    new Proxy(of, {
      get(target, prop) {
        if (prop === 'setDoc')
          return (updater: (prev: State['doc']) => State['doc']) => {
            state = { ...state, doc: updater(state.doc) }
          }
        if (prop === 'setSections')
          return (sections: State['sections']) => {
            state = { ...state, sections }
          }
        if (prop === 'settleFontSettings') return async () => rendered
        if (prop in target) return Reflect.get(target, prop)
        if (prop === 'endManualHyphenation') return noop
        if (typeof prop === 'string' && (/^set[A-Z]/.test(prop) || /^on[A-Z]/.test(prop)))
          return noop
        return undefined
      },
    }) as unknown as FileActionContext
  const render = () => {
    rendered = snapshot(state)
    return rendered
  }
  render()
  /** run once, right after the next write: what goes wrong between the write and the reload */
  let afterWrite: (() => void) | null = null
  window.desktop = {
    saveDocx: async (target: string, data: ArrayBuffer) => {
      await writeFile(target, new Uint8Array(data))
      afterWrite?.()
      afterWrite = null
      return { ok: true }
    },
  } as unknown as typeof window.desktop
  /** the reload's `failing`-th editor command throws: the first rewrites the document, the second puts the caret back */
  const failReloadAt = (failing: number) => {
    afterWrite = () => {
      const chain = editor.chain.bind(editor)
      let calls = 0
      editor.chain = () => {
        if (++calls < failing) return chain()
        editor.chain = chain
        throw new Error('reload failed')
      }
    }
  }

  /** a section break after the paragraph at `index`, as the Layout tab inserts it */
  const insertBreakAfter = (index: number) => {
    let pos = 0
    for (let i = 0; i <= index; i++) pos += editor.state.doc.child(i).nodeSize
    editor.commands.insertContentAt(pos, {
      type: 'docProtected',
      attrs: {
        blockType: 'passthrough',
        label: 'Section break paragraph',
        previewText: '',
        genXml: `<w:p><w:pPr>${SECT_PR}</w:pPr></w:p>`,
      },
    })
    dirtyRef.current = true
  }
  const type = (text: string) => {
    editor.view.dispatch(editor.state.tr.insertText(text, 1))
    dirtyRef.current = true
  }
  const read = async (bytes: Uint8Array) => {
    const zip = await JSZip.loadAsync(bytes)
    const xml = await zip.file('word/document.xml')!.async('string')
    return {
      sections: (xml.match(/<w:sectPr\b/g) ?? []).length,
      texts: [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]),
    }
  }
  const saved = async () => read(await readFile(path))
  return { editor, render, insertBreakAfter, type, saved, read, failReloadAt }
}

const nextTask = () => new Promise((resolve) => setTimeout(resolve, 0))

describe('a save queued behind a save that reloads the document', () => {
  it('does not write from the context captured before that reload', async () => {
    const { editor, render, insertBreakAfter, saved } = await fixture()
    insertBreakAfter(1)
    // both asked for before any render: the automatic save a section break
    // brings, and the user's own
    const ctx = render()
    const first = save(ctx, false, true)
    const second = save(ctx, false, false)
    expect(await first).toBe(true)
    // the editor is on the re-parsed document; the context follows a task later
    await nextTask()
    render()
    expect(await second).toBe(true)
    expect(await saved()).toEqual({ sections: 2, texts: LINES })
    expect(editor.state.doc.childCount).toBe(6)
  })

  it('still writes what was typed between the reload and its own turn', async () => {
    const { render, insertBreakAfter, type, saved } = await fixture()
    insertBreakAfter(1)
    const ctx = render()
    const first = save(ctx, false, true)
    const second = save(ctx, false, false)
    expect(await first).toBe(true)
    type('Typed ')
    await nextTask()
    render()
    expect(await second).toBe(true)
    expect(await saved()).toEqual({
      sections: 2,
      texts: ['Typed line 01', ...LINES.slice(1)],
    })
  })

  it('still writes what was typed right after both saves were asked for', async () => {
    const { render, insertBreakAfter, type, saved } = await fixture()
    insertBreakAfter(1)
    const ctx = render()
    const first = save(ctx, false, true)
    const second = save(ctx, false, false)
    type('Typed ')
    expect(await first).toBe(true)
    await nextTask()
    render()
    expect(await second).toBe(true)
    expect(await saved()).toEqual({
      sections: 2,
      texts: ['Typed line 01', ...LINES.slice(1)],
    })
  })

  it('writes nothing when no context for the reloaded document ever arrives', async () => {
    const { render, insertBreakAfter, type, saved } = await fixture()
    insertBreakAfter(1)
    const ctx = render()
    const first = save(ctx, false, true)
    const second = save(ctx, false, false)
    expect(await first).toBe(true)
    type('Typed ')
    expect(await second).toBe(false)
    expect(await saved()).toEqual({ sections: 2, texts: LINES })
  }, 20_000)
})

describe('a reload that fails part-way', () => {
  it('before the editor was rewritten: the next save works from the document as it was', async () => {
    const { editor, render, insertBreakAfter, type, saved, failReloadAt } = await fixture()
    insertBreakAfter(1)
    const before = editor.state.doc
    failReloadAt(1)
    expect(await save(render(), false, true)).toBe(false)
    expect(editor.state.doc).toBe(before)
    type('Typed ')
    await nextTask()
    expect(await save(render(), false, false)).toBe(true)
    expect(await saved()).toEqual({ sections: 2, texts: ['Typed line 01', ...LINES.slice(1)] })
  })

  it('after the editor was rewritten: the next save works from the re-parsed document', async () => {
    const { editor, render, insertBreakAfter, type, saved, failReloadAt } = await fixture()
    insertBreakAfter(1)
    const before = editor.state.doc
    failReloadAt(2)
    expect(await save(render(), false, true)).toBe(false)
    expect(editor.state.doc).not.toBe(before)
    type('Typed ')
    await nextTask()
    expect(await save(render(), false, false)).toBe(true)
    expect(await saved()).toEqual({ sections: 2, texts: ['Typed line 01', ...LINES.slice(1)] })
  })
})

describe("a document of the caller's making, serialized across a reload", () => {
  it('is made from the document as the reload left it, not from an earlier snapshot', async () => {
    const { editor, render, insertBreakAfter, read } = await fixture()
    insertBreakAfter(1)
    const ctx = render()
    const early = editor.getJSON()
    expect(await save(ctx, false, true)).toBe(true)
    // asked for with the context from before the reload; the render comes while it waits
    let handed: unknown
    const pending = buildDocBytes(ctx, (live) => {
      handed = live
      const [firstPara, ...rest] = live.content ?? []
      return {
        ...live,
        content: [{ ...firstPara, content: [{ type: 'text', text: 'merged 01' }] }, ...rest],
      }
    })
    await nextTask()
    render()
    const bytes = await pending
    expect(handed).toEqual(editor.getJSON())
    expect(handed).not.toEqual(early)
    expect(await read(bytes!)).toEqual({ sections: 2, texts: ['merged 01', ...LINES.slice(1)] })
  })
})
