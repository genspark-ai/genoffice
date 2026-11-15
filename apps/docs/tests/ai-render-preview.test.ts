/**
 * render_preview: the visual self-check for agent-generated documents. The
 * guard is a pure function over the executed tool list; the executor runs
 * the print-to-PDF buffer through pdf.js and hands the page images to the
 * media-analysis pipeline. The vision model itself is stubbed — what these
 * tests pin is the wiring and the guard's triggers.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createDocsSkill } from '../src/renderer/ai/docs-skill'
import { executeTool, renderPreviewGuard } from '../src/renderer/ai/tools'
import type { NumIds } from '../src/renderer/ai/protocol'

vi.mock('pdfjs-dist/legacy/build/pdf.worker.min.mjs?url', () => ({ default: 'worker.js' }))

const fakePage = () => ({
  getViewport: () => ({ width: 100, height: 140 }),
  render: () => ({ promise: Promise.resolve() }),
  cleanup: () => {},
})

vi.mock('pdfjs-dist/legacy/build/pdf.mjs', () => ({
  GlobalWorkerOptions: { workerSrc: '' },
  getDocument: () => ({
    promise: Promise.resolve({
      numPages: 2,
      getPage: async () => fakePage(),
    }),
    destroy: async () => {},
  }),
}))

const liveEditors: Editor[] = []

function createEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: null },
          content: [{ type: 'text', text: 'Body text.' }],
        },
      ],
    } as never,
  })
  liveEditors.push(editor)
  return editor
}

afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

const numIds = {} as NumIds

const pageSetup = {
  current: () =>
    ({
      settings: { pageWidth: 12240, pageHeight: 15840 },
    }) as never,
} as never

describe('renderPreviewGuard', () => {
  const gen = (name: string, ok = true) => ({ name, ok })

  it('forces a look after a wholesale generation', () => {
    const correction = renderPreviewGuard(
      [gen('create_document'), gen('get_document_context')],
      true,
    )
    expect(correction).toContain('render_preview')
  })

  it('forces a look after a streamed rewrite', () => {
    expect(renderPreviewGuard([gen('write_document')], true)).toContain('render_preview')
  })

  it('stays silent once the run rendered and audited', () => {
    expect(renderPreviewGuard([gen('create_document'), gen('render_preview')], true)).toBeNull()
  })

  it('stays silent for small edits and failed runs', () => {
    expect(renderPreviewGuard([gen('apply_ops')], true)).toBeNull()
    expect(renderPreviewGuard([gen('create_document', false)], true)).toBeNull()
    expect(renderPreviewGuard([], true)).toBeNull()
  })

  it('disables with no vision model', () => {
    expect(renderPreviewGuard([gen('create_document')], false)).toBeNull()
  })
})

describe('the render_preview tool', () => {
  it('renders the buffer, audits through the media pipeline, returns the text', async () => {
    const printPdfBuffer = vi.fn(async () => ({ ok: true, base64: 'ZmFrZQ==' }))
    const analyzeMedia = vi.fn(async (_op: { mediaUrls: string[] }) => ({
      text: 'No layout problems found.',
    }))
    vi.stubGlobal('window', Object.assign(window, { desktop: { printPdfBuffer, analyzeMedia } }))
    vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue(
      'data:image/jpeg;base64,QUJD',
    )

    const editor = createEditor()
    const r = await executeTool(
      editor,
      { name: 'render_preview', input: {} } as never,
      numIds,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      pageSetup,
    )
    expect(r.isError).toBeFalsy()
    expect(r.output).toContain('No layout problems found.')
    expect(printPdfBuffer).toHaveBeenCalledWith(12240, 15840)
    expect(analyzeMedia).toHaveBeenCalledTimes(1)
    const op = analyzeMedia.mock.calls[0]![0] as { mediaUrls: string[]; requirements: string }
    expect(op.mediaUrls.every((u) => u.startsWith('data:image/jpeg;base64,'))).toBe(true)
    expect(op.requirements).toContain('Audit the visual layout')
  })

  it('surfaces a render failure as a tool error', async () => {
    vi.stubGlobal(
      'window',
      Object.assign(window, {
        desktop: { printPdfBuffer: vi.fn(async () => ({ ok: false, error: 'print failed' })) },
      }),
    )
    const editor = createEditor()
    const r = await executeTool(
      editor,
      { name: 'render_preview', input: {} } as never,
      numIds,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      pageSetup,
    )
    expect(r.isError).toBe(true)
    expect(r.output).toContain('print failed')
  })

  it('is hidden when no media provider is available', () => {
    const skill = createDocsSkill(
      () => createEditor(),
      () => numIds,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      () => false,
    )
    const names = skill.tools.map((t) => t.name)
    expect(names).not.toContain('render_preview')
    expect(names).not.toContain('analyze_media')
    // and the guard matches: eyes off, no correction
    expect(skill.verifyResponse?.('done', [{ name: 'create_document', ok: true }])).toBeNull()
  })

  it('the skill guard fires after a generation without a render', () => {
    const skill = createDocsSkill(
      () => createEditor(),
      () => numIds,
    )
    expect(
      skill.verifyResponse?.('The document is ready.', [{ name: 'create_document', ok: true }]),
    )!.toContain('render_preview')
    expect(
      skill.verifyResponse?.('Ready.', [
        { name: 'create_document', ok: true },
        { name: 'render_preview', ok: true },
      ]),
    ).toBeNull()
  })
})
