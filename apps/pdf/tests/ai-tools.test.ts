import { describe, expect, it, vi } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { AGENT_TOOLS, executePdfTool, type PdfAiDeps } from '../src/renderer/ai/tools'
import type { SearchIndex } from '../src/renderer/search'
import type { FormValueInput, PageImageRef } from '../src/shared/ipc'
import {
  blockMoveInput,
  editCarriesBlock,
  isBlockEditOf,
  patchPendingEdits,
  resolveTextEdit,
  shiftPendingEdit,
} from '../src/renderer/text-edit-preview'
import type { LocalTextEdit, LocalTextInsert } from '../src/renderer/text-edit-preview'
import type { TextBlock } from '../src/renderer/text-block'
import type { TextEditInput } from '../src/shared/ipc'
import { measurePt } from '../src/renderer/text-wrap'

/** Two pages: "Hello World" and "foo bar foo" */
const INDEX: SearchIndex = [
  {
    text: 'Hello World',
    lower: 'hello world',
    items: [{ start: 0, end: 11, x: 0, y: 700, w: 110, h: 12 }],
  },
  {
    text: 'foo bar foo',
    lower: 'foo bar foo',
    items: [{ start: 0, end: 11, x: 0, y: 700, w: 110, h: 12 }],
  },
]

interface Widget {
  id?: string
  subtype?: string
  fieldType?: string
  fieldName?: string
  fieldValue?: unknown
  buttonValue?: string
  rect?: number[]
  readOnly?: boolean
  checkBox?: boolean
  radioButton?: boolean
  options?: { exportValue?: unknown; displayValue?: unknown }[]
}

function fakeDoc(pageTexts: string[], annotsByPage: Widget[][] = []): PDFDocumentProxy {
  return {
    numPages: pageTexts.length,
    getPage: async (n: number) => ({
      getTextContent: async () => ({ items: [{ str: pageTexts[n - 1], hasEOL: true }] }),
      getAnnotations: async () =>
        (annotsByPage[n - 1] ?? []).map((widget, index) =>
          widget.subtype === 'Widget'
            ? { id: `${n}-${index}`, rect: [0, 0, 100, 20], ...widget }
            : widget,
        ),
      cleanup: () => {},
    }),
  } as unknown as PDFDocumentProxy
}

/** Page 1 (600 × 800 pt) holds two images: top-left-ish (below text) and lower-right (above text) */
const PAGE_IMAGES: PageImageRef[] = [
  { pageIndex: 0, rect: [200, 100, 300, 200], aboveText: true },
  { pageIndex: 0, rect: [50, 600, 150, 700], aboveText: false },
]

function makeDeps(over: Partial<PdfAiDeps> = {}): PdfAiDeps {
  return {
    doc: () => fakeDoc(['Hello World', 'foo bar foo']),
    fileName: () => 'test.pdf',
    pageCount: () => 2,
    currentPage: () => 1,
    readOnly: () => false,
    ocrText: () => null,
    selection: () => null,
    pendingSummary: () => '',
    annotationSummary: () => '',
    annotationsOn: vi.fn(async () => ({ threads: [], markups: [] })),
    addNote: vi.fn(() => 'Pn1'),
    findNoteRoot: vi.fn(async () => null),
    replyToThread: vi.fn(),
    editNote: vi.fn(),
    deleteMarkups: vi.fn(async () => {}),
    deleteNoteThread: vi.fn(),
    outline: () => null,
    searchIndex: () => Promise.resolve(INDEX),
    isDeleted: () => false,
    gotoPage: vi.fn(() => true),
    addMarkup: vi.fn(),
    formEdits: () => new Map<string, FormValueInput>(),
    applyFormEdit: vi.fn(),
    rotatePages: vi.fn(),
    deletePage: vi.fn(() => true),
    metadata: () => ({ title: 'Report', author: 'Ann', subject: '', keywords: '' }),
    setMetadata: vi.fn(),
    pageOrder: () => [0, 1],
    movePage: vi.fn(),
    reversePages: vi.fn(),
    editText: vi.fn(async () => null),
    moveTextBlock: vi.fn(async (_idx, _block, d: [number, number]) => ({ moveBy: d })),
    addFormMark: vi.fn(),
    insertText: vi.fn(() => 'i1'),
    textInserts: () => [],
    updateTextInsert: vi.fn(),
    moveTextInsert: vi.fn(),
    deleteTextInsert: vi.fn(),
    editFonts: () => ['arial', 'times', 'courier'],
    pageGeom: () => ({ pw: 600, ph: 800, rot: 0 }),
    listImages: vi.fn(async () => PAGE_IMAGES),
    isImageClaimed: () => false,
    insertImage: vi.fn(),
    transformImage: vi.fn(),
    replaceImage: vi.fn(),
    bakeImage: vi.fn(async () => true),
    deleteImage: vi.fn(),
    searchImages: vi.fn(async () => ({
      images: [
        {
          title: 'A cat',
          imageUrl: 'https://img.example/cat.jpg',
          sourceUrl: 'https://example.com',
          source: 'example',
          width: 800,
          height: 600,
        },
      ],
      method: 'gsk',
    })),
    generateImage: vi.fn(async () => ({ url: 'https://img.example/generated.png' })),
    fetchImage: vi.fn(async () => ({ png: 'PNGB64', width: 400, height: 300 })),
    createDocument: vi.fn(async () => ({ ok: true, path: '/tmp/out.pdf' })),
    stamps: () => null,
    setStamps: vi.fn(),
    confirmFileOp: vi.fn(async () => true),
    insertBlankPage: vi.fn(async () => ({ ok: true, pageCount: 3 })),
    setPageSize: vi.fn(async () => ({ ok: true, pageCount: 2 })),
    cropPages: vi.fn(async () => ({ ok: true, pageCount: 2 })),
    replacePages: vi.fn(async () => ({ ok: true, pageCount: 4 })),
    extractPages: vi.fn(async () => ({ ok: true, savedPath: '/tmp/test-p1.pdf' })),
    splitPdf: vi.fn(async () => ({ ok: true, savedDir: '/tmp/split', count: 2 })),
    splitPages: vi.fn(async () => ({ ok: true, savedPath: '/tmp/test-split.pdf' })),
    mergePages: vi.fn(async () => ({ ok: true, savedPath: '/tmp/test-2in1.pdf' })),
    ...over,
  }
}

const call = (name: string, input: Record<string, unknown> = {}) => ({ id: 't1', name, input })

describe('AGENT_TOOLS definitions', () => {
  it('declares unique names and object input schemas with required fields present', () => {
    const names = AGENT_TOOLS.map((t) => t.name)
    expect(new Set(names).size).toBe(names.length)
    for (const tool of AGENT_TOOLS) {
      expect(tool.description ?? tool.name).toBeTruthy()
      expect(tool.inputSchema.type).toBe('object')
      const props = tool.inputSchema.properties as Record<string, unknown>
      for (const req of (tool.inputSchema.required as string[] | undefined) ?? []) {
        expect(props).toHaveProperty(req)
      }
    }
  })

  it('every declared tool is handled by executePdfTool', async () => {
    // A tool falling into the default branch would return "Unknown tool"
    const deps = makeDeps({ doc: () => null, searchIndex: () => null })
    for (const tool of AGENT_TOOLS) {
      const result = await executePdfTool(deps, call(tool.name, { page: 1, start: 1 }))
      expect(result.output).not.toContain('Unknown tool')
    }
  })
})

describe('read_pages', () => {
  it('reads a page range with [Page N] markers', async () => {
    const result = await executePdfTool(makeDeps(), call('read_pages', { start: 1, end: 2 }))
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('[Page 1]')
    expect(result.output).toContain('Hello World')
    expect(result.output).toContain('[Page 2]')
    expect(result.output).toContain('foo bar foo')
  })

  it('rejects an out-of-range start page', async () => {
    const result = await executePdfTool(makeDeps(), call('read_pages', { start: 5 }))
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Invalid page range')
  })

  it('errors when the document is not loaded', async () => {
    const result = await executePdfTool(
      makeDeps({ doc: () => null }),
      call('read_pages', { start: 1 }),
    )
    expect(result.isError).toBe(true)
  })
})

describe('search_text', () => {
  it('returns page numbers with context excerpts', async () => {
    const result = await executePdfTool(makeDeps(), call('search_text', { query: 'FOO' }))
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('Page 2')
    expect(result.output).toContain('foo bar foo')
  })

  it('rejects an empty query', async () => {
    const result = await executePdfTool(makeDeps(), call('search_text', { query: '  ' }))
    expect(result.isError).toBe(true)
  })

  it('reports no matches without erroring', async () => {
    const result = await executePdfTool(makeDeps(), call('search_text', { query: 'zzz' }))
    expect(result.output).toBe('No matches found')
  })
})

describe('goto_page', () => {
  it('scrolls to a valid page', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(deps, call('goto_page', { page: 2 }))
    expect(result.isError).toBeUndefined()
    expect(deps.gotoPage).toHaveBeenCalledWith(2)
  })

  it('rejects out-of-range and deleted pages', async () => {
    const bad = await executePdfTool(makeDeps(), call('goto_page', { page: 3 }))
    expect(bad.isError).toBe(true)
    expect(bad.output).toContain('out of range')

    const deleted = await executePdfTool(
      makeDeps({ isDeleted: (i) => i === 0 }),
      call('goto_page', { page: 1 }),
    )
    expect(deleted.isError).toBe(true)
    expect(deleted.output).toContain('deleted')
  })
})

describe('markup_text', () => {
  it('marks the first occurrence and jumps to the page', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('markup_text', { page: 2, text: 'foo', type: 'highlight' }),
    )
    expect(result.mutated).toBe(true)
    expect(deps.addMarkup).toHaveBeenCalledTimes(1)
    expect(deps.addMarkup).toHaveBeenCalledWith('highlight', 1, expect.any(Array), undefined)
    expect(deps.gotoPage).toHaveBeenCalledWith(2)
  })

  it('marks every occurrence with all=true', async () => {
    const deps = makeDeps()
    await executePdfTool(
      deps,
      call('markup_text', { page: 2, text: 'foo', type: 'underline', all: true }),
    )
    expect(deps.addMarkup).toHaveBeenCalledTimes(2)
  })

  it('rejects text not present on the target page', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('markup_text', { page: 1, text: 'foo', type: 'highlight' }),
    )
    expect(result.isError).toBe(true)
    expect(deps.addMarkup).not.toHaveBeenCalled()
  })

  it('rejects invalid types and read-only documents', async () => {
    const badType = await executePdfTool(
      makeDeps(),
      call('markup_text', { page: 1, text: 'Hello', type: 'wavy' }),
    )
    expect(badType.isError).toBe(true)

    const ro = await executePdfTool(
      makeDeps({ readOnly: () => true }),
      call('markup_text', { page: 1, text: 'Hello', type: 'highlight' }),
    )
    expect(ro.isError).toBe(true)
    expect(ro.output).toContain('read-only')
  })
})

describe('edit_text', () => {
  it('queues an edit for the first occurrence with its located rect and font size', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('edit_text', { page: 2, old_text: 'bar', new_text: 'baz' }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(deps.editText).toHaveBeenCalledWith({
      pageIndex: 1,
      // 'bar' is chars 4-7 of the 11-char item spanning x 0-110
      rect: [40, 700, 70, 712],
      oldText: 'bar',
      newText: 'baz',
      fontSize: 12,
      newFontSize: undefined,
      newColor: undefined,
    })
    expect(deps.gotoPage).toHaveBeenCalledWith(2)
  })

  it('targets the nth occurrence and passes style overrides through', async () => {
    const deps = makeDeps()
    await executePdfTool(
      deps,
      call('edit_text', {
        page: 2,
        old_text: 'FOO',
        new_text: 'qux',
        occurrence: 2,
        font_size: 9,
        color: '#ff8000',
      }),
    )
    expect(deps.editText).toHaveBeenCalledWith(
      expect.objectContaining({
        rect: [80, 700, 110, 712],
        oldText: 'foo', // verbatim page text, not the model's casing
        newFontSize: 9,
        newColor: [255, 128, 0],
      }),
    )
  })

  it('passes a valid font through and rejects unavailable ones', async () => {
    const deps = makeDeps()
    await executePdfTool(
      deps,
      call('edit_text', { page: 1, old_text: 'Hello', new_text: 'Hi', font: 'times' }),
    )
    expect(deps.editText).toHaveBeenCalledWith(expect.objectContaining({ newFont: 'times' }))

    const bad = await executePdfTool(
      makeDeps({ editFonts: () => ['arial'] }),
      call('edit_text', { page: 1, old_text: 'Hello', new_text: 'Hi', font: 'times' }),
    )
    expect(bad.isError).toBe(true)
    expect(bad.output).toContain('arial')
  })

  it('notes remaining occurrences when occurrence is omitted', async () => {
    const result = await executePdfTool(
      makeDeps(),
      call('edit_text', { page: 2, old_text: 'foo', new_text: 'baz' }),
    )
    expect(result.output).toContain('2 occurrences')
  })

  it('rejects text not present on the page without queueing', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('edit_text', { page: 1, old_text: 'foo', new_text: 'baz' }),
    )
    expect(result.isError).toBe(true)
    expect(deps.editText).not.toHaveBeenCalled()
  })

  it('deletes the run when new_text is empty or whitespace-only', async () => {
    for (const newText of ['', '  ']) {
      const deps = makeDeps()
      const result = await executePdfTool(
        deps,
        call('edit_text', { page: 1, old_text: 'Hello', new_text: newText }),
      )
      expect(result.isError).toBeUndefined()
      expect(result.mutated).toBe(true)
      expect(result.output).toContain('Deleted')
      expect(deps.editText).toHaveBeenCalledWith(expect.objectContaining({ newText: '' }))
    }
  })

  it('rejects bad colors and read-only documents', async () => {
    const badColor = await executePdfTool(
      makeDeps(),
      call('edit_text', { page: 1, old_text: 'Hello', new_text: 'Hi', color: 'red' }),
    )
    expect(badColor.isError).toBe(true)

    const ro = await executePdfTool(
      makeDeps({ readOnly: () => true }),
      call('edit_text', { page: 1, old_text: 'Hello', new_text: 'Hi' }),
    )
    expect(ro.isError).toBe(true)
    expect(ro.output).toContain('read-only')
  })

  it('surfaces the rejection reason when the edit fails validation', async () => {
    const deps = makeDeps({ editText: vi.fn(async () => 'no match') })
    const result = await executePdfTool(
      deps,
      call('edit_text', { page: 1, old_text: 'Hello', new_text: 'Hi' }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('no match')
  })
})

describe('edit_block', () => {
  it('deletes the whole paragraph when new_text is empty', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('edit_block', { page: 1, paragraph_text: 'Hello', new_text: '' }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('Deleted the paragraph')
    expect(deps.editText).toHaveBeenCalledWith(
      expect.objectContaining({ pageIndex: 0, oldText: 'Hello World', newText: '' }),
    )
  })

  it('keeps the detected alignment when align is omitted', async () => {
    const deps = makeDeps()
    await executePdfTool(
      deps,
      call('edit_block', { page: 1, paragraph_text: 'Hello', new_text: 'Hi' }),
    )
    const input = vi.mocked(deps.editText).mock.calls[0]![0]
    expect(input.align).toBeUndefined()
    expect(input.lineXOffsets).toBeUndefined()
  })

  it('re-aligns the reflowed lines for align=center / right', async () => {
    const deps = makeDeps()
    await executePdfTool(
      deps,
      call('edit_block', { page: 1, paragraph_text: 'Hello', new_text: 'Hi', align: 'center' }),
    )
    const centered = vi.mocked(deps.editText).mock.calls[0]![0]
    expect(centered.align).toBe('center')
    expect(centered.lineXOffsets).toHaveLength(1)
    expect(centered.lineXOffsets![0]).toBeGreaterThanOrEqual(0)

    await executePdfTool(
      deps,
      call('edit_block', { page: 1, paragraph_text: 'Hello', new_text: 'Hi', align: 'right' }),
    )
    const right = vi.mocked(deps.editText).mock.calls[1]![0]
    expect(right.align).toBe('right')
    expect(right.lineXOffsets![0]).toBeGreaterThanOrEqual(centered.lineXOffsets![0]!)
  })

  it('rejects an unknown align value', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('edit_block', { page: 1, paragraph_text: 'Hello', new_text: 'Hi', align: 'justify' }),
    )
    expect(result.isError).toBe(true)
    expect(deps.editText).not.toHaveBeenCalled()
  })
})

describe('edit_text align', () => {
  it('points the model at edit_block instead of silently dropping align', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('edit_text', { page: 1, old_text: 'Hello', new_text: 'Hi', align: 'center' }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('edit_block')
    expect(deps.editText).not.toHaveBeenCalled()
  })
})

describe('move_text_block', () => {
  it('translates the located paragraph by the display offset and reports the new position', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 10, dy: 20 }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    // screen-down is PDF-up: dy 20 → -20 in user space
    expect(deps.moveTextBlock).toHaveBeenCalledWith(
      0,
      expect.objectContaining({ rect: [0, 697.6, 110, 712] }),
      [10, -20],
    )
    // block top-left as displayed: x 0+10, y (800-712)+20
    expect(result.output).toContain('x 10, y 108')
    expect(result.output).toContain('unsaved')
    expect(deps.gotoPage).toHaveBeenCalledWith(1)
  })

  it('reports the position from the total pending displacement when stacking on a pending edit', async () => {
    const deps = makeDeps({ moveTextBlock: vi.fn(async () => ({ moveBy: [10, -50] })) })
    const result = await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 10, dy: 20 }),
    )
    expect(deps.moveTextBlock).toHaveBeenCalledWith(0, expect.anything(), [10, -20])
    expect(result.output).toContain('by dx 10, dy 20 pt')
    expect(result.output).toContain('x 10, y 138')
  })

  it('maps the offset through the page rotation', async () => {
    const deps = makeDeps({ pageGeom: () => ({ pw: 600, ph: 800, rot: 90 }) })
    await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 10, dy: 20 }),
    )
    expect(deps.moveTextBlock).toHaveBeenCalledWith(0, expect.anything(), [20, 10])
  })

  it('rejects a zero offset, a missing paragraph, and read-only documents', async () => {
    const deps = makeDeps()
    const zero = await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 0, dy: 0 }),
    )
    expect(zero.isError).toBe(true)
    const missing = await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Nope', dx: 0, dy: -30 }),
    )
    expect(missing.isError).toBe(true)
    expect(missing.output).toContain('No paragraph on page 1')
    expect(deps.moveTextBlock).not.toHaveBeenCalled()
    const ro = await executePdfTool(
      makeDeps({ readOnly: () => true }),
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 0, dy: -30 }),
    )
    expect(ro.isError).toBe(true)
  })

  it('composes with a same-turn edit_block on the same paragraph into one pending edit', async () => {
    // Stateful deps mirroring App: pending edits live in one list the tools mutate in turn
    const pending: LocalTextEdit[] = []
    const deps = makeDeps({
      moveTextBlock: vi.fn(async (idx, block, d: [number, number]) => {
        pending.push({ id: 'mv', input: blockMoveInput(idx, block, d), moveBy: d })
        return { moveBy: d }
      }),
      editText: vi.fn(async (raw) => {
        const r = resolveTextEdit(pending, raw)
        if ('reason' in r) return r.reason
        const gone = new Set(r.replaces.map((e) => e.id))
        pending.splice(0, pending.length, ...pending.filter((e) => !gone.has(e.id)), {
          id: `ed${pending.length}`,
          input: r.input,
          moveBy: r.moveBy,
        })
        return null
      }),
    })
    const moved = await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 0, dy: -30 }),
    )
    expect(moved.isError).toBeUndefined()
    const edited = await executePdfTool(
      deps,
      call('edit_block', { page: 1, paragraph_text: 'Hello', new_text: 'Hi', align: 'center' }),
    )
    expect(edited.isError).toBeUndefined()
    expect(pending).toHaveLength(1)
    const only = pending[0]!
    expect(only.moveBy).toEqual([0, 30])
    expect(only.input.newText).toBe('Hi')
    expect(only.input.align).toBe('center')
    expect(only.input.translate).toBeUndefined()
    // rebuild anchored at the moved corner: block left edge, first baseline + 30
    expect(only.input.origin).toEqual([0, 730])

    const line = await executePdfTool(
      deps,
      call('edit_text', { page: 1, old_text: 'Hello', new_text: 'Hey' }),
    )
    expect(line.isError).toBe(true)
    expect(line.output).toContain('edit_block')
    expect(pending).toHaveLength(1)

    // a second rewrite supersedes the first and keeps the paragraph's displacement
    const again = await executePdfTool(
      deps,
      call('edit_block', { page: 1, paragraph_text: 'Hello', new_text: 'Hello again' }),
    )
    expect(again.isError).toBeUndefined()
    expect(pending).toHaveLength(1)
    expect(pending[0]!.input.newText).toBe('Hello again')
    expect(pending[0]!.input.origin).toEqual([0, 730])
    expect(pending[0]!.moveBy).toEqual([0, 30])
  })

  it('edit_block supersedes a moved line edit of the same paragraph', async () => {
    const pending: LocalTextEdit[] = []
    const deps = makeDeps({
      editText: vi.fn(async (raw) => {
        const r = resolveTextEdit(pending, raw)
        if ('reason' in r) return r.reason
        const gone = new Set(r.replaces.map((e) => e.id))
        pending.splice(0, pending.length, ...pending.filter((e) => !gone.has(e.id)), {
          id: `ed${pending.length}`,
          input: r.input,
          moveBy: r.moveBy,
        })
        return null
      }),
    })
    const line = await executePdfTool(
      deps,
      call('edit_text', { page: 1, old_text: 'Hello', new_text: 'Hey' }),
    )
    expect(line.isError).toBeUndefined()
    // the manual drag shifts a pending line edit like this (it gains an origin)
    pending[0] = {
      ...pending[0]!,
      input: { ...pending[0]!.input, origin: [0, 720], lineLeading: 14.4 },
      moveBy: [0, 20],
    }
    const block = await executePdfTool(
      deps,
      call('edit_block', { page: 1, paragraph_text: 'Hello', new_text: 'Hi' }),
    )
    expect(block.isError).toBeUndefined()
    expect(pending).toHaveLength(1)
    expect(pending[0]!.input.newText).toBe('Hi')
    expect(pending[0]!.input.origin).toEqual([0, 720])
    expect(pending[0]!.moveBy).toEqual([0, 20])
  })

  /** Deps mirroring App's moveTextBlock: shift a pending edit that claims the block,
      else queue a pure move; dry-run everything except a shifted block-level edit, then
      patch the LIVE list by id (the dry-run may have let other edits land meanwhile) */
  const movingDeps = (
    pending: LocalTextEdit[],
    validate: (input: TextEditInput) => string | null | Promise<string | null>,
  ) =>
    makeDeps({
      moveTextBlock: vi.fn(async (idx: number, block: TextBlock, d: [number, number]) => {
        const claimed = pending.find((e) => e.input.pageIndex === idx)
        const te = claimed
          ? shiftPendingEdit(claimed, block, d)
          : { id: 'mv', input: blockMoveInput(idx, block, d), moveBy: d }
        if (!editCarriesBlock(te, block)) {
          return {
            reason:
              'a pending edit of a line inside this paragraph keeps it from moving as one unit',
          }
        }
        if (!(claimed && isBlockEditOf(te, block))) {
          const reason = await validate(te.input)
          if (reason) return { reason }
        }
        if (claimed && !pending.some((e) => e.id === te.id)) {
          return { reason: 'the pending edits changed while the move was validated; retry' }
        }
        pending.splice(
          0,
          pending.length,
          ...patchPendingEdits(pending, te, new Set(), claimed ? 'replace' : 'upsert'),
        )
        return { moveBy: te.moveBy ?? d }
      }),
    })

  it('keeps an edit queued during the dry-run await and still lands the move', async () => {
    const pending: LocalTextEdit[] = []
    const other: LocalTextEdit = {
      id: 'other',
      input: {
        pageIndex: 1,
        rect: [0, 700, 110, 712],
        oldText: 'foo bar foo',
        newText: 'x',
        fontSize: 12,
      },
    }
    const validate = vi.fn(async () => {
      pending.push(other)
      return null
    })
    const deps = movingDeps(pending, validate)
    const result = await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 0, dy: -30 }),
    )
    expect(result.isError).toBeUndefined()
    expect(pending.map((e) => e.id)).toEqual(['other', 'mv'])
    expect(pending[1]!.input.translate).toEqual([0, 30])
  })

  it('refuses when the owner it planned to shift vanished during validation', async () => {
    const pending: LocalTextEdit[] = [
      {
        id: 'line',
        input: {
          pageIndex: 0,
          rect: [0, 700, 110, 712],
          oldText: 'Hello World',
          newText: 'Hey World',
          fontSize: 12,
        },
      },
    ]
    const validate = vi.fn(async () => {
      pending.splice(0, pending.length)
      return null
    })
    const deps = movingDeps(pending, validate)
    const result = await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 0, dy: -30 }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('changed while')
    expect(result.output).not.toContain('Moved')
    expect(pending).toHaveLength(0)
  })

  it('does not resurrect an owner a background dry-run dropped even if the guard is bypassed', async () => {
    // Deps whose live list drops the owner during the await and that land without the
    // presence guard: replace mode must still keep the dropped edit out
    const pending: LocalTextEdit[] = [
      {
        id: 'line',
        input: {
          pageIndex: 0,
          rect: [0, 700, 110, 712],
          oldText: 'Hello World',
          newText: 'Hey World',
          fontSize: 12,
        },
      },
    ]
    const deps = makeDeps({
      moveTextBlock: vi.fn(async (_idx: number, block: TextBlock, d: [number, number]) => {
        const te = shiftPendingEdit(pending[0]!, block, d)
        await Promise.resolve().then(() => pending.splice(0, pending.length))
        const next = patchPendingEdits(pending, te, new Set(), 'replace')
        if (next === pending)
          return { reason: 'the pending edits changed while the move was validated; retry' }
        pending.splice(0, pending.length, ...next)
        return { moveBy: te.moveBy ?? d }
      }),
    })
    const result = await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 0, dy: -30 }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).not.toContain('Moved')
    expect(pending).toHaveLength(0)
  })

  it('dry-runs the shifted edit when only a line edit claims the paragraph', async () => {
    const pending: LocalTextEdit[] = [
      {
        id: 'line',
        input: {
          pageIndex: 0,
          rect: [0, 700, 110, 712],
          oldText: 'Hello World',
          newText: 'Hey World',
          fontSize: 12,
        },
      },
    ]
    const validate = vi.fn(() => null)
    const deps = movingDeps(pending, validate)
    const result = await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 0, dy: -30 }),
    )
    expect(result.isError).toBeUndefined()
    expect(validate).toHaveBeenCalledTimes(1)
    expect(validate.mock.calls[0]![0]).toMatchObject({ newText: 'Hey World', origin: [0, 730] })
    expect(pending).toHaveLength(1)
    expect(pending[0]!.moveBy).toEqual([0, 30])
    expect(result.output).toContain('Moved the paragraph')
  })

  it('reports the engine rejection instead of a move when the dry-run fails', async () => {
    const pending: LocalTextEdit[] = [
      {
        id: 'line',
        input: {
          pageIndex: 0,
          rect: [0, 700, 110, 712],
          oldText: 'Hello World',
          newText: 'Hey World',
          fontSize: 12,
        },
      },
    ]
    const validate = vi.fn(() => 'the text block cannot be moved as one unit')
    const deps = movingDeps(pending, validate)
    const result = await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 0, dy: -30 }),
    )
    expect(result.isError).toBe(true)
    expect(result.mutated).toBeUndefined()
    expect(result.output).not.toContain('Moved')
    expect(result.output).toContain('cannot be moved as one unit')
    expect(pending[0]!.moveBy).toBeUndefined()
  })

  it('skips the dry-run only for a pending block-level edit of the same block', async () => {
    const validate = vi.fn(() => null)
    const pending: LocalTextEdit[] = []
    const deps = movingDeps(pending, validate)
    await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 0, dy: -30 }),
    )
    expect(validate).toHaveBeenCalledTimes(1)
    const again = await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 0, dy: -10 }),
    )
    expect(again.isError).toBeUndefined()
    expect(validate).toHaveBeenCalledTimes(1)
    expect(pending).toHaveLength(1)
    expect(pending[0]!.input.translate).toEqual([0, 40])
    expect(again.output).toContain('y 48')
  })

  it('surfaces the rejection reason from the pending-edit pipeline', async () => {
    const deps = makeDeps({
      moveTextBlock: vi.fn(async () => ({ reason: 'the text block cannot be moved as one unit' })),
    })
    const result = await executePdfTool(
      deps,
      call('move_text_block', { page: 1, paragraph_text: 'Hello', dx: 0, dy: -30 }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('cannot be moved as one unit')
    expect(result.mutated).toBeUndefined()
  })
})

describe('insert_text', () => {
  it('queues a new text block at explicit coordinates with defaults', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('insert_text', { page: 1, text: 'Hello\nWorld', x: 100, y: 72 }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(deps.insertText).toHaveBeenCalledWith(
      expect.objectContaining({
        pageIndex: 0,
        // baseline of the first line: y 72 + default 14 pt → PDF y = 800 - 86
        origin: [100, 714],
        text: 'Hello\nWorld',
        fontSize: 14,
        color: [0, 0, 0],
        lineLeading: 14 * 1.2,
        lineXOffsets: undefined,
        align: undefined,
        rotate: 0,
      }),
    )
    expect(deps.gotoPage).toHaveBeenCalledWith(1)
  })

  it('passes style overrides through and validates the font', async () => {
    const deps = makeDeps()
    await executePdfTool(
      deps,
      call('insert_text', {
        page: 1,
        text: 'Hi',
        x: 0,
        y: 0,
        font_size: 20,
        color: '#ff8000',
        font: 'times',
        bold: true,
      }),
    )
    expect(deps.insertText).toHaveBeenCalledWith(
      expect.objectContaining({
        fontSize: 20,
        color: [255, 128, 0],
        font: 'times',
        bold: true,
        italic: undefined,
      }),
    )

    const bad = await executePdfTool(
      makeDeps({ editFonts: () => ['arial'] }),
      call('insert_text', { page: 1, text: 'Hi', font: 'times' }),
    )
    expect(bad.isError).toBe(true)
    expect(bad.output).toContain('arial')
  })

  it('centers lines within the widest line for align=center', async () => {
    const deps = makeDeps()
    await executePdfTool(
      deps,
      call('insert_text', { page: 1, text: 'aaaa\naa', x: 50, y: 100, align: 'center' }),
    )
    const input = vi.mocked(deps.insertText).mock.calls[0]![0]
    expect(input.align).toBe('center')
    // origin.x is the block's center: every line hangs left of it by half its width
    expect(input.lineXOffsets![0]).toBeLessThan(0)
    expect(input.lineXOffsets![1]).toBeGreaterThan(input.lineXOffsets![0])
    expect(input.lineXOffsets![1]).toBeLessThan(0)
    expect(input.origin[0]).toBeGreaterThan(50)
  })

  it('reports the same top-left that list_inserted_text derives for a narrow centered block', async () => {
    const deps = makeDeps({ insertText: vi.fn(() => 'dabc') })
    const inserted = await executePdfTool(
      deps,
      call('insert_text', { page: 1, text: 'aa', x: 100, y: 50, max_width: 200, align: 'center' }),
    )
    const input = vi.mocked(deps.insertText).mock.calls[0]![0]
    const listed = await executePdfTool(
      makeDeps({ textInserts: () => [{ id: 'dabc', input }] }),
      call('list_inserted_text', {}),
    )
    const pos = /x=(\d+), y=(\d+)/
    expect(inserted.output.match(pos)!.slice(1)).toEqual(listed.output.match(pos)!.slice(1))
    expect(listed.output).toContain('y=50')
    // the wrapped line is narrower than max_width, so the box left (100) is not the text left
    expect(Number(listed.output.match(pos)![1])).toBeGreaterThan(100)
  })

  it('reports the pending id of the new block', async () => {
    const deps = makeDeps({ insertText: vi.fn(() => 'dabc') })
    const result = await executePdfTool(
      deps,
      call('insert_text', { page: 1, text: 'Hi', x: 10, y: 10 }),
    )
    expect(result.output).toContain('Tdabc')
  })

  it('wraps paragraphs to max_width', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('insert_text', {
        page: 1,
        text: 'aaa bbb ccc',
        x: 0,
        y: 0,
        font_size: 10,
        max_width: 20,
      }),
    )
    expect(result.isError).toBeUndefined()
    const input = vi.mocked(deps.insertText).mock.calls[0]![0]
    expect(input.text.split('\n').length).toBeGreaterThan(1)
  })

  it('carries the page rotation into the insert', async () => {
    const deps = makeDeps({ pageGeom: () => ({ pw: 600, ph: 800, rot: 90 }) })
    await executePdfTool(deps, call('insert_text', { page: 1, text: 'Hi', x: 10, y: 10 }))
    expect(deps.insertText).toHaveBeenCalledWith(expect.objectContaining({ rotate: 90 }))
  })

  it('rejects empty text, bad colors, bad pages, and read-only documents', async () => {
    const empty = await executePdfTool(makeDeps(), call('insert_text', { page: 1, text: '  ' }))
    expect(empty.isError).toBe(true)

    const badColor = await executePdfTool(
      makeDeps(),
      call('insert_text', { page: 1, text: 'Hi', color: 'red' }),
    )
    expect(badColor.isError).toBe(true)

    const badPage = await executePdfTool(makeDeps(), call('insert_text', { page: 9, text: 'Hi' }))
    expect(badPage.isError).toBe(true)

    const deps = makeDeps({ readOnly: () => true })
    const ro = await executePdfTool(deps, call('insert_text', { page: 1, text: 'Hi' }))
    expect(ro.isError).toBe(true)
    expect(ro.output).toContain('read-only')
    expect(deps.insertText).not.toHaveBeenCalled()
  })
})

describe('form tools', () => {
  const widgets: Widget[][] = [
    [
      { subtype: 'Widget', fieldType: 'Tx', fieldName: 'name', fieldValue: 'Bob' },
      {
        subtype: 'Widget',
        fieldType: 'Btn',
        checkBox: true,
        fieldName: 'agree',
        fieldValue: 'Off',
      },
      {
        subtype: 'Widget',
        fieldType: 'Btn',
        radioButton: true,
        fieldName: 'color',
        buttonValue: 'red',
        fieldValue: '',
      },
      {
        subtype: 'Widget',
        fieldType: 'Btn',
        radioButton: true,
        fieldName: 'color',
        buttonValue: 'blue',
        fieldValue: '',
      },
      {
        subtype: 'Widget',
        fieldType: 'Ch',
        fieldName: 'size',
        fieldValue: 'M',
        options: [{ exportValue: 'S' }, { exportValue: 'M' }, { displayValue: 'Large' }],
      },
      { subtype: 'Widget', fieldType: 'Tx', fieldName: 'locked', readOnly: true },
      { subtype: 'Link' },
    ],
  ]
  const withForm = (over: Partial<PdfAiDeps> = {}) =>
    makeDeps({ doc: () => fakeDoc(['form page'], widgets), pageCount: () => 1, ...over })

  it('lists fields with kinds, aggregated radio options, and skips read-only widgets', async () => {
    const result = await executePdfTool(withForm(), call('list_form_fields'))
    expect(result.output).toContain('name (text, page 1)')
    expect(result.output).toContain('agree (checkbox, page 1)')
    expect(result.output).toContain('options[red, blue]')
    expect(result.output).toContain('options[S, M, Large]')
    expect(result.output).not.toContain('locked')
  })

  it('shows pending unsaved edits instead of stored values', async () => {
    const edits = new Map<string, FormValueInput>([
      ['name', { name: 'name', kind: 'text', value: 'Alice' }],
      ['agree', { name: 'agree', kind: 'checkbox', checked: true }],
    ])
    const result = await executePdfTool(
      withForm({ formEdits: () => edits }),
      call('list_form_fields'),
    )
    expect(result.output).toContain('current value: Alice')
    expect(result.output).toContain('current value: true')
  })

  it('fills a text field and jumps to its page', async () => {
    const deps = withForm()
    const result = await executePdfTool(
      deps,
      call('fill_form_field', { name: 'name', value: 'Alice' }),
    )
    expect(result.mutated).toBe(true)
    expect(deps.applyFormEdit).toHaveBeenCalledWith({ name: 'name', kind: 'text', value: 'Alice' })
    expect(deps.gotoPage).toHaveBeenCalledWith(1)
  })

  it('requires checked for checkboxes and validates choice options', async () => {
    const noChecked = await executePdfTool(withForm(), call('fill_form_field', { name: 'agree' }))
    expect(noChecked.isError).toBe(true)

    const badOption = await executePdfTool(
      withForm(),
      call('fill_form_field', { name: 'size', value: 'XXL' }),
    )
    expect(badOption.isError).toBe(true)
    expect(badOption.output).toContain('not among the options')

    const deps = withForm()
    const ok = await executePdfTool(deps, call('fill_form_field', { name: 'color', value: 'blue' }))
    expect(ok.mutated).toBe(true)
    expect(deps.applyFormEdit).toHaveBeenCalledWith({ name: 'color', kind: 'radio', value: 'blue' })
  })

  it('rejects unknown field names', async () => {
    const result = await executePdfTool(
      withForm(),
      call('fill_form_field', { name: 'nope', value: 'x' }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('No field named')
  })
})

describe('rotate_page / delete_page', () => {
  it('maps direction to a signed 90-degree delta', async () => {
    const deps = makeDeps()
    await executePdfTool(deps, call('rotate_page', { page: 1, direction: 'left' }))
    expect(deps.rotatePages).toHaveBeenCalledWith([0], -90)
    await executePdfTool(deps, call('rotate_page', { page: 2, direction: 'right' }))
    expect(deps.rotatePages).toHaveBeenCalledWith([1], 90)
  })

  it('rotates a page by 180 and rejects an unknown direction', async () => {
    const deps = makeDeps()
    const half = await executePdfTool(deps, call('rotate_page', { page: 2, direction: '180' }))
    expect(half.mutated).toBe(true)
    expect(half.output).toContain('180°')
    expect(deps.rotatePages).toHaveBeenCalledWith([1], 180)
    const bad = await executePdfTool(deps, call('rotate_page', { page: 1, direction: 'up' }))
    expect(bad.isError).toBe(true)
    expect(deps.rotatePages).toHaveBeenCalledTimes(1)
  })

  it('rotates every visible page in one call when all is true', async () => {
    const deps = makeDeps({ pageCount: () => 3, pageOrder: () => [2, 0] })
    const result = await executePdfTool(
      deps,
      call('rotate_page', { direction: 'right', all: true }),
    )
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('all 2 pages')
    expect(deps.rotatePages).toHaveBeenCalledTimes(1)
    expect(deps.rotatePages).toHaveBeenCalledWith([2, 0], 90)
    expect(deps.gotoPage).not.toHaveBeenCalled()
    const ro = await executePdfTool(
      makeDeps({ readOnly: () => true }),
      call('rotate_page', { direction: 'left', all: true }),
    )
    expect(ro.isError).toBe(true)
  })

  it('deletes a page and reports failure when the last page must remain', async () => {
    const deps = makeDeps()
    const ok = await executePdfTool(deps, call('delete_page', { page: 2 }))
    expect(ok.mutated).toBe(true)
    expect(deps.deletePage).toHaveBeenCalledWith(1)

    const blocked = await executePdfTool(
      makeDeps({ deletePage: () => false }),
      call('delete_page', { page: 1 }),
    )
    expect(blocked.isError).toBe(true)
  })

  it('blocks mutations on read-only documents', async () => {
    const deps = makeDeps({ readOnly: () => true })
    for (const c of [
      call('rotate_page', { page: 1, direction: 'left' }),
      call('delete_page', { page: 1 }),
    ]) {
      const result = await executePdfTool(deps, c)
      expect(result.isError).toBe(true)
      expect(deps.rotatePages).not.toHaveBeenCalled()
      expect(deps.deletePage).not.toHaveBeenCalled()
    }
  })
})

describe('get_outline / unknown tools', () => {
  it('renders the outline tree with indentation', async () => {
    const deps = makeDeps({
      outline: () => [
        { title: 'Chapter 1', items: [{ title: 'Section 1.1' }] },
        { title: 'Chapter 2' },
      ],
    })
    const result = await executePdfTool(deps, call('get_outline'))
    expect(result.output).toBe('Chapter 1\n  Section 1.1\nChapter 2')
  })

  it('reports when the document has no outline', async () => {
    const result = await executePdfTool(makeDeps(), call('get_outline'))
    expect(result.output).toContain('no outline')
  })

  it('errors on unknown tool names', async () => {
    const result = await executePdfTool(makeDeps(), call('nope'))
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Unknown tool')
  })
})

describe('image_search / generate_image', () => {
  it('lists numbered results with direct links', async () => {
    const result = await executePdfTool(makeDeps(), call('image_search', { query: 'cat' }))
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('1. A cat [800x600]')
    expect(result.output).toContain('https://img.example/cat.jpg')
  })

  it('surfaces backend failures as retryable errors, not empty galleries', async () => {
    const deps = makeDeps({
      searchImages: async () => ({ images: [], method: 'error', error: 'boom' }),
    })
    const result = await executePdfTool(deps, call('image_search', { query: 'cat' }))
    expect(result.isError).toBe(true)
    expect(result.output).toContain('boom')
  })

  it('returns the generated image URL and errors when generation fails', async () => {
    const ok = await executePdfTool(makeDeps(), call('generate_image', { prompt: 'a diagram' }))
    expect(ok.isError).toBeUndefined()
    expect(ok.output).toContain('https://img.example/generated.png')

    const failed = await executePdfTool(
      makeDeps({ generateImage: async () => ({ error: 'not logged in' }) }),
      call('generate_image', { prompt: 'a diagram' }),
    )
    expect(failed.isError).toBe(true)
    expect(failed.output).toContain('not logged in')
  })
})

describe('list_page_images', () => {
  it('numbers images top-to-bottom with top-left-origin coordinates', async () => {
    const result = await executePdfTool(makeDeps(), call('list_page_images', {}))
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('Page 1 (600 × 800 pt):')
    // [50,600,150,700] sits higher on the page (y=100 from top) → image 1
    expect(result.output).toContain('image 1: 100 × 100 pt at x=50, y=100, below the text')
    expect(result.output).toContain('image 2: 100 × 100 pt at x=200, y=600, above the text')
  })

  it('marks images that already have a pending edit', async () => {
    const deps = makeDeps({ isImageClaimed: (ref) => ref.rect[0] === 50 })
    const result = await executePdfTool(deps, call('list_page_images', { page: 1 }))
    expect(result.output).toContain(
      'image 1: 100 × 100 pt at x=50, y=100, below the text — has a pending unsaved edit',
    )
  })

  it('reports pages without embedded images', async () => {
    const result = await executePdfTool(makeDeps(), call('list_page_images', { page: 2 }))
    expect(result.output).toContain('Page 2 has no embedded images')
  })
})

describe('insert_image', () => {
  it('downloads, sizes to natural dimensions, and centers by default', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('insert_image', { page: 2, url: 'https://img.example/cat.jpg' }),
    )
    expect(result.mutated).toBe(true)
    // 400×300 px → 300×225 pt (capped at half the 600 pt page width), centered on 600×800
    expect(deps.insertImage).toHaveBeenCalledWith(
      1,
      'PNGB64',
      [150, 287.5, 450, 512.5],
      'belowText',
    )
    expect(deps.gotoPage).toHaveBeenCalledWith(2)
  })

  it('places the image below anchor text and honors the layer parameter', async () => {
    const deps = makeDeps()
    // Page 1 anchor "Hello World" occupies [0,700,110,712] → below = y 108 from the top
    const result = await executePdfTool(
      deps,
      call('insert_image', {
        page: 1,
        url: 'https://img.example/cat.jpg',
        anchor_text: 'Hello World',
        layer: 'above_text',
      }),
    )
    expect(result.mutated).toBe(true)
    expect(deps.insertImage).toHaveBeenCalledWith(0, 'PNGB64', [0, 467, 300, 692], 'aboveText')
  })

  it('rejects non-http urls and reports download failures', async () => {
    const bad = await executePdfTool(
      makeDeps(),
      call('insert_image', { page: 1, url: 'file:///etc/passwd' }),
    )
    expect(bad.isError).toBe(true)

    const deps = makeDeps({ fetchImage: async () => null })
    const failed = await executePdfTool(
      deps,
      call('insert_image', { page: 1, url: 'https://img.example/cat.jpg' }),
    )
    expect(failed.isError).toBe(true)
    expect(deps.insertImage).not.toHaveBeenCalled()
  })

  it('rejects anchor text that is not on the page', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('insert_image', { page: 2, url: 'https://x.example/a.png', anchor_text: 'Hello World' }),
    )
    expect(result.isError).toBe(true)
    expect(deps.insertImage).not.toHaveBeenCalled()
  })
})

describe('transform_image / delete_image', () => {
  it('moves and resizes by listing number, keeping the aspect ratio', async () => {
    const deps = makeDeps()
    // image 1 = [50,600,150,700]; width 50 → height 50 by aspect; y kept at 100 from top
    const result = await executePdfTool(
      deps,
      call('transform_image', { page: 1, image_number: 1, x: 10, width: 50 }),
    )
    expect(result.mutated).toBe(true)
    expect(deps.transformImage).toHaveBeenCalledWith(PAGE_IMAGES[1], [10, 650, 60, 700], undefined)
  })

  it('deletes by listing number', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(deps, call('delete_image', { page: 1, image_number: 2 }))
    expect(result.mutated).toBe(true)
    expect(deps.deleteImage).toHaveBeenCalledWith(PAGE_IMAGES[0])
  })

  it('rotate_image cw swaps the footprint about the center and passes quarter turns', async () => {
    const deps = makeDeps({
      listImages: vi.fn(async () => [
        {
          pageIndex: 0,
          rect: [100, 100, 300, 200] as [number, number, number, number],
          aboveText: true,
        },
      ]),
    })
    const result = await executePdfTool(
      deps,
      call('rotate_image', { page: 1, image_number: 1, direction: 'cw' }),
    )
    expect(result.mutated).toBe(true)
    // 200×100 about center (200,150) → 100×200 footprint
    expect(deps.transformImage).toHaveBeenCalledWith(
      expect.objectContaining({ rect: [100, 100, 300, 200] }),
      [150, 50, 250, 250],
      undefined,
      1,
    )
  })

  it('rotate_image 180 keeps the footprint; bad direction is rejected', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('rotate_image', { page: 1, image_number: 1, direction: '180' }),
    )
    expect(result.mutated).toBe(true)
    expect(deps.transformImage).toHaveBeenCalledWith(
      PAGE_IMAGES[1],
      [50, 600, 150, 700],
      undefined,
      2,
    )
    const bad = await executePdfTool(
      makeDeps(),
      call('rotate_image', { page: 1, image_number: 1, direction: 'flip' }),
    )
    expect(bad.isError).toBe(true)
  })

  it('replace_image fetches the url and swaps pixels in place', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('replace_image', { page: 1, image_number: 1, url: 'https://img.example/new.png' }),
    )
    expect(result.mutated).toBe(true)
    expect(deps.fetchImage).toHaveBeenCalledWith('https://img.example/new.png')
    expect(deps.replaceImage).toHaveBeenCalledWith(PAGE_IMAGES[1], 'PNGB64')
  })

  it('replace_image rejects non-http urls and does not mutate after a stop', async () => {
    const deps = makeDeps()
    const bad = await executePdfTool(
      deps,
      call('replace_image', { page: 1, image_number: 1, url: 'file:///etc/passwd' }),
    )
    expect(bad.isError).toBe(true)
    expect(deps.replaceImage).not.toHaveBeenCalled()

    const ctl = new AbortController()
    const aborted = makeDeps({
      fetchImage: vi.fn(async () => {
        ctl.abort()
        return { png: 'PNGB64', width: 400, height: 300 }
      }),
    })
    const result = await executePdfTool(
      aborted,
      call('replace_image', { page: 1, image_number: 1, url: 'https://img.example/new.png' }),
      ctl.signal,
    )
    expect(result.isError).toBe(true)
    expect(aborted.replaceImage).not.toHaveBeenCalled()
  })

  it('refuses images that already have a pending edit and out-of-range numbers', async () => {
    const claimed = await executePdfTool(
      makeDeps({ isImageClaimed: () => true }),
      call('delete_image', { page: 1, image_number: 1 }),
    )
    expect(claimed.isError).toBe(true)
    expect(claimed.output).toContain('pending unsaved edit')

    const missing = await executePdfTool(
      makeDeps(),
      call('transform_image', { page: 1, image_number: 9 }),
    )
    expect(missing.isError).toBe(true)
    expect(missing.output).toContain('only has 2 image(s)')
  })

  it('maps display coordinates through page rotation (90°)', async () => {
    const deps = makeDeps({ pageGeom: () => ({ pw: 600, ph: 800, rot: 90 }) })
    // Displayed page is 800 × 600. [50,600,150,700] displays at (600,50); [200,100,300,200] at (100,200)
    const listed = await executePdfTool(deps, call('list_page_images', { page: 1 }))
    expect(listed.output).toContain('Page 1 (800 × 600 pt):')
    expect(listed.output).toContain('image 1: 100 × 100 pt at x=600, y=50')
    expect(listed.output).toContain('image 2: 100 × 100 pt at x=100, y=200')

    // Display box (100,50)-(400,275) → PDF space [50,100,275,400] under rot 90
    const inserted = await executePdfTool(
      deps,
      call('insert_image', { page: 1, url: 'https://img.example/cat.jpg', x: 100, y: 50 }),
    )
    expect(inserted.mutated).toBe(true)
    expect(deps.insertImage).toHaveBeenCalledWith(0, 'PNGB64', [50, 100, 275, 400], 'belowText')
  })

  it('does not mutate after the run was stopped', async () => {
    const ctl = new AbortController()
    const deps = makeDeps({
      fetchImage: vi.fn(async () => {
        ctl.abort()
        return { png: 'PNGB64', width: 400, height: 300 }
      }),
    })
    const result = await executePdfTool(
      deps,
      call('insert_image', { page: 1, url: 'https://img.example/cat.jpg' }),
      ctl.signal,
    )
    expect(result.isError).toBe(true)
    expect(deps.insertImage).not.toHaveBeenCalled()
  })

  it('blocks image mutations on read-only documents', async () => {
    const deps = makeDeps({ readOnly: () => true })
    for (const c of [
      call('insert_image', { page: 1, url: 'https://x.example/a.png' }),
      call('transform_image', { page: 1, image_number: 1 }),
      call('delete_image', { page: 1, image_number: 1 }),
    ]) {
      const result = await executePdfTool(deps, c)
      expect(result.isError).toBe(true)
    }
    expect(deps.insertImage).not.toHaveBeenCalled()
    expect(deps.transformImage).not.toHaveBeenCalled()
    expect(deps.deleteImage).not.toHaveBeenCalled()
  })
})

describe('image pixel bakes (flip / opacity / crop / remove background)', () => {
  it('flip_image maps the axis and reports the unsaved edit', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('flip_image', { page: 1, image_number: 1, axis: 'horizontal' }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('unsaved')
    expect(deps.bakeImage).toHaveBeenCalledWith(
      PAGE_IMAGES[1],
      { kind: 'flip', axis: 'h' },
      undefined,
    )
    expect(deps.gotoPage).toHaveBeenCalledWith(1)
    const vertical = await executePdfTool(
      deps,
      call('flip_image', { page: 1, image_number: 2, axis: 'vertical' }),
    )
    expect(vertical.mutated).toBe(true)
    expect(deps.bakeImage).toHaveBeenLastCalledWith(
      PAGE_IMAGES[0],
      { kind: 'flip', axis: 'v' },
      undefined,
    )
  })

  it('flip_image swaps the axis on a quarter-turned page so the flip is horizontal as displayed', async () => {
    const rot90 = makeDeps({ pageGeom: () => ({ pw: 600, ph: 800, rot: 90 }) })
    const result = await executePdfTool(
      rot90,
      call('flip_image', { page: 1, image_number: 1, axis: 'horizontal' }),
    )
    expect(result.mutated).toBe(true)
    expect(rot90.bakeImage).toHaveBeenCalledWith(
      expect.objectContaining({ pageIndex: 0 }),
      { kind: 'flip', axis: 'v' },
      undefined,
    )
    const rot180 = makeDeps({ pageGeom: () => ({ pw: 600, ph: 800, rot: 180 }) })
    await executePdfTool(rot180, call('flip_image', { page: 1, image_number: 1, axis: 'vertical' }))
    expect(rot180.bakeImage).toHaveBeenCalledWith(
      expect.objectContaining({ pageIndex: 0 }),
      { kind: 'flip', axis: 'v' },
      undefined,
    )
  })

  it('a rotate_page earlier in the same turn is honored by a later flip', async () => {
    let rot = 0
    const deps = makeDeps({
      pageGeom: () => ({ pw: 600, ph: 800, rot }),
      rotatePages: vi.fn((_pages: number[], dir: number) => {
        rot = (rot + dir + 360) % 360
      }),
    })
    await executePdfTool(deps, call('rotate_page', { page: 1, direction: 'right' }))
    await executePdfTool(deps, call('flip_image', { page: 1, image_number: 1, axis: 'horizontal' }))
    expect(deps.bakeImage).toHaveBeenCalledWith(
      expect.objectContaining({ pageIndex: 0 }),
      { kind: 'flip', axis: 'v' },
      undefined,
    )
  })

  it('flip_image rejects an unknown axis without baking', async () => {
    const deps = makeDeps()
    const bad = await executePdfTool(
      deps,
      call('flip_image', { page: 1, image_number: 1, axis: 'diagonal' }),
    )
    expect(bad.isError).toBe(true)
    expect(bad.output).toContain('axis')
    expect(deps.bakeImage).not.toHaveBeenCalled()
  })

  it('set_image_opacity passes the alpha and rejects values outside 0..1', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('set_image_opacity', { page: 1, image_number: 1, opacity: 0.35 }),
    )
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('35%')
    expect(deps.bakeImage).toHaveBeenCalledWith(
      PAGE_IMAGES[1],
      { kind: 'opacity', alpha: 0.35 },
      undefined,
    )
    for (const opacity of [1.5, -0.1, 'half']) {
      const bad = await executePdfTool(
        deps,
        call('set_image_opacity', { page: 1, image_number: 1, opacity }),
      )
      expect(bad.isError).toBe(true)
    }
    expect(deps.bakeImage).toHaveBeenCalledTimes(1)
  })

  it('crop_image turns insets into kept fractions and reports the new footprint', async () => {
    const deps = makeDeps()
    // image 1 = [50,600,150,700] → displayed 100×100 at x=50, y=100
    const result = await executePdfTool(
      deps,
      call('crop_image', { page: 1, image_number: 1, left: 0.25, bottom: 0.5 }),
    )
    expect(result.mutated).toBe(true)
    expect(deps.bakeImage).toHaveBeenCalledWith(
      PAGE_IMAGES[1],
      {
        kind: 'crop',
        crop: { l: 0.25, t: 0, r: 1, b: 0.5 },
      },
      undefined,
    )
    expect(result.output).toContain('now 75 × 50 pt at x=75, y=100')
  })

  it('crop_image rotates display-space insets into the object-space crop on rotated pages', async () => {
    // rot 90: image [50,600,150,700] displays as 100×100 at x=600, y=50; its displayed
    // left edge is the object-space bottom edge
    const one = vi.fn(async () => [PAGE_IMAGES[1]!])
    const rot90 = makeDeps({ listImages: one, pageGeom: () => ({ pw: 600, ph: 800, rot: 90 }) })
    const left = await executePdfTool(
      rot90,
      call('crop_image', { page: 1, image_number: 1, left: 0.25 }),
    )
    expect(left.mutated).toBe(true)
    expect(rot90.bakeImage).toHaveBeenCalledWith(
      PAGE_IMAGES[1],
      { kind: 'crop', crop: { l: 0, t: 0, r: 1, b: 0.75 } },
      undefined,
    )
    expect(left.output).toContain('now 75 × 100 pt at x=625, y=50')
    // rot 270: the displayed top edge is the object-space right edge
    const rot270 = makeDeps({ listImages: one, pageGeom: () => ({ pw: 600, ph: 800, rot: 270 }) })
    await executePdfTool(rot270, call('crop_image', { page: 1, image_number: 1, top: 0.25 }))
    expect(rot270.bakeImage).toHaveBeenCalledWith(
      PAGE_IMAGES[1],
      { kind: 'crop', crop: { l: 0, t: 0, r: 0.75, b: 1 } },
      undefined,
    )
  })

  it('crop_image rejects insets that leave nothing, out-of-range insets, and a no-op crop', async () => {
    const deps = makeDeps()
    const nothing = await executePdfTool(
      deps,
      call('crop_image', { page: 1, image_number: 1, left: 0.6, right: 0.4 }),
    )
    expect(nothing.isError).toBe(true)
    expect(nothing.output).toContain('leave no image')
    const range = await executePdfTool(
      deps,
      call('crop_image', { page: 1, image_number: 1, top: 1 }),
    )
    expect(range.isError).toBe(true)
    const noop = await executePdfTool(deps, call('crop_image', { page: 1, image_number: 1 }))
    expect(noop.isError).toBe(true)
    expect(deps.bakeImage).not.toHaveBeenCalled()
  })

  it('remove_image_background defaults the tolerance and validates an explicit one', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('remove_image_background', { page: 1, image_number: 1 }),
    )
    expect(result.mutated).toBe(true)
    expect(deps.bakeImage).toHaveBeenCalledWith(
      PAGE_IMAGES[1],
      { kind: 'cutout', tolerance: 30 },
      undefined,
    )
    const custom = await executePdfTool(
      deps,
      call('remove_image_background', { page: 1, image_number: 1, tolerance: 55 }),
    )
    expect(custom.output).toContain('tolerance 55')
    expect(deps.bakeImage).toHaveBeenLastCalledWith(
      PAGE_IMAGES[1],
      {
        kind: 'cutout',
        tolerance: 55,
      },
      undefined,
    )
    const bad = await executePdfTool(
      deps,
      call('remove_image_background', { page: 1, image_number: 1, tolerance: 120 }),
    )
    expect(bad.isError).toBe(true)
    expect(deps.bakeImage).toHaveBeenCalledTimes(2)
  })

  it('refuses an image that already has a pending edit, read-only files, and a failed bake', async () => {
    const claimed = makeDeps({ isImageClaimed: () => true })
    const conflict = await executePdfTool(
      claimed,
      call('flip_image', { page: 1, image_number: 1, axis: 'vertical' }),
    )
    expect(conflict.isError).toBe(true)
    expect(conflict.output).toContain('already has a pending unsaved edit')
    expect(claimed.bakeImage).not.toHaveBeenCalled()

    const ro = makeDeps({ readOnly: () => true })
    for (const [name, extra] of [
      ['flip_image', { axis: 'horizontal' }],
      ['set_image_opacity', { opacity: 0.5 }],
      ['crop_image', { left: 0.1 }],
      ['remove_image_background', {}],
    ] as const) {
      const result = await executePdfTool(ro, call(name, { page: 1, image_number: 1, ...extra }))
      expect(result.isError).toBe(true)
    }
    expect(ro.bakeImage).not.toHaveBeenCalled()

    const failing = makeDeps({ bakeImage: vi.fn(async () => false) })
    const failed = await executePdfTool(
      failing,
      call('crop_image', { page: 1, image_number: 1, right: 0.2 }),
    )
    expect(failed.isError).toBe(true)
    expect(failed.mutated).toBeUndefined()
    expect(failing.gotoPage).not.toHaveBeenCalled()
  })

  it('a second edit on an image claimed earlier in the same turn is refused', async () => {
    const claimedKeys = new Set<string>()
    const deps = makeDeps({
      isImageClaimed: (ref) => claimedKeys.has(`${ref.pageIndex}:${ref.rect.join()}`),
      bakeImage: vi.fn(async (ref) => {
        claimedKeys.add(`${ref.pageIndex}:${ref.rect.join()}`)
        return true
      }),
    })
    const first = await executePdfTool(
      deps,
      call('flip_image', { page: 1, image_number: 1, axis: 'horizontal' }),
    )
    expect(first.mutated).toBe(true)
    const second = await executePdfTool(
      deps,
      call('set_image_opacity', { page: 1, image_number: 1, opacity: 0.5 }),
    )
    expect(second.isError).toBe(true)
    expect(second.output).toContain('already has a pending unsaved edit')
    expect(deps.bakeImage).toHaveBeenCalledTimes(1)
  })

  it('passes the abort signal to the bake and reports a stop instead of a bake failure', async () => {
    const ctl = new AbortController()
    const deps = makeDeps({
      bakeImage: vi.fn(async (_ref, _op, signal?: AbortSignal) => {
        ctl.abort()
        return !signal?.aborted
      }),
    })
    const result = await executePdfTool(
      deps,
      call('remove_image_background', { page: 1, image_number: 1 }),
      ctl.signal,
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('stopped by the user')
    expect(deps.bakeImage).toHaveBeenCalledWith(
      PAGE_IMAGES[1],
      { kind: 'cutout', tolerance: 30 },
      ctl.signal,
    )
    expect(deps.gotoPage).not.toHaveBeenCalled()
  })
})

describe('annotations tools', () => {
  const thread = {
    key: 'S12',
    saved: null,
    pendingId: null,
    author: 'Alice',
    contents: 'Fix this total',
    timeMs: Date.UTC(2026, 7, 20),
    color: null,
    at: [10, 700] as [number, number],
    replies: [],
  }

  it('read_annotations lists threads and quotes markup-covered text', async () => {
    const deps = makeDeps({
      annotationsOn: vi.fn(async (idx: number) =>
        idx === 0
          ? {
              threads: [thread],
              // covers the first ~half of "Hello World" (x 0..50 of the 110pt item)
              markups: [
                {
                  key: 'S7',
                  type: 'highlight' as const,
                  quads: [[0, 712, 50, 712, 0, 700, 50, 700]],
                  saved: true,
                },
              ],
            }
          : { threads: [], markups: [] },
      ),
    })
    const result = await executePdfTool(deps, call('read_annotations', {}))
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('[Page 1]')
    expect(result.output).toContain('Note S12')
    expect(result.output).toContain('Alice')
    expect(result.output).toContain('2026-08-20')
    expect(result.output).toContain('Fix this total')
    expect(result.output).toContain('Markup S7: highlight')
    expect(result.output).toContain('Hello')
  })

  it('read_annotations gives replies their own ids', async () => {
    const withReply = { ...thread, replies: [{ ...thread, key: 'Pd3k', saved: null, replies: [] }] }
    const deps = makeDeps({
      annotationsOn: vi.fn(async (idx: number) =>
        idx === 0 ? { threads: [withReply], markups: [] } : { threads: [], markups: [] },
      ),
    })
    const result = await executePdfTool(deps, call('read_annotations', {}))
    expect(result.output).toContain('- Note S12')
    expect(result.output).toContain('  - reply Pd3k (unsaved)')
  })

  const PAGE_MARKUPS = [
    {
      key: 'S7',
      type: 'highlight' as const,
      quads: [[0, 712, 50, 712, 0, 700, 50, 700]],
      saved: true,
    },
    {
      key: 'Pd1',
      type: 'underline' as const,
      quads: [[60, 712, 110, 712, 60, 700, 110, 700]],
      saved: false,
    },
  ]

  it('delete_markup removes the listed ids, or every markup of a type on the page', async () => {
    const deps = makeDeps({
      annotationsOn: vi.fn(async () => ({ threads: [], markups: PAGE_MARKUPS })),
    })
    const one = await executePdfTool(deps, call('delete_markup', { page: 1, markup_ids: ['S7'] }))
    expect(one.isError).toBeUndefined()
    expect(one.mutated).toBe(true)
    expect(one.output).toContain('1 highlight')
    expect(deps.deleteMarkups).toHaveBeenLastCalledWith(0, ['S7'])

    const byType = await executePdfTool(deps, call('delete_markup', { page: 1, type: 'underline' }))
    expect(byType.isError).toBeUndefined()
    expect(deps.deleteMarkups).toHaveBeenLastCalledWith(0, ['Pd1'])

    const all = await executePdfTool(deps, call('delete_markup', { page: 1 }))
    expect(all.output).toContain('1 highlight, 1 underline')
    expect(deps.deleteMarkups).toHaveBeenLastCalledWith(0, ['S7', 'Pd1'])
  })

  it('delete_markup rejects unknown ids, empty pages, and read-only documents', async () => {
    const deps = makeDeps({
      annotationsOn: vi.fn(async () => ({ threads: [], markups: PAGE_MARKUPS })),
    })
    const unknown = await executePdfTool(
      deps,
      call('delete_markup', { page: 1, markup_ids: ['S99'] }),
    )
    expect(unknown.isError).toBe(true)
    expect(unknown.output).toContain('S99')
    const none = await executePdfTool(deps, call('delete_markup', { page: 1, type: 'strikeout' }))
    expect(none.isError).toBe(true)
    const notList = await executePdfTool(deps, call('delete_markup', { page: 1, markup_ids: 7 }))
    expect(notList.isError).toBe(true)
    expect(deps.deleteMarkups).not.toHaveBeenCalled()
    const single = await executePdfTool(deps, call('delete_markup', { page: 1, markup_ids: 'S7' }))
    expect(single.isError).toBeUndefined()
    expect(deps.deleteMarkups).toHaveBeenLastCalledWith(0, ['S7'])
    const ro = await executePdfTool(
      makeDeps({ readOnly: () => true }),
      call('delete_markup', { page: 1 }),
    )
    expect(ro.isError).toBe(true)
  })

  it('delete_note deletes the thread found by id and reports its replies', async () => {
    const withReply = { ...thread, replies: [{ ...thread, key: 'S13', replies: [] }] }
    const deps = makeDeps({ findNoteRoot: vi.fn(async () => withReply) })
    const result = await executePdfTool(deps, call('delete_note', { page: 1, note_id: 'S12' }))
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('Deleted note S12 and its 1 reply')
    expect(deps.deleteNoteThread).toHaveBeenCalledWith(0, withReply)

    const missing = await executePdfTool(
      makeDeps(),
      call('delete_note', { page: 1, note_id: 'S99' }),
    )
    expect(missing.isError).toBe(true)
    expect(missing.output).toContain('S99')
  })

  it('read_annotations reports an annotation-free range', async () => {
    const result = await executePdfTool(makeDeps(), call('read_annotations', {}))
    expect(result.output).toContain('No notes or markups')
  })

  it('add_note anchors at the end of anchor_text', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('add_note', { page: 1, text: 'Check spelling', anchor_text: 'World' }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(deps.addNote).toHaveBeenCalledWith(0, [110, 712], 'Check spelling', undefined)
    expect(result.output).toContain('Pn1')
  })

  it('add_note forwards an optional pin color as rgb 0-1 and rejects bad hex', async () => {
    const deps = makeDeps()
    const ok = await executePdfTool(
      deps,
      call('add_note', { page: 1, text: 'Check', anchor_text: 'World', color: '#FF8000' }),
    )
    expect(ok.isError).toBeUndefined()
    expect(deps.addNote).toHaveBeenCalledWith(0, [110, 712], 'Check', [1, 128 / 255, 0])
    const bad = await executePdfTool(
      deps,
      call('add_note', { page: 1, text: 'Check', anchor_text: 'World', color: 'orange' }),
    )
    expect(bad.isError).toBe(true)
    expect(bad.output).toContain('#RRGGBB')
    expect(deps.addNote).toHaveBeenCalledTimes(1)
  })

  it('add_note requires a position', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(deps, call('add_note', { page: 1, text: 'hi' }))
    expect(result.isError).toBe(true)
    expect(deps.addNote).not.toHaveBeenCalled()
  })

  it('reply_note replies to a found root and errors on unknown ids', async () => {
    const deps = makeDeps({
      findNoteRoot: vi.fn(async (_i: number, key: string) => (key === 'S12' ? thread : null)),
    })
    const ok = await executePdfTool(
      deps,
      call('reply_note', { page: 1, note_id: 'S12', text: 'Done' }),
    )
    expect(ok.mutated).toBe(true)
    expect(deps.replyToThread).toHaveBeenCalledWith(0, thread, 'Done')
    const bad = await executePdfTool(
      deps,
      call('reply_note', { page: 1, note_id: 'S99', text: 'x' }),
    )
    expect(bad.isError).toBe(true)
    expect(bad.output).toContain('read_annotations')
  })

  describe('edit_note', () => {
    const reply = { ...thread, key: 'Pd3k', saved: null, pendingId: 'd3k', replies: [] }
    const roots = [{ ...thread, replies: [reply] }]
    const depsWith = (over: Partial<PdfAiDeps> = {}) =>
      makeDeps({
        annotationsOn: vi.fn(async (idx: number) =>
          idx === 0 ? { threads: roots, markups: [] } : { threads: [], markups: [] },
        ),
        ...over,
      })

    it('edits a saved thread root by S id', async () => {
      const deps = depsWith()
      const result = await executePdfTool(
        deps,
        call('edit_note', { page: 1, note_id: 'S12', text: ' Fix this subtotal ' }),
      )
      expect(result.isError).toBeUndefined()
      expect(result.mutated).toBe(true)
      expect(result.output).toContain('Edited note S12 on page 1 (unsaved')
      expect(deps.editNote).toHaveBeenCalledWith(0, roots[0], 'Fix this subtotal')
      expect(deps.gotoPage).toHaveBeenCalledWith(1)
    })

    it('edits a pending reply by P id', async () => {
      const deps = depsWith()
      const result = await executePdfTool(
        deps,
        call('edit_note', { page: 1, note_id: 'Pd3k', text: 'Done in v2' }),
      )
      expect(result.mutated).toBe(true)
      expect(deps.editNote).toHaveBeenCalledWith(0, reply, 'Done in v2')
    })

    it('rejects empty text and points at delete_note', async () => {
      const deps = depsWith()
      const result = await executePdfTool(
        deps,
        call('edit_note', { page: 1, note_id: 'S12', text: '   ' }),
      )
      expect(result.isError).toBe(true)
      expect(result.output).toContain('delete_note')
      expect(deps.editNote).not.toHaveBeenCalled()
    })

    it('rejects unknown ids, unchanged text, and read-only documents', async () => {
      const deps = depsWith()
      const missing = await executePdfTool(
        deps,
        call('edit_note', { page: 1, note_id: 'S99', text: 'x' }),
      )
      expect(missing.isError).toBe(true)
      expect(missing.output).toContain('S99')
      expect(missing.output).toContain('read_annotations')
      const same = await executePdfTool(
        deps,
        call('edit_note', { page: 1, note_id: 'S12', text: 'Fix this total' }),
      )
      expect(same.isError).toBeUndefined()
      expect(same.mutated).toBeUndefined()
      expect(deps.editNote).not.toHaveBeenCalled()
      const ro = await executePdfTool(
        depsWith({ readOnly: () => true }),
        call('edit_note', { page: 1, note_id: 'S12', text: 'x' }),
      )
      expect(ro.isError).toBe(true)
    })
  })

  it('markup_text forwards a custom color as rgb 0-1', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('markup_text', { page: 1, text: 'Hello', type: 'highlight', color: '#ff0000' }),
    )
    expect(result.isError).toBeUndefined()
    expect(deps.addMarkup).toHaveBeenCalledWith('highlight', 0, expect.anything(), [1, 0, 0])
  })

  it('markup_text without color leaves the default to the app', async () => {
    const deps = makeDeps()
    await executePdfTool(deps, call('markup_text', { page: 1, text: 'Hello', type: 'highlight' }))
    expect(deps.addMarkup).toHaveBeenCalledWith('highlight', 0, expect.anything(), undefined)
  })

  it('insert_text places anchored text to the right of the anchor', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('insert_text', { page: 1, text: 'John Doe', anchor_text: 'World' }),
    )
    expect(result.isError).toBeUndefined()
    const input = (deps.insertText as ReturnType<typeof vi.fn>).mock.calls[0]![0]
    // anchor "World" ends at x=110 (display); the new text starts right of it
    expect(input.origin[0]).toBeGreaterThan(110)
  })

  it('insert_text rejects an anchor that is not on the page', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('insert_text', { page: 1, text: 'x', anchor_text: 'nope' }),
    )
    expect(result.isError).toBe(true)
    expect(deps.insertText).not.toHaveBeenCalled()
  })
})

describe('add_form_mark', () => {
  it('centers a check mark against the right side of the anchor', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('add_form_mark', { page: 1, kind: 'check', anchor_text: 'Hello World' }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('unsaved')
    // anchor box is [0,700,110,712]: mark (22 pt) starts 4 pt right of it, centered on its middle
    expect(deps.addFormMark).toHaveBeenCalledWith(0, 'check', [114, 695, 136, 717])
    expect(deps.gotoPage).toHaveBeenCalledWith(1)
  })

  it('places a cross at explicit top-left coordinates with a custom size', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('add_form_mark', { page: 1, kind: 'cross', x: 100, y: 50, size: 20 }),
    )
    expect(result.isError).toBeUndefined()
    expect(deps.addFormMark).toHaveBeenCalledWith(0, 'cross', [100, 730, 120, 750])
  })

  it('rejects an unknown kind', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('add_form_mark', { page: 1, kind: 'tick', x: 0, y: 0 }),
    )
    expect(result.isError).toBe(true)
    expect(deps.addFormMark).not.toHaveBeenCalled()
  })

  it('rejects an anchor that is not on the page and a missing position', async () => {
    const deps = makeDeps()
    const missing = await executePdfTool(
      deps,
      call('add_form_mark', { page: 1, kind: 'check', anchor_text: 'nope' }),
    )
    expect(missing.isError).toBe(true)
    expect(missing.output).toContain('not found')
    const none = await executePdfTool(deps, call('add_form_mark', { page: 1, kind: 'check' }))
    expect(none.isError).toBe(true)
    expect(deps.addFormMark).not.toHaveBeenCalled()
  })

  it('refuses read-only documents', async () => {
    const deps = makeDeps({ readOnly: () => true })
    const result = await executePdfTool(
      deps,
      call('add_form_mark', { page: 1, kind: 'check', anchor_text: 'Hello World' }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('read-only')
    expect(deps.addFormMark).not.toHaveBeenCalled()
  })
})

describe('create_document', () => {
  it('defaults to a new PDF and reports the saved path', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('create_document', { title: 'Summary', content: '<h1>S</h1><p>Body</p>' }),
    )
    expect(result.isError).toBeUndefined()
    expect(deps.createDocument).toHaveBeenCalledWith({
      type: 'pdf',
      title: 'Summary',
      content: '<h1>S</h1><p>Body</p>',
    })
    expect(result.output).toContain('/tmp/out.pdf')
  })

  it('passes an explicit cross-type (docx) request through', async () => {
    const deps = makeDeps({ createDocument: vi.fn(async () => ({ ok: true })) })
    const result = await executePdfTool(
      deps,
      call('create_document', { type: 'docx', title: 'Report', content: '<p>x</p>' }),
    )
    expect(result.isError).toBeUndefined()
    expect(deps.createDocument).toHaveBeenCalledWith({
      type: 'docx',
      title: 'Report',
      content: '<p>x</p>',
    })
    expect(result.output).toContain('Report.docx')
  })

  it('rejects bad input without calling the bridge', async () => {
    const deps = makeDeps()
    for (const input of [
      { type: 'xlsx', title: 'T', content: '<p>x</p>' },
      { title: '  ', content: '<p>x</p>' },
      { title: 'T', content: '   ' },
    ]) {
      const result = await executePdfTool(deps, call('create_document', input))
      expect(result.isError).toBe(true)
    }
    expect(deps.createDocument).not.toHaveBeenCalled()
  })

  it('surfaces a main-process failure', async () => {
    const deps = makeDeps({ createDocument: vi.fn(async () => ({ ok: false, error: 'boom' })) })
    const result = await executePdfTool(
      deps,
      call('create_document', { title: 'T', content: '<p>x</p>' }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('boom')
  })

  it('rejects echoed tool-protocol output as content', async () => {
    const deps = makeDeps()
    for (const content of [
      '<tool_response>{"ok":true}</tool_response>',
      '{"index": 3, "type": "paragraph"}',
    ]) {
      const result = await executePdfTool(deps, call('create_document', { title: 'T', content }))
      expect(result.isError).toBe(true)
    }
    expect(deps.createDocument).not.toHaveBeenCalled()
  })
})

describe('move_page / reverse_pages', () => {
  it('move_page translates the original page number to its visible position', async () => {
    let order = [2, 0, 1]
    const deps = makeDeps({
      pageCount: () => 3,
      pageOrder: () => order,
      movePage: vi.fn((origIdx: number, to: number) => {
        const next = order.filter((i) => i !== origIdx)
        next.splice(to, 0, origIdx)
        order = next
      }),
    })
    const result = await executePdfTool(deps, call('move_page', { page: 1, to: 1 }))
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(deps.movePage).toHaveBeenCalledWith(0, 0)
    expect(result.output).toContain('Current order (original page numbers): 1, 3, 2')
  })

  it('move_page is a no-op for the current position and rejects bad targets', async () => {
    const deps = makeDeps()
    const same = await executePdfTool(deps, call('move_page', { page: 2, to: 2 }))
    expect(same.isError).toBeUndefined()
    expect(same.mutated).toBeUndefined()
    const out = await executePdfTool(deps, call('move_page', { page: 2, to: 3 }))
    expect(out.isError).toBe(true)
    expect(deps.movePage).not.toHaveBeenCalled()
    const ro = await executePdfTool(
      makeDeps({ readOnly: () => true }),
      call('move_page', { page: 2, to: 1 }),
    )
    expect(ro.isError).toBe(true)
  })

  it('reverse_pages flips the order and refuses single-page documents', async () => {
    let order = [0, 1]
    const deps = makeDeps({
      pageOrder: () => order,
      reversePages: vi.fn(() => {
        order = [...order].reverse()
      }),
    })
    const result = await executePdfTool(deps, call('reverse_pages', {}))
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('2, 1')
    const single = await executePdfTool(
      makeDeps({ pageCount: () => 1, pageOrder: () => [0] }),
      call('reverse_pages', {}),
    )
    expect(single.isError).toBe(true)
  })
})

describe('set_watermark / set_header_footer', () => {
  const HF = {
    headerLeft: 'ACME',
    headerCenter: '',
    headerRight: '',
    footerLeft: '',
    footerCenter: '',
    footerRight: '',
    pageNumber: true,
    startAt: 1,
    fontSize: 9,
    color: '#666666',
  }

  it('set_watermark fills defaults, keeps the session header/footer, and validates inputs', async () => {
    const deps = makeDeps({ stamps: () => ({ wm: null, hf: HF }) })
    const result = await executePdfTool(
      deps,
      call('set_watermark', { text: 'DRAFT', opacity: 0.3, color: '#FF0000' }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(deps.setStamps).toHaveBeenCalledWith({
      wm: { text: 'DRAFT', angle: 35, opacity: 0.3, color: '#ff0000', sizeRatio: 0.11 },
      hf: HF,
    })
    const bad = await executePdfTool(deps, call('set_watermark', { text: 'X', opacity: 2 }))
    expect(bad.isError).toBe(true)
    const badColor = await executePdfTool(deps, call('set_watermark', { text: 'X', color: 'red' }))
    expect(badColor.isError).toBe(true)
  })

  it('set_watermark with empty text removes only the watermark', async () => {
    const wm = { text: 'OLD', angle: 35, opacity: 0.18, color: '#d0342c', sizeRatio: 0.11 }
    const both = makeDeps({ stamps: () => ({ wm, hf: HF }) })
    await executePdfTool(both, call('set_watermark', { text: '' }))
    expect(both.setStamps).toHaveBeenCalledWith({ wm: null, hf: HF })
    const only = makeDeps({ stamps: () => ({ wm, hf: null }) })
    await executePdfTool(only, call('set_watermark', { text: '' }))
    expect(only.setStamps).toHaveBeenCalledWith(null)
    const none = await executePdfTool(makeDeps(), call('set_watermark', { text: '' }))
    expect(none.isError).toBe(true)
  })

  it('set_header_footer builds the config with defaults and removes on an all-empty call', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('set_header_footer', { header_left: 'ACME', page_number: true }),
    )
    expect(result.isError).toBeUndefined()
    expect(deps.setStamps).toHaveBeenCalledWith({ wm: null, hf: HF })
    expect(result.output).toContain('page numbers')

    const clear = makeDeps({ stamps: () => ({ wm: null, hf: HF }) })
    const removed = await executePdfTool(clear, call('set_header_footer', {}))
    expect(removed.mutated).toBe(true)
    expect(clear.setStamps).toHaveBeenCalledWith(null)
    const nothing = await executePdfTool(makeDeps(), call('set_header_footer', {}))
    expect(nothing.isError).toBe(true)
  })
})

describe('set_metadata', () => {
  it('merges the given fields over the current properties', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(deps, call('set_metadata', { subject: 'Q3 numbers' }))
    expect(result.mutated).toBe(true)
    expect(deps.setMetadata).toHaveBeenCalledWith({
      title: 'Report',
      author: 'Ann',
      subject: 'Q3 numbers',
      keywords: '',
    })
    expect(result.output).toContain('title: "Report"')
    expect(result.output).toContain('subject: "Q3 numbers"')
    expect(result.output).toContain('on save')
  })

  it('clears a field with "" and reports the current values on a no-op', async () => {
    const deps = makeDeps()
    const cleared = await executePdfTool(deps, call('set_metadata', { author: '' }))
    expect(deps.setMetadata).toHaveBeenCalledWith(expect.objectContaining({ author: '' }))
    expect(cleared.output).toContain('author: (empty)')
    const noop = await executePdfTool(deps, call('set_metadata', { title: 'Report' }))
    expect(noop.mutated).toBeUndefined()
    expect(noop.isError).toBeUndefined()
    expect(noop.output).toContain('unchanged')
    expect(noop.output).toContain('author: "Ann"')
    expect(deps.setMetadata).toHaveBeenCalledTimes(1)
    const ro = await executePdfTool(
      makeDeps({ readOnly: () => true }),
      call('set_metadata', { title: 'New' }),
    )
    expect(ro.isError).toBe(true)
  })
})

describe('inserted text lifecycle', () => {
  const block = (over: Partial<LocalTextInsert['input']> = {}): LocalTextInsert => ({
    id: 'dabc',
    input: {
      pageIndex: 1,
      origin: [100, 700],
      text: 'Hello\nWorld',
      fontSize: 14,
      color: [0, 0, 0],
      ...over,
    },
  })

  it('list_inserted_text reports id, page, displayed top-left and a preview', async () => {
    const centered = block({
      id: 'dcen',
      origin: [300, 500],
      lineXOffsets: [-40, -20],
      align: 'center',
    })
    const deps = makeDeps({ textInserts: () => [block(), { ...centered, id: 'dcen' }] })
    const all = await executePdfTool(deps, call('list_inserted_text', {}))
    expect(all.isError).toBeUndefined()
    expect(all.output).toContain(
      'Tdabc: page 2, x=100, y=86, 14 pt, left, 2 line(s): "Hello ⏎ World"',
    )
    expect(all.output).toContain('Tdcen: page 2, x=260, y=286')
    const none = await executePdfTool(deps, call('list_inserted_text', { page: 1 }))
    expect(none.output).toContain('No pending inserted text blocks on page 1')
    expect(none.output).toContain('edit_text')
  })

  it('edit_inserted_text re-commits the block with the merged fields', async () => {
    const deps = makeDeps({ textInserts: () => [block()] })
    const result = await executePdfTool(
      deps,
      call('edit_inserted_text', { id: 'Tdabc', text: 'Bye', color: '#ff0000' }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('unsaved')
    expect(deps.updateTextInsert).toHaveBeenCalledWith('dabc', {
      text: 'Bye',
      color: [255, 0, 0],
      lineXOffsets: [0],
    })
    expect(deps.gotoPage).toHaveBeenCalledWith(2)
  })

  it('edit_inserted_text keeps the offsets on a color-only change and remeasures with the face', async () => {
    const bold = block({ bold: true, align: 'center', lineXOffsets: [-30, -10] })
    const deps = makeDeps({ textInserts: () => [bold] })
    await executePdfTool(deps, call('edit_inserted_text', { id: 'dabc', color: '#0000ff' }))
    expect(deps.updateTextInsert).toHaveBeenLastCalledWith('dabc', { color: [0, 0, 255] })

    await executePdfTool(deps, call('edit_inserted_text', { id: 'dabc', text: 'aaaa\naa' }))
    const family = getComputedStyle(document.body).fontFamily
    expect(deps.updateTextInsert).toHaveBeenLastCalledWith('dabc', {
      text: 'aaaa\naa',
      lineXOffsets: ['aaaa', 'aa'].map((l) => -measurePt(l, 14, family, 'bold') / 2),
    })

    await executePdfTool(deps, call('edit_inserted_text', { id: 'dabc', font_size: 20 }))
    const last = vi.mocked(deps.updateTextInsert).mock.lastCall![1]
    expect(last.fontSize).toBe(20)
    expect(last.lineLeading).toBe(24)
    expect(last.lineXOffsets).toEqual(
      ['Hello', 'World'].map((l) => -measurePt(l, 20, family, 'bold') / 2),
    )
  })

  it('edit_inserted_text rejects empty edits, empty text and bad values', async () => {
    const deps = makeDeps({ textInserts: () => [block()] })
    const nothing = await executePdfTool(deps, call('edit_inserted_text', { id: 'dabc' }))
    expect(nothing.isError).toBe(true)
    const empty = await executePdfTool(deps, call('edit_inserted_text', { id: 'dabc', text: ' ' }))
    expect(empty.isError).toBe(true)
    const color = await executePdfTool(
      deps,
      call('edit_inserted_text', { id: 'dabc', color: 'red' }),
    )
    expect(color.isError).toBe(true)
    const align = await executePdfTool(
      deps,
      call('edit_inserted_text', { id: 'dabc', align: 'justify' }),
    )
    expect(align.isError).toBe(true)
    expect(deps.updateTextInsert).not.toHaveBeenCalled()
  })

  it('move_inserted_text takes an absolute top-left or a relative offset', async () => {
    const deps = makeDeps({ textInserts: () => [block()] })
    const abs = await executePdfTool(
      deps,
      call('move_inserted_text', { id: 'Tdabc', x: 150, y: 100 }),
    )
    expect(abs.isError).toBeUndefined()
    expect(abs.mutated).toBe(true)
    expect(abs.output).toContain('x=150, y=100')
    // displayed top-left (100, 86) → (150, 100): +50 right, +14 down = baseline y drops by 14
    expect(deps.moveTextInsert).toHaveBeenCalledWith('dabc', [150, 686])
    const rel = await executePdfTool(deps, call('move_inserted_text', { id: 'dabc', dy: 20 }))
    expect(rel.output).toContain('x=100, y=106')
    expect(deps.moveTextInsert).toHaveBeenLastCalledWith('dabc', [100, 680])
  })

  it('move_inserted_text rotates the delta with the page and validates its arguments', async () => {
    const deps = makeDeps({
      textInserts: () => [block()],
      pageGeom: () => ({ pw: 600, ph: 800, rot: 90 }),
    })
    await executePdfTool(deps, call('move_inserted_text', { id: 'dabc', dx: 10, dy: 0 }))
    // rot 90: display x runs along PDF y
    expect(deps.moveTextInsert).toHaveBeenCalledWith('dabc', [100, 710])
    const none = await executePdfTool(deps, call('move_inserted_text', { id: 'dabc' }))
    expect(none.isError).toBe(true)
    const both = await executePdfTool(deps, call('move_inserted_text', { id: 'dabc', x: 1, dx: 1 }))
    expect(both.isError).toBe(true)
    const same = await executePdfTool(deps, call('move_inserted_text', { id: 'dabc', dx: 0 }))
    expect(same.isError).toBeUndefined()
    expect(same.mutated).toBeUndefined()
  })

  it('delete_inserted_text removes the block', async () => {
    const deps = makeDeps({ textInserts: () => [block()] })
    const result = await executePdfTool(deps, call('delete_inserted_text', { id: 'Tdabc' }))
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('unsaved')
    expect(deps.deleteTextInsert).toHaveBeenCalledWith('dabc')
  })

  it('unknown ids point at list_inserted_text and read-only documents refuse', async () => {
    const deps = makeDeps({ textInserts: () => [block()] })
    for (const name of ['edit_inserted_text', 'move_inserted_text', 'delete_inserted_text']) {
      const args =
        name === 'edit_inserted_text'
          ? { text: 'x' }
          : name === 'move_inserted_text'
            ? { dx: 1 }
            : {}
      const missing = await executePdfTool(deps, call(name, { id: 'Tnope', ...args }))
      expect(missing.isError).toBe(true)
      expect(missing.output).toContain('list_inserted_text')
      const ro = await executePdfTool(
        makeDeps({ readOnly: () => true, textInserts: () => [block()] }),
        call(name, { id: 'Tdabc', ...args }),
      )
      expect(ro.isError).toBe(true)
      expect(ro.output).toContain('read-only')
    }
    expect(deps.updateTextInsert).not.toHaveBeenCalled()
    expect(deps.moveTextInsert).not.toHaveBeenCalled()
    expect(deps.deleteTextInsert).not.toHaveBeenCalled()
  })
})

describe('file-level page operations', () => {
  const fourPages = (over: Partial<PdfAiDeps> = {}) =>
    makeDeps({ pageCount: () => 4, pageOrder: () => [0, 1, 2, 3], ...over })

  it('insert_blank_page runs after the user confirms and tells the model to re-read', async () => {
    const deps = fourPages({ insertBlankPage: vi.fn(async () => ({ ok: true, pageCount: 5 })) })
    const result = await executePdfTool(deps, call('insert_blank_page', { after_page: 2 }))
    expect(deps.confirmFileOp).toHaveBeenCalledTimes(1)
    expect(deps.insertBlankPage).toHaveBeenCalledWith(1)
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('page 3')
    expect(result.output).toContain('5 pages')
    expect(result.output).toContain('re-read')
  })

  it('insert_blank_page 0 inserts at the front', async () => {
    const deps = fourPages()
    const result = await executePdfTool(deps, call('insert_blank_page', { after_page: 0 }))
    expect(deps.insertBlankPage).toHaveBeenCalledWith(-1)
    expect(result.output).toContain('page 1')
  })

  it('a declined card is a plain non-error output and nothing runs', async () => {
    const deps = fourPages({ confirmFileOp: vi.fn(async () => false) })
    const result = await executePdfTool(deps, call('insert_blank_page', { after_page: 1 }))
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBeUndefined()
    expect(result.output).toContain('declined')
    expect(deps.insertBlankPage).not.toHaveBeenCalled()
  })

  it('stopping the run while the card is open reports the stop', async () => {
    const ctl = new AbortController()
    const deps = fourPages({
      confirmFileOp: vi.fn(async (_req, signal) => {
        ctl.abort()
        return !signal?.aborted
      }),
    })
    const result = await executePdfTool(deps, call('set_page_size', { preset: 'A4' }), ctl.signal)
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('stopped')
    expect(deps.setPageSize).not.toHaveBeenCalled()
  })

  it('aborts when the pages change while the card is open', async () => {
    let order = [0, 1, 2, 3]
    const deps = fourPages({
      pageOrder: () => order,
      confirmFileOp: vi.fn(async () => {
        order = [0, 2, 1, 3]
        return true
      }),
    })
    for (const c of [
      call('insert_blank_page', { after_page: 2 }),
      call('crop_pages', { pages: '2', left: 0.1, top: 0, right: 0.9, bottom: 1 }),
      call('extract_pages', { pages: '2-3' }),
      call('replace_pages', { pages: '2' }),
    ]) {
      order = [0, 1, 2, 3]
      const result = await executePdfTool(deps, c)
      expect(result.output, c.name).toContain('changed while the confirmation card was open')
      expect(result.isError, c.name).toBeUndefined()
      expect(result.mutated, c.name).toBeUndefined()
    }
    order = [0, 1, 2, 3]
    const deleted = fourPages({
      pageOrder: () => order,
      confirmFileOp: vi.fn(async () => {
        order = [0, 1, 2]
        return true
      }),
    })
    const split = await executePdfTool(deleted, call('split_pdf', { pages_per_file: 2 }))
    expect(split.output).toContain('changed while the confirmation card was open')
    order = [0, 1, 2, 3]
    const gone = await executePdfTool(deleted, call('extract_pages', { pages: '4' }))
    expect(gone.output).toContain('changed while the confirmation card was open')
    expect(deps.insertBlankPage).not.toHaveBeenCalled()
    expect(deps.cropPages).not.toHaveBeenCalled()
    expect(deps.extractPages).not.toHaveBeenCalled()
    expect(deps.replacePages).not.toHaveBeenCalled()
    expect(deleted.splitPdf).not.toHaveBeenCalled()
    expect(deleted.extractPages).not.toHaveBeenCalled()
  })

  it('translates original page numbers through the visible order', async () => {
    const deps = fourPages({ pageOrder: () => [2, 0, 1, 3] })
    await executePdfTool(deps, call('insert_blank_page', { after_page: 1 }))
    expect(deps.insertBlankPage).toHaveBeenCalledWith(1)
    await executePdfTool(
      deps,
      call('crop_pages', { pages: '1,3', left: 0.1, top: 0, right: 0.9, bottom: 1 }),
    )
    expect(deps.cropPages).toHaveBeenCalledWith([0, 1], { l: 0.1, t: 0, r: 0.9, b: 1 })
    const gone = await executePdfTool(
      fourPages({ pageOrder: () => [0, 1, 3] }),
      call('extract_pages', { pages: '3' }),
    )
    expect(gone.isError).toBe(true)
    expect(gone.output).toContain('deleted')
  })

  it('refuses read-only documents before asking for confirmation', async () => {
    const deps = fourPages({ readOnly: () => true })
    for (const c of [
      call('insert_blank_page', { after_page: 1 }),
      call('set_page_size', { preset: 'A4' }),
      call('crop_pages', { pages: 'all', left: 0.1, top: 0.1, right: 0.9, bottom: 0.9 }),
      call('extract_pages', { pages: '1-2' }),
      call('split_pdf', { pages_per_file: 2 }),
      call('split_pages', { per_page: 4 }),
      call('merge_pages', { per_sheet: 2 }),
      call('replace_pages', { pages: '2' }),
    ]) {
      const result = await executePdfTool(deps, c)
      expect(result.isError, c.name).toBe(true)
    }
    expect(deps.confirmFileOp).not.toHaveBeenCalled()
  })

  it('validates parameters without showing the card', async () => {
    const deps = fourPages()
    const bad = [
      call('insert_blank_page', { after_page: 5 }),
      call('set_page_size', {}),
      call('set_page_size', { preset: 'B5' }),
      call('set_page_size', { width: 500 }),
      call('set_page_size', { width: 10, height: 500 }),
      call('crop_pages', { pages: 'all', left: 0, top: 0, right: 1, bottom: 1 }),
      call('crop_pages', { pages: 'all', left: 0.6, top: 0, right: 0.4, bottom: 1 }),
      call('crop_pages', { pages: 'all', left: 0, top: 0, right: 1.2, bottom: 1 }),
      call('crop_pages', { pages: '1-9', left: 0.1, top: 0, right: 1, bottom: 1 }),
      call('extract_pages', { pages: '' }),
      call('extract_pages', { pages: '3-1' }),
      call('split_pdf', { pages_per_file: 0 }),
      call('split_pdf', { pages_per_file: 4 }),
      call('split_pages', { per_page: 3 }),
      call('merge_pages', { per_sheet: 1 }),
      call('merge_pages', { per_sheet: 17 }),
      call('merge_pages', { per_sheet: 2, direction: 'diagonal' }),
      call('replace_pages', { pages: 'x' }),
    ]
    for (const c of bad) {
      const result = await executePdfTool(deps, c)
      expect(result.isError, `${c.name} ${JSON.stringify(c.input)}`).toBe(true)
    }
    expect(deps.confirmFileOp).not.toHaveBeenCalled()
  })

  it('set_page_size takes presets case-insensitively or explicit points', async () => {
    const deps = fourPages()
    const a4 = await executePdfTool(deps, call('set_page_size', { preset: 'a4' }))
    expect(deps.setPageSize).toHaveBeenCalledWith(595, 842)
    expect(a4.output).toContain('A4')
    expect(a4.mutated).toBe(true)
    await executePdfTool(deps, call('set_page_size', { width: 500, height: 700 }))
    expect(deps.setPageSize).toHaveBeenLastCalledWith(500, 700)
  })

  it('crop_pages "all" covers every visible page and reports the kept area', async () => {
    const deps = fourPages()
    const result = await executePdfTool(
      deps,
      call('crop_pages', { pages: 'all', left: 0.1, top: 0.2, right: 0.9, bottom: 1 }),
    )
    expect(deps.cropPages).toHaveBeenCalledWith([0, 1, 2, 3], { l: 0.1, t: 0.2, r: 0.9, b: 1 })
    expect(result.output).toContain('1-4')
    expect(result.output).toContain('left 10%')
    expect(result.mutated).toBe(true)
  })

  it('new-file tools report the saved path and picker cancellations', async () => {
    const deps = fourPages({
      splitPdf: vi.fn(async () => ({ ok: true, canceled: true, flushed: false })),
      replacePages: vi.fn(async () => ({ ok: true, canceled: true, flushed: false })),
    })
    const extracted = await executePdfTool(deps, call('extract_pages', { pages: '1-2' }))
    expect(deps.extractPages).toHaveBeenCalledWith([0, 1])
    expect(extracted.output).toContain('/tmp/test-p1.pdf')
    expect(extracted.output).toContain('new tab')
    expect(extracted.mutated).toBe(true)

    const split = await executePdfTool(deps, call('split_pdf', { pages_per_file: 2 }))
    expect(split.isError).toBeUndefined()
    expect(split.mutated).toBeUndefined()
    expect(split.output).toContain('canceled')

    const replaced = await executePdfTool(deps, call('replace_pages', { pages: '2-3' }))
    expect(deps.replacePages).toHaveBeenCalledWith([1, 2])
    expect(replaced.isError).toBeUndefined()
    expect(replaced.output).toContain('canceled')
    expect(replaced.output).toContain('nothing was changed')
    expect(replaced.output).not.toContain('saved')

    await executePdfTool(deps, call('merge_pages', { per_sheet: 2 }))
    expect(deps.mergePages).toHaveBeenCalledWith(2, 'vertical', false)
    await executePdfTool(
      deps,
      call('merge_pages', { per_sheet: 4, direction: 'horizontal', separator: true }),
    )
    expect(deps.mergePages).toHaveBeenLastCalledWith(4, 'horizontal', true)

    const pieces = await executePdfTool(deps, call('split_pages', { per_page: 9 }))
    expect(deps.splitPages).toHaveBeenCalledWith(9)
    expect(pieces.output).toContain('/tmp/test-split.pdf')
  })

  it('a picker canceled after the flush admits the document was saved', async () => {
    const deps = fourPages({
      splitPdf: vi.fn(async () => ({ ok: true, canceled: true, flushed: true })),
      replacePages: vi.fn(async () => ({ ok: true, canceled: true, flushed: true })),
    })
    for (const c of [
      call('replace_pages', { pages: '2' }),
      call('split_pdf', { pages_per_file: 2 }),
    ]) {
      const result = await executePdfTool(deps, c)
      expect(result.isError, c.name).toBeUndefined()
      expect(result.mutated, c.name).toBeUndefined()
      expect(result.output, c.name).toContain('canceled')
      expect(result.output, c.name).toContain('already been saved')
      expect(result.output, c.name).toContain('undo history is cleared')
      expect(result.output, c.name).toContain('Re-read')
      expect(result.output, c.name).not.toContain('nothing was changed')
    }
  })

  it('surfaces a failed operation as a tool error', async () => {
    const deps = fourPages({ splitPdf: vi.fn(async () => ({ ok: false, error: 'boom' })) })
    const written = await executePdfTool(fourPages(), call('split_pdf', { pages_per_file: 3 }))
    expect(written.output).toContain('/tmp/split')
    const result = await executePdfTool(deps, call('split_pdf', { pages_per_file: 3 }))
    expect(result.isError).toBe(true)
    expect(result.output).toContain('boom')
  })
})
