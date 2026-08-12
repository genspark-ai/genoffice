import { describe, expect, it, vi } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { AGENT_TOOLS, executePdfTool, type PdfAiDeps } from '../src/renderer/ai/tools'
import type { SearchIndex } from '../src/renderer/search'
import type { FormValueInput, PageImageRef } from '../src/shared/ipc'

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
    outline: () => null,
    searchIndex: () => Promise.resolve(INDEX),
    isDeleted: () => false,
    gotoPage: vi.fn(() => true),
    addMarkup: vi.fn(),
    formEdits: () => new Map<string, FormValueInput>(),
    applyFormEdit: vi.fn(),
    rotatePage: vi.fn(),
    deletePage: vi.fn(() => true),
    editText: vi.fn(async () => null),
    editFonts: () => ['arial', 'times', 'courier'],
    pageGeom: () => ({ pw: 600, ph: 800, rot: 0 }),
    listImages: vi.fn(async () => PAGE_IMAGES),
    isImageClaimed: () => false,
    insertImage: vi.fn(),
    transformImage: vi.fn(),
    replaceImage: vi.fn(),
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
    expect(deps.addMarkup).toHaveBeenCalledWith('highlight', 1, expect.any(Array))
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

  it('rejects empty replacements, bad colors, and read-only documents', async () => {
    const empty = await executePdfTool(
      makeDeps(),
      call('edit_text', { page: 1, old_text: 'Hello', new_text: '  ' }),
    )
    expect(empty.isError).toBe(true)

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
    expect(deps.rotatePage).toHaveBeenCalledWith(0, -90)
    await executePdfTool(deps, call('rotate_page', { page: 2, direction: 'right' }))
    expect(deps.rotatePage).toHaveBeenCalledWith(1, 90)
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
      expect(deps.rotatePage).not.toHaveBeenCalled()
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
