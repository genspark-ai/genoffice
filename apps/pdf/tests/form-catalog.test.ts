import { describe, expect, it } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { VISUAL_SIGNATURE_CONTENT_PREFIX } from '../src/shared/ipc'
import {
  buildFormCatalog,
  documentFormFeatures,
  hasXfaMarker,
  visibleFormWidgets,
  widgetsFromAnnotations,
  type RawFormAnnotation,
} from '../src/renderer/form-catalog'

function fakeDoc(pages: RawFormAnnotation[][]): PDFDocumentProxy {
  return {
    numPages: pages.length,
    getPage: async (page: number) => ({
      getAnnotations: async () => pages[page - 1] ?? [],
    }),
  } as unknown as PDFDocumentProxy
}

describe('form catalog', () => {
  it('builds one document catalog with page order and aggregated radio options', async () => {
    const catalog = await buildFormCatalog(
      fakeDoc([
        [
          {
            id: 'name',
            subtype: 'Widget',
            fieldType: 'Tx',
            fieldName: 'name',
            fieldValue: 'Alice',
            rect: [10, 100, 100, 120],
            required: true,
          },
          {
            id: 'red',
            subtype: 'Widget',
            fieldType: 'Btn',
            fieldName: 'color',
            fieldValue: '',
            buttonValue: 'red',
            radioButton: true,
            rect: [10, 50, 20, 60],
          },
          {
            id: 'blue',
            subtype: 'Widget',
            fieldType: 'Btn',
            fieldName: 'color',
            fieldValue: '',
            buttonValue: 'blue',
            radioButton: true,
            rect: [30, 50, 40, 60],
          },
        ],
        [
          {
            id: 'sig',
            subtype: 'Widget',
            fieldType: 'Sig',
            fieldName: 'signature',
            fieldValue: null,
            rect: [20, 20, 200, 50],
          },
        ],
      ]),
    )

    expect(catalog.widgets.map((widget) => widget.id)).toEqual(['name', 'red', 'blue', 'sig'])
    expect(catalog.byPage.get(1)?.[0]).toMatchObject({
      kind: 'signature',
      fieldName: 'signature',
      pageIndex: 1,
    })
    expect(catalog.fields.get('name')).toMatchObject({
      kind: 'text',
      value: 'Alice',
      required: true,
    })
    expect(catalog.fields.get('color')?.options).toEqual(['red', 'blue'])
  })

  it('filters deleted pages and follows the current page order', async () => {
    const catalog = await buildFormCatalog(
      fakeDoc([
        [
          {
            id: 'page-1',
            subtype: 'Widget',
            fieldType: 'Tx',
            fieldName: 'first',
            rect: [10, 100, 100, 120],
          },
        ],
        [
          {
            id: 'page-2',
            subtype: 'Widget',
            fieldType: 'Tx',
            fieldName: 'second',
            rect: [10, 100, 100, 120],
          },
        ],
      ]),
    )

    expect(visibleFormWidgets(catalog, [1]).map((widget) => widget.fieldName)).toEqual(['second'])
    expect(visibleFormWidgets(catalog, [1, 0]).map((widget) => widget.fieldName)).toEqual([
      'second',
      'first',
    ])
  })

  it('parses signature widgets but still skips push buttons', () => {
    const widgets = widgetsFromAnnotations([
      {
        id: 'sig',
        subtype: 'Widget',
        fieldType: 'Sig',
        fieldName: 'signature',
        rect: [0, 0, 100, 20],
      },
      {
        id: 'submit',
        subtype: 'Widget',
        fieldType: 'Btn',
        fieldName: 'submit',
        pushButton: true,
        rect: [0, 30, 100, 50],
      },
    ])

    expect(widgets).toHaveLength(1)
    expect(widgets[0]?.kind).toBe('signature')
  })

  it('recognizes a saved visual signature inside a signature field', async () => {
    const catalog = await buildFormCatalog(
      fakeDoc([
        [
          {
            id: 'sig',
            subtype: 'Widget',
            fieldType: 'Sig',
            fieldName: 'signature',
            rect: [20, 20, 200, 60],
          },
          {
            subtype: 'Ink',
            rect: [35, 25, 185, 55],
          },
        ],
      ]),
    )

    expect(catalog.widgets[0]).toMatchObject({ kind: 'signature', signed: true })
  })

  it('prefers persisted field association over annotation position', async () => {
    const catalog = await buildFormCatalog(
      fakeDoc([
        [
          {
            id: 'sig',
            subtype: 'Widget',
            fieldType: 'Sig',
            fieldName: 'signature',
            rect: [20, 20, 200, 60],
          },
          {
            id: 'other-sig',
            subtype: 'Widget',
            fieldType: 'Sig',
            fieldName: 'otherSignature',
            rect: [300, 300, 400, 350],
          },
          {
            subtype: 'Stamp',
            rect: [300, 300, 400, 350],
            contentsObj: { str: `${VISUAL_SIGNATURE_CONTENT_PREFIX}signature` },
          },
        ],
      ]),
    )

    expect(catalog.widgets.find((widget) => widget.fieldName === 'signature')?.signed).toBe(true)
    expect(catalog.widgets.find((widget) => widget.fieldName === 'otherSignature')?.signed).toBe(
      false,
    )
  })

  it('detects an XFA catalog marker without decoding the full PDF', () => {
    expect(hasXfaMarker(new TextEncoder().encode('<< /AcroForm << /XFA 12 0 R >> >>'))).toBe(true)
    expect(hasXfaMarker(new TextEncoder().encode('<< /AcroForm 12 0 R >>'))).toBe(false)
  })

  it('uses document metadata for compressed XFA and encryption flags', () => {
    const bytes = new TextEncoder().encode('compressed catalog data')
    expect(
      documentFormFeatures({ IsXFAPresent: true, EncryptFilterName: 'Standard' }, bytes),
    ).toEqual({ hasXfa: true, encrypted: true })
    expect(documentFormFeatures({}, bytes)).toEqual({ hasXfa: false, encrypted: false })
  })
})
