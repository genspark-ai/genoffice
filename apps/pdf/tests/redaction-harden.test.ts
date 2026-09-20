/**
 * PDF redaction hardening: index-aware validation messages plus
 * RedactionLayer cancel/tiny-drag/key/label behavior.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { redactPdf, validateRedactionRegions } from '../src/main/redaction'
import { RedactionLayer } from '../src/renderer/RedactionLayer'

async function fixture(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([300, 200])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('PUBLIC', { x: 20, y: 90, size: 18, font })
  return doc.save({ useObjectStreams: false })
}

describe('validateRedactionRegions messages', () => {
  it('names the operation for a non-object region', () => {
    expect(() => validateRedactionRegions([null])).toThrow(/redaction 1.*invalid redaction region/)
  })

  it('names the operation for a non-integer page index', () => {
    expect(() =>
      validateRedactionRegions([
        { pageIndex: 0, rect: [1, 1, 2, 2] },
        { pageIndex: 1.5, rect: [1, 1, 2, 2] },
      ]),
    ).toThrow(/redaction 2.*page index must be an integer/)
  })

  it('rejects a negative page index before touching bytes', () => {
    expect(() => validateRedactionRegions([{ pageIndex: -1, rect: [1, 1, 2, 2] }])).toThrow(
      /redaction 1.*page index -1 is out of range/,
    )
  })

  it('names operation, page, and rect for a malformed rectangle', () => {
    expect(() =>
      validateRedactionRegions([
        { pageIndex: 0, rect: [1, 1, 2, 2] },
        { pageIndex: 0, rect: [1, 1, 2] },
      ]),
    ).toThrow(/redaction 2 on page 1.*rectangle must contain four finite coordinates.*\[1,1,2\]/)
  })

  it('names operation and page for an empty rectangle', () => {
    expect(() => validateRedactionRegions([{ pageIndex: 0, rect: [1, 1, 1, 2] }])).toThrow(
      /redaction 1 on page 1.*rectangle must be nonempty/,
    )
  })

  it('names operation and page for coordinates outside PDFium range', () => {
    expect(() => validateRedactionRegions([{ pageIndex: 0, rect: [0, 0, 1e39, 1] }])).toThrow(
      /redaction 1 on page 1.*rectangle is outside PDFium coordinate range/,
    )
  })

  it('rejects a trailing out-of-range page with its operation number', async () => {
    const before = await fixture()
    await expect(
      redactPdf(before, [
        { pageIndex: 0, rect: [1, 1, 2, 2] },
        { pageIndex: 9, rect: [1, 1, 2, 2] },
      ]),
    ).rejects.toThrow(/redaction 2.*page index 9 is out of range/)
  })

  it('rejects a negative page through redactPdf fail-closed', async () => {
    const before = await fixture()
    await expect(redactPdf(before, [{ pageIndex: -1, rect: [1, 1, 2, 2] }])).rejects.toThrow(
      /page index/,
    )
  })
})

// Renderer harness mirrors dialog-a11y.test.ts: real mount with createRoot + act.
let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
})

const geom = { pw: 200, ph: 200, rot: 0 }

function stubLayer(layer: HTMLElement) {
  ;(layer as unknown as Record<string, unknown>).setPointerCapture = () => {}
  layer.getBoundingClientRect = () =>
    ({ left: 0, top: 0, right: 200, bottom: 200, width: 200, height: 200, x: 0, y: 0 }) as DOMRect
}

async function renderLayer(props: Record<string, unknown>) {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(
      createElement(RedactionLayer, {
        active: true,
        geom,
        scale: 1,
        pageWidth: 200,
        pageHeight: 200,
        marks: [],
        markLabel: 'Redact area',
        onCommit: () => {},
        ...props,
      }),
    )
    await Promise.resolve()
  })
  const layer = container.querySelector<HTMLElement>('.pdf-redaction-layer')!
  stubLayer(layer)
  return { container, layer }
}

function firePointer(layer: HTMLElement, type: string, clientX: number, clientY: number) {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    button: 0,
    clientX,
    clientY,
  })
  ;(event as unknown as Record<string, unknown>).pointerId = 1
  layer.dispatchEvent(event)
}

describe('RedactionLayer behavior', () => {
  it('labels pending marks with the passed redact string', async () => {
    const { container } = await renderLayer({
      marks: [{ id: 'a', pageIndex: 0, rect: [10, 10, 20, 20] }],
      markLabel: 'Redact area',
    })
    const marks = container.querySelectorAll('.pdf-redaction-mark')
    expect(marks).toHaveLength(1)
    expect(marks[0]!.getAttribute('aria-label')).toBe('Redact area')
    expect(container.innerHTML).not.toContain('Pending redaction')
  })

  it('keeps mark identity by id when order changes', async () => {
    const first = [
      { id: 'a', pageIndex: 0, rect: [10, 10, 20, 20] },
      { id: 'b', pageIndex: 0, rect: [30, 30, 40, 40] },
    ]
    const { container } = await renderLayer({ marks: first })
    const before = [...container.querySelectorAll('.pdf-redaction-mark')]
    expect(before).toHaveLength(2)
    await act(async () => {
      root!.render(
        createElement(RedactionLayer, {
          active: true,
          geom,
          scale: 1,
          pageWidth: 200,
          pageHeight: 200,
          marks: [first[1], first[0]],
          markLabel: 'Redact area',
          onCommit: () => {},
        }),
      )
      await Promise.resolve()
    })
    const after = [...container.querySelectorAll('.pdf-redaction-mark')]
    expect(after).toHaveLength(2)
    expect(after[0]).toBe(before[1])
    expect(after[1]).toBe(before[0])
  })

  it('commits a drag above the minimum size', async () => {
    const onCommit = vi.fn()
    const { layer } = await renderLayer({ onCommit })
    await act(async () => {
      firePointer(layer, 'pointerdown', 10, 10)
      await Promise.resolve()
    })
    await act(async () => {
      firePointer(layer, 'pointermove', 50, 50)
      await Promise.resolve()
    })
    await act(async () => {
      firePointer(layer, 'pointerup', 50, 50)
      await Promise.resolve()
    })
    expect(onCommit).toHaveBeenCalledTimes(1)
    const rect = onCommit.mock.calls[0]![0] as [number, number, number, number]
    expect(rect[2] - rect[0]).toBeGreaterThanOrEqual(3)
    expect(rect[3] - rect[1]).toBeGreaterThanOrEqual(3)
  })

  it('discards a cancelled drag instead of committing it', async () => {
    const onCommit = vi.fn()
    const onTooSmall = vi.fn()
    const { container, layer } = await renderLayer({ onCommit, onTooSmall })
    await act(async () => {
      firePointer(layer, 'pointerdown', 10, 10)
      await Promise.resolve()
    })
    await act(async () => {
      firePointer(layer, 'pointermove', 60, 60)
      await Promise.resolve()
    })
    await act(async () => {
      firePointer(layer, 'pointercancel', 60, 60)
      await Promise.resolve()
    })
    expect(onCommit).not.toHaveBeenCalled()
    expect(onTooSmall).not.toHaveBeenCalled()
    expect(container.querySelectorAll('.pdf-redaction-mark')).toHaveLength(0)
  })

  it('reports a below-minimum drag instead of silently dropping it', async () => {
    const onCommit = vi.fn()
    const onTooSmall = vi.fn()
    const { container, layer } = await renderLayer({ onCommit, onTooSmall })
    await act(async () => {
      firePointer(layer, 'pointerdown', 10, 10)
      await Promise.resolve()
    })
    await act(async () => {
      firePointer(layer, 'pointermove', 11, 11)
      await Promise.resolve()
    })
    await act(async () => {
      firePointer(layer, 'pointerup', 11, 11)
      await Promise.resolve()
    })
    expect(onCommit).not.toHaveBeenCalled()
    expect(onTooSmall).toHaveBeenCalledTimes(1)
    expect(container.querySelectorAll('.pdf-redaction-mark')).toHaveLength(0)
  })
})
