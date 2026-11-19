import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { PdfThumb } from '../src/renderer/PdfThumb'

interface FakeRenderTask {
  promise: Promise<void>
  cancel: ReturnType<typeof vi.fn>
  resolve: () => void
  reject: (reason?: Error) => void
}

function makeRenderTask(): FakeRenderTask {
  let resolve!: () => void
  let reject!: (reason?: Error) => void
  const promise = new Promise<void>((res, rej) => {
    resolve = res
    reject = rej
  })
  return {
    promise,
    resolve,
    reject,
    cancel: vi.fn(() => reject(new Error('render cancelled'))),
  }
}

function makePdf() {
  const tasks: FakeRenderTask[] = []
  const page = {
    rotate: 0,
    getViewport: vi.fn(({ scale, rotation }: { scale: number; rotation: number }) => ({
      width: 200 * scale,
      height: 100 * scale,
      rotation,
    })),
    render: vi.fn(() => {
      const task = makeRenderTask()
      tasks.push(task)
      return task
    }),
  }
  const doc = {
    getPage: vi.fn(async () => page),
  } as unknown as PDFDocumentProxy
  return { doc, page, tasks }
}

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  Object.defineProperty(window, 'devicePixelRatio', { configurable: true, value: 1 })
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
})

async function renderThumb(
  doc: PDFDocumentProxy,
  props: { visible: boolean; rotationDelta?: number; rasterW?: number },
): Promise<HTMLCanvasElement> {
  if (!container) {
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
  }
  await act(async () => {
    root!.render(
      createElement(PdfThumb, {
        doc,
        pageNo: 1,
        rotationDelta: props.rotationDelta ?? 0,
        visible: props.visible,
        rasterW: props.rasterW ?? 120,
      }),
    )
    await Promise.resolve()
  })
  return container.querySelector('canvas')!
}

describe('PdfThumb render lifecycle', () => {
  it('cancels and releases an offscreen thumbnail, then renders it again when visible', async () => {
    const { doc, page, tasks } = makePdf()
    const canvas = await renderThumb(doc, { visible: true })
    expect(page.render).toHaveBeenCalledTimes(1)
    expect(canvas.width).toBe(120)
    expect(canvas.height).toBe(60)

    await renderThumb(doc, { visible: false })
    expect(tasks[0]!.cancel).toHaveBeenCalledOnce()
    expect(canvas.width).toBe(0)
    expect(canvas.height).toBe(0)

    await renderThumb(doc, { visible: true })
    expect(page.render).toHaveBeenCalledTimes(2)
    expect(canvas.width).toBe(120)
    tasks[1]!.resolve()
    await act(async () => Promise.resolve())
  })

  it('cancels stale renders and rerenders for rotation and raster-width changes', async () => {
    const { doc, page, tasks } = makePdf()
    await renderThumb(doc, { visible: true, rotationDelta: 0, rasterW: 100 })

    await renderThumb(doc, { visible: true, rotationDelta: 90, rasterW: 100 })
    expect(tasks[0]!.cancel).toHaveBeenCalledOnce()
    expect(page.render).toHaveBeenCalledTimes(2)

    const canvas = await renderThumb(doc, { visible: true, rotationDelta: 90, rasterW: 160 })
    expect(tasks[1]!.cancel).toHaveBeenCalledOnce()
    expect(page.render).toHaveBeenCalledTimes(3)
    expect(canvas.width).toBe(160)
    expect(page.getViewport).toHaveBeenLastCalledWith({ scale: 0.8, rotation: 90 })

    tasks[2]!.resolve()
    await act(async () => Promise.resolve())
  })

  it('waits for asynchronous cancellation settlement before reusing the canvas', async () => {
    const { doc, page, tasks } = makePdf()
    await renderThumb(doc, { visible: true, rotationDelta: 0 })
    // Model pdf.js cancel(): it requests cancellation now, but its promise does
    // not reject until asynchronous cleanup has released ownership of the canvas.
    tasks[0]!.cancel.mockImplementation(() => undefined)

    await renderThumb(doc, { visible: true, rotationDelta: 90 })
    expect(tasks[0]!.cancel).toHaveBeenCalledOnce()
    expect(page.render).toHaveBeenCalledTimes(1)

    await act(async () => {
      tasks[0]!.reject(new Error('render cancelled after cleanup'))
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(page.render).toHaveBeenCalledTimes(2)

    tasks[1]!.resolve()
    await act(async () => Promise.resolve())
  })

  it('cancels an in-flight render and releases the canvas when unmounted', async () => {
    const { doc, tasks } = makePdf()
    const canvas = await renderThumb(doc, { visible: true })
    expect(canvas.width).toBeGreaterThan(0)

    await act(async () => root!.unmount())
    root = null

    expect(tasks[0]!.cancel).toHaveBeenCalledOnce()
    expect(canvas.width).toBe(0)
    expect(canvas.height).toBe(0)
  })
})
