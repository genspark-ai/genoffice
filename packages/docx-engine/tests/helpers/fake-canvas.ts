/**
 * Minimal canvas stand-in for environments without a real 2D context.
 *
 * The vendored emf-converter needs a canvas to replay a metafile; in Node
 * there is none, so the tests stub OffscreenCanvas/ImageData/FileReader with
 * a recording context. The converter's own assertions live in
 * metafile-render.test.ts — this helper only exists so other suites can
 * exercise the docx parse path end to end.
 */
export interface FakeCall {
  method: string
  args: unknown[]
}

export const fakeCalls: FakeCall[] = []

function makeCtx(canvas: FakeOffscreenCanvas): Record<string, unknown> {
  const props: Record<string | symbol, unknown> = {
    canvas,
    globalCompositeOperation: 'source-over',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    font: '10px sans-serif',
    textAlign: 'left',
    textBaseline: 'alphabetic',
  }
  const fns = new Map<string, (...args: unknown[]) => unknown>()
  return new Proxy(props, {
    get(target, prop) {
      if (prop in target) return target[prop]
      const name = String(prop)
      let fn = fns.get(name)
      if (!fn) {
        fn = (...args: unknown[]) => {
          fakeCalls.push({ method: name, args })
          if (name === 'measureText') return { width: 10, actualBoundingBoxAscent: 8 }
          if (name === 'createPattern') return {}
          if (name === 'createImageData') {
            const [w, h] = args as [number, number]
            return { data: new Uint8ClampedArray(w * h * 4), width: w, height: h }
          }
          if (name === 'getImageData') {
            const [, , w, h] = args as [number, number, number, number]
            return { data: new Uint8ClampedArray(w * h * 4), width: w, height: h }
          }
          if (name === 'createLinearGradient' || name === 'createRadialGradient') {
            return { addColorStop() {} }
          }
          return undefined
        }
        fns.set(name, fn)
      }
      return fn
    },
    set(target, prop, value) {
      target[prop] = value
      return true
    },
  })
}

export class FakeOffscreenCanvas {
  width: number
  height: number
  private ctx: Record<string, unknown> | null = null
  constructor(width: number, height: number) {
    this.width = width
    this.height = height
  }
  getContext(type: string) {
    if (type !== '2d') return null
    this.ctx ??= makeCtx(this)
    return this.ctx
  }
  convertToBlob() {
    return Promise.resolve(
      new Blob([new Uint8Array([0x89, 0x50, 0x4e, 0x47])], { type: 'image/png' }),
    )
  }
}

export class FakeImageData {
  data: Uint8ClampedArray
  width: number
  height: number
  constructor(data: Uint8ClampedArray, width: number, height: number) {
    this.data = data
    this.width = width
    this.height = height
  }
}

export class FakeFileReader {
  result: string | null = null
  onload: (() => void) | null = null
  onerror: (() => void) | null = null
  error: Error | null = null
  readAsDataURL(blob: Blob) {
    void blob.arrayBuffer().then((buf) => {
      this.result = `data:${blob.type};base64,${Buffer.from(buf).toString('base64')}`
      this.onload?.()
    })
  }
}

const globals = globalThis as Record<string, unknown>
const saved: Record<string, unknown> = {}

/** install the stubs; returns a restore function */
export function installFakeCanvas(): () => void {
  for (const [key, value] of Object.entries({
    OffscreenCanvas: FakeOffscreenCanvas,
    ImageData: FakeImageData,
    FileReader: FakeFileReader,
  })) {
    saved[key] = globals[key]
    globals[key] = value
  }
  return () => {
    for (const [key, value] of Object.entries(saved)) {
      if (value === undefined) delete globals[key]
      else globals[key] = value
    }
  }
}
