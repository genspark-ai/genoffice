import { test, expect } from '@playwright/test'
import { createServer, type ViteDevServer } from 'vite'
import react from '@vitejs/plugin-react'
import { resolve } from 'node:path'
import { PNG } from 'pngjs'
import type { RenderNode, RenderSlide } from '../packages/pptx-render/src/render-tree'

let server: ViteDevServer
let url: string

test.beforeAll(async () => {
  server = await createServer({
    configFile: false,
    root: resolve('apps/slides/src/renderer'),
    plugins: [react()],
    server: { host: '127.0.0.1', port: 0, fs: { allow: [process.cwd()] } },
  })
  server.middlewares.use('/clipboard-test', async (_req, res) => {
    res.setHeader('Content-Type', 'text/html')
    res.end(await server.transformIndexHtml('/clipboard-test', '<html><body></body></html>'))
  })
  await server.listen()
  url = server.resolvedUrls!.local[0]!
})

test.afterAll(async () => {
  await server?.close()
})

const shape = (
  id: string,
  x: number,
  y: number,
  w: number,
  h: number,
  color: string,
): RenderNode => ({
  id,
  sourceId: id,
  type: 'shape',
  box: { x, y, w, h, rotationDeg: 0 },
  presetGeometry: 'rect',
  fill: { kind: 'solid', color },
})
const slide = (nodes: RenderNode[]): RenderSlide => ({
  widthPx: 1280,
  heightPx: 720,
  scale: 1,
  background: { kind: 'solid', color: '#ff00ff' },
  nodes,
})
const pixel = (png: PNG, x: number, y: number) => [
  ...png.data.subarray((y * png.width + x) * 4, (y * png.width + x) * 4 + 4),
]

test('selection PNG keeps z-order, excludes unselected shapes and page background', async ({
  page,
}, info) => {
  await page.goto(`${url}clipboard-test`)
  const pngData = await page.evaluate(
    async (fixture) => {
      const moduleUrl = '/selection-image.tsx'
      const { renderSelectionToPngBase64 } = await import(moduleUrl)
      return renderSelectionToPngBase64(fixture, ['blue', 'red'], new Map())
    },
    slide([
      shape('red', 100, 100, 100, 80, '#ff0000'),
      shape('unselected', 90, 90, 200, 200, '#00ff00'),
      shape('blue', 150, 140, 80, 60, '#0000ff'),
    ]),
  )
  const bytes = Buffer.from(pngData, 'base64')
  const png = PNG.sync.read(bytes)
  expect(png.width).toBe(264) // 130px union + 1px padding on each side, at 2x
  expect(png.height).toBe(204)
  expect(pixel(png, 22, 22)).toEqual([255, 0, 0, 255])
  expect(pixel(png, 122, 102)).toEqual([0, 0, 255, 255])
  expect(pixel(png, 242, 22)[3]).toBe(0)
  await info.attach('selection.png', { body: bytes, contentType: 'image/png' })
  expect(await page.locator('canvas').count()).toBe(0)
})

test('rotated groups outside the page are not clipped and include their shadows', async ({
  page,
}) => {
  await page.goto(`${url}clipboard-test`)
  const child = shape('child', 0, 0, 100, 40, '#ff0000')
  if (child.type === 'shape')
    child.shadow = { color: '#000000', blurPx: 4, offsetX: 12, offsetY: 6 }
  const pngData = await page.evaluate(
    async (fixture) => {
      const moduleUrl = '/selection-image.tsx'
      const { renderSelectionToPngBase64 } = await import(moduleUrl)
      return renderSelectionToPngBase64(fixture, ['group'], new Map())
    },
    slide([
      {
        id: 'group',
        sourceId: 'group',
        type: 'group',
        box: { x: -100, y: -100, w: 100, h: 40, rotationDeg: 90 },
        children: [child],
      },
    ]),
  )
  const png = PNG.sync.read(Buffer.from(pngData, 'base64'))
  expect(png.height).toBeGreaterThan(200)
  expect(png.width).toBeLessThan(png.height)
  let red = 0,
    shadow = 0
  for (let i = 0; i < png.data.length; i += 4) {
    if (png.data[i] === 255 && png.data[i + 3] === 255) red++
    if (png.data[i] === 0 && png.data[i + 3]! > 0) shadow++
  }
  expect(red).toBeGreaterThan(14000)
  expect(shadow).toBeGreaterThan(100)
})

test('pictures not yet present in the editor image cache are decoded before capture', async ({
  page,
}) => {
  await page.goto(`${url}clipboard-test`)
  const pngData = await page.evaluate(async (fixture) => {
    const canvas = document.createElement('canvas')
    canvas.width = 10
    canvas.height = 10
    const c = canvas.getContext('2d')!
    c.fillStyle = '#00ff00'
    c.fillRect(0, 0, 10, 10)
    fixture.nodes = [
      {
        id: 'picture',
        sourceId: 'picture',
        type: 'picture',
        box: { x: 30, y: 50, w: 40, h: 40, rotationDeg: 0 },
        dataUrl: canvas.toDataURL(),
      },
    ]
    const moduleUrl = '/selection-image.tsx'
    const { renderSelectionToPngBase64 } = await import(moduleUrl)
    return renderSelectionToPngBase64(fixture, ['picture'], new Map())
  }, slide([]))
  const png = PNG.sync.read(Buffer.from(pngData, 'base64'))
  expect(pixel(png, 40, 40)).toEqual([0, 255, 0, 255])
})

for (const type of ['table', 'chart'] as const) {
  test(`${type} picture fills are loaded even with an empty image cache`, async ({ page }) => {
    await page.goto(`${url}clipboard-test`)
    const pngData = await page.evaluate(
      async ({ fixture, type }) => {
        const canvas = document.createElement('canvas')
        canvas.width = 10
        canvas.height = 10
        const c = canvas.getContext('2d')!
        c.fillStyle = '#00ff00'
        c.fillRect(0, 0, 10, 10)
        const fill = {
          kind: 'image' as const,
          mode: 'stretch' as const,
          dataUrl: canvas.toDataURL(),
        }
        const common = {
          id: 'item',
          sourceId: 'item',
          box: { x: 30, y: 50, w: 40, h: 40, rotationDeg: 0 },
        }
        fixture.nodes = [
          type === 'table'
            ? {
                ...common,
                type: 'table',
                bgFill: fill,
                gridX: [0, 40],
                gridY: [0, 40],
                cells: [{ row: 0, col: 0, x: 0, y: 0, w: 40, h: 40, fill: { kind: 'none' } }],
              }
            : {
                ...common,
                type: 'chart',
                plotRect: { x: 0, y: 0, w: 40, h: 40, fill },
                gridLines: [],
                axisLines: [],
                labels: [],
                bars: [],
                polylines: [],
                markers: [],
                swatches: [],
              },
        ]
        const moduleUrl = '/selection-image.tsx'
        const { renderSelectionToPngBase64 } = await import(moduleUrl)
        return renderSelectionToPngBase64(fixture, ['item'], new Map())
      },
      { fixture: slide([]), type },
    )
    const png = PNG.sync.read(Buffer.from(pngData, 'base64'))
    expect(pixel(png, 40, 40)).toEqual([0, 255, 0, 255])
  })
}
