import { readFileSync } from 'node:fs'
import { URL as NodeURL } from 'node:url'
import { runInNewContext } from 'node:vm'
import ts from 'typescript'
import { afterEach, expect, it } from 'vitest'
import { createZoomWheelClassifier } from '@genoffice/ui'
import * as docsZoom from '../src/renderer/wheel-zoom'

// Exercise App's actual effect and DOM WheelEvents without mounting its unrelated services.
const source = ts.createSourceFile(
  'App.tsx',
  readFileSync(new NodeURL('../src/renderer/App.tsx', import.meta.url), 'utf8'),
  ts.ScriptTarget.Latest,
  true,
  ts.ScriptKind.TSX,
)
let wheelEffect: ts.Expression | undefined
function visit(node: ts.Node): void {
  if (
    ts.isCallExpression(node) &&
    node.expression.getText(source) === 'useEffect' &&
    node.arguments[0]?.getText(source).includes("window.addEventListener('wheel', onWheel") &&
    node.arguments[0]?.getText(source).includes('.editor-scroll')
  )
    wheelEffect = node.arguments[0]
  ts.forEachChild(node, visit)
}
visit(source)
if (!wheelEffect) throw new Error('Docs wheel effect not found')
const effectJs = ts.transpile(`(${wheelEffect.getText(source)})`, {
  target: ts.ScriptTarget.ES2022,
})
const cleanups: Array<() => void> = []
afterEach(() => {
  for (const cleanup of cleanups.splice(0)) cleanup()
  document.body.replaceChildren()
})
function harness(initial = 100) {
  document.body.innerHTML =
    '<main class="editor-scroll"><p>Synthetic document</p></main><aside>Outside editor</aside>'
  const container = document.querySelector('main')!,
    content = container.querySelector('p')!
  Object.defineProperty(container, 'getBoundingClientRect', {
    value: () => ({ left: 20, top: 30 }),
  })
  let zoom = initial
  const anchor = { current: null as { vx: number; vy: number } | null }
  const setup = runInNewContext(effectJs, {
    ...docsZoom,
    createZoomWheelClassifier,
    window,
    zoomAnchorRef: anchor,
    scrollContainerRef: { current: container },
    setZoom: (update: (z: number) => number) => {
      zoom = update(zoom)
    },
  })
  cleanups.push(setup())
  return {
    zoom: () => zoom,
    anchor,
    wheel: (
      deltaY: number,
      timeStamp: number,
      options: Partial<WheelEventInit> = {},
      target: Element = content,
    ) => {
      const event = new WheelEvent('wheel', {
        deltaY,
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
        clientX: 150,
        clientY: 180,
        ...options,
      })
      Object.defineProperty(event, 'timeStamp', { value: timeStamp })
      target.dispatchEvent(event)
      return event
    },
  }
}
it('retains every fast Ctrl+mouse notch through the installed Docs handler', () => {
  const h = harness()
  for (const t of [0, 8, 16, 24]) expect(h.wheel(-120, t).defaultPrevented).toBe(true)
  expect(h.zoom()).toBe(140)
  expect(h.anchor.current).toEqual({ vx: 130, vy: 150 })
})
it('slow and fast notches produce the same zoom change', () => {
  const slow = harness()
  for (const t of [0, 80, 160, 240]) slow.wheel(-120, t)
  expect(slow.zoom()).toBe(140)
})
it('handles rapid zoom-out, reversal and caps without leaving a clamped anchor', () => {
  const h = harness()
  for (const t of [0, 8, 16]) h.wheel(120, t)
  h.wheel(-120, 24)
  expect(h.zoom()).toBe(80)
  const cap = harness(500)
  cap.wheel(-120, 0)
  expect(cap.zoom()).toBe(500)
  expect(cap.anchor.current).toBeNull()
})
it('keeps fractional pinch continuous and line-mode notches discrete', () => {
  const h = harness()
  h.wheel(-3.4, 0)
  h.wheel(-7.25, 8)
  expect(h.zoom()).toBeCloseTo(106.39, 8)
  h.wheel(-3, 16, { deltaMode: 1 })
  expect(h.zoom()).toBe(116)
})
it('ignores ordinary scrolling, events outside the document and zero delta', () => {
  const h = harness()
  expect(h.wheel(-120, 0, { ctrlKey: false }).defaultPrevented).toBe(false)
  expect(h.wheel(-120, 8, {}, document.querySelector('aside')!).defaultPrevented).toBe(false)
  h.wheel(0, 16)
  expect(h.zoom()).toBe(100)
  expect(h.anchor.current).toBeNull()
})
it('keeps Cmd+trackpad gesture tail suppression', () => {
  const h = harness()
  for (const [delta, t] of [
    [-10, 0],
    [-20, 16],
    [-40, 32],
    [-120, 48],
    [-80, 64],
  ])
    h.wheel(delta, t, { ctrlKey: false, metaKey: true })
  expect(h.zoom()).toBe(110)
})
