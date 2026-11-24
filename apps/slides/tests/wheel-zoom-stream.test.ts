// @vitest-environment jsdom
import { readFileSync } from 'node:fs'
import { URL as NodeURL } from 'node:url'
import { runInNewContext } from 'node:vm'
import ts from 'typescript'
import { afterEach, expect, it, vi } from 'vitest'
import { createWheelPager, createZoomWheelClassifier } from '@genoffice/ui'
import { clampZoom, notchStep } from '../src/renderer/zoom-steps'

// Execute the actual wheel effect AND preview callback, including the pending
// functional updates that must compose before the delayed zoom state commit.
const source = ts.createSourceFile(
  'App.tsx',
  readFileSync(new NodeURL('../src/renderer/App.tsx', import.meta.url), 'utf8'),
  ts.ScriptTarget.Latest,
  true,
  ts.ScriptKind.TSX,
)
let effect: ts.Expression | undefined
let preview: ts.Expression | undefined
function visit(node: ts.Node): void {
  if (
    ts.isCallExpression(node) &&
    node.expression.getText(source) === 'useEffect' &&
    node.arguments[0]?.getText(source).includes("el.addEventListener('wheel', onWheel") &&
    node.arguments[0]?.getText(source).includes('zoomWheel.feed')
  )
    effect = node.arguments[0]
  if (
    ts.isVariableDeclaration(node) &&
    node.name.getText(source) === 'previewZoom' &&
    node.initializer &&
    ts.isCallExpression(node.initializer)
  )
    preview = node.initializer.arguments[0]
  ts.forEachChild(node, visit)
}
visit(source)
if (!effect || !preview) throw new Error('Slides wheel effect or preview callback not found')
const compile = (node: ts.Expression) =>
  ts.transpile(`(${node.getText(source)})`, { target: ts.ScriptTarget.ES2022 })
const effectJs = compile(effect),
  previewJs = compile(preview)
const cleanups: Array<() => void> = []
afterEach(() => {
  for (const cleanup of cleanups.splice(0)) cleanup()
  vi.useRealTimers()
  document.body.replaceChildren()
})
function harness(initial = 1) {
  vi.useFakeTimers()
  const host = document.createElement('main')
  document.body.append(host)
  let zoom = initial,
    page = 0
  const gesture = {
    current: {
      pending: null as number | null,
      anchor: null as { x: number; y: number } | null,
      raf: 0,
      timer: 0,
    },
  }
  const zoomLiveRef = { current: initial }
  const environment = {
    window,
    createWheelPager,
    createZoomWheelClassifier,
    clampZoom,
    notchStep,
    zoomGestureRef: gesture,
    zoomLiveRef,
    stageWrapRef: { current: host },
    requestAnimationFrame: () => 1,
    resetFitScroll: () => {},
    setZoom: (next: number) => {
      zoom = next
      zoomLiveRef.current = next
    },
    editingActiveRef: { current: false },
    askOpenRef: { current: false },
    rawFitRef: { current: () => 1 },
    slideLiveRef: { current: {} },
    slides: [{}, {}, {}],
    setCurrent: (update: (n: number) => number) => {
      page = update(page)
    },
  }
  const previewZoom = runInNewContext(previewJs, environment)
  const setup = runInNewContext(effectJs, { ...environment, previewZoom })
  const dispose = setup()
  cleanups.push(() => {
    dispose()
    window.clearTimeout(gesture.current.timer)
  })
  return {
    pending: () => gesture.current.pending,
    zoom: () => zoom,
    page: () => page,
    wheel(deltaY: number, timeStamp: number, modifier: Partial<WheelEventInit> = {}) {
      const event = new WheelEvent('wheel', {
        deltaY,
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
        clientX: 100,
        clientY: 120,
        ...modifier,
      })
      Object.defineProperty(event, 'timeStamp', { value: timeStamp })
      host.dispatchEvent(event)
      return event
    },
  }
}
for (const modifier of [
  { ctrlKey: true, metaKey: false },
  { ctrlKey: false, metaKey: true },
]) {
  it(`retains fast notches in the actual Slides pending preview and commit: ${JSON.stringify(modifier)}`, () => {
    const h = harness()
    for (const t of [0, 8, 16, 24]) expect(h.wheel(-120, t, modifier).defaultPrevented).toBe(true)
    expect(h.pending()).toBe(1.4)
    vi.advanceTimersByTime(150)
    expect(h.zoom()).toBe(1.4)
  })
  it(`keeps a swipe to one preview step: ${JSON.stringify(modifier)}`, () => {
    const h = harness()
    for (const [i, delta] of [-10, -20, -40, -120, -80].entries()) h.wheel(delta, i * 16, modifier)
    expect(h.pending()).toBe(1.1)
    vi.advanceTimersByTime(150)
    expect(h.zoom()).toBe(1.1)
  })
}
it('preserves continuous pinch, line mode and the slide zoom cap', () => {
  const h = harness()
  h.wheel(-3.4, 0)
  expect(h.pending()).toBeCloseTo(Math.exp(0.034), 10)
  const line = harness()
  line.wheel(-3, 0, { deltaMode: 1 })
  expect(line.pending()).toBe(1.1)
  const cap = harness(4)
  cap.wheel(-120, 0)
  expect(cap.pending()).toBe(4)
})
it('leaves ordinary slide paging on its existing gesture pager', () => {
  const h = harness()
  for (const [i, delta] of [10, 20, 40, 120, 80].entries())
    h.wheel(delta, i * 16, { ctrlKey: false })
  expect(h.pending()).toBeNull()
  expect(h.page()).toBe(1)
})
