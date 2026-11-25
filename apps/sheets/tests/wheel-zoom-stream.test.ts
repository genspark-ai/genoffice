// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest'
import { installWheelZoomSteps } from '../src/renderer/wheel-zoom'

const cleanups: Array<() => void> = []
afterEach(() => {
  for (const dispose of cleanups.splice(0)) dispose()
  document.body.replaceChildren()
})
function harness(editing = false, initial = 1) {
  const host = document.createElement('main')
  host.innerHTML = '<div class="chart-editor">Chart</div><p>Worksheet</p>'
  document.body.append(host)
  let zoom = initial
  cleanups.push(
    installWheelZoomSteps(host, {
      getZoom: () => zoom,
      setZoom: (value) => {
        zoom = value
      },
      isCellEditing: () => editing,
    }),
  )
  return {
    zoom: () => zoom,
    wheel(
      deltaY: number,
      timeStamp: number,
      modifier: Partial<WheelEventInit> = {},
      chart = false,
    ) {
      const event = new WheelEvent('wheel', {
        deltaY,
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
        ...modifier,
      })
      Object.defineProperty(event, 'timeStamp', { value: timeStamp })
      host.querySelector(chart ? '.chart-editor' : 'p')!.dispatchEvent(event)
      return event
    },
  }
}
for (const modifier of [
  { ctrlKey: true, metaKey: false },
  { ctrlKey: false, metaKey: true },
]) {
  it(`retains fast notches through the installed Sheets listener: ${JSON.stringify(modifier)}`, () => {
    const h = harness()
    for (const t of [0, 8, 16, 24]) expect(h.wheel(-120, t, modifier).defaultPrevented).toBe(true)
    expect(h.zoom()).toBe(1.4)
  })
  it(`keeps a swipe and its repeated large tail to one step: ${JSON.stringify(modifier)}`, () => {
    const h = harness()
    for (const [i, delta] of [-10, -20, -40, -120, -120, -80].entries())
      h.wheel(delta, i * 16, modifier)
    expect(h.zoom()).toBe(1.1)
  })
}
it('keeps pinch fallthrough, editing and chart guards, line notches and limits', () => {
  const h = harness()
  expect(h.wheel(-3.4, 0).defaultPrevented).toBe(false)
  expect(h.zoom()).toBe(1)
  expect(h.wheel(-120, 8, {}, true).defaultPrevented).toBe(false)
  const edit = harness(true)
  expect(edit.wheel(-120, 0).defaultPrevented).toBe(false)
  h.wheel(-3, 300, { deltaMode: 1 })
  expect(h.zoom()).toBe(1.1)
  const cap = harness(false, 4)
  cap.wheel(-120, 0)
  expect(cap.zoom()).toBe(4)
})
