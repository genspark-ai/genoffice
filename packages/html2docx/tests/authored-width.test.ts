import { afterEach, describe, expect, it } from 'vitest'
import { widestAuthoredWidth } from '../src/convert'

type Stub = Record<string, unknown>
const g = globalThis as unknown as Stub

const realDocument = g.document
const realWindow = g.window
const realGetComputedStyle = g.getComputedStyle

afterEach(() => {
  g.document = realDocument
  g.window = realWindow
  g.getComputedStyle = realGetComputedStyle
})

/** stub a page whose every element is a qualifying authored-width candidate */
function stubPage(count: number, maxWidth = 880): void {
  const els: Stub[] = []
  for (let i = 0; i < count; i++) {
    els.push({ getBoundingClientRect: () => ({ width: 900, height: 40 }) })
  }
  g.document = {
    body: {
      querySelectorAll: () => els,
      getBoundingClientRect: () => ({ width: 900, height: 40 }),
    },
  }
  g.window = { innerWidth: 794 }
  g.getComputedStyle = () => ({ maxWidth: `${maxWidth}px` })
}

describe('widestAuthoredWidth', () => {
  it('returns the widest qualifying max-width, or null when none qualify', () => {
    stubPage(3, 880)
    expect(widestAuthoredWidth()).toBe(880)
    stubPage(3, 1200)
    expect(widestAuthoredWidth()).toBe(1200)
    g.document = {
      body: {
        querySelectorAll: () => [],
        getBoundingClientRect: () => ({ width: 900, height: 40 }),
      },
    }
    g.window = { innerWidth: 794 }
    g.getComputedStyle = () => ({ maxWidth: 'none' })
    expect(widestAuthoredWidth()).toBeNull()
  })

  it('handles a document with more candidates than the argument limit', () => {
    stubPage(200_000)
    expect(widestAuthoredWidth()).toBe(880)
  })
})
