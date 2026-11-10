import { describe, expect, it } from 'vitest'
import { createRenderContext } from '../src/generate/render-context'
import { makeNodeRuns } from '../src/generate/word-utils'

const NBSP = ' '

/** the text of a docx run, read out of its w:t child */
const runText = (run: unknown): string => {
  const out: string[] = []
  const walk = (node: unknown): void => {
    if (!node || typeof node !== 'object') return
    if (Array.isArray(node)) {
      node.forEach(walk)
      return
    }
    const entry = node as { rootKey?: string; root?: unknown }
    if (entry.rootKey === 'w:t') {
      for (const child of (entry.root ?? []) as unknown[]) {
        if (typeof child === 'string') out.push(child)
      }
    }
    walk(entry.root)
  }
  walk(run)
  return out.join('')
}

/** a boxed chip whose only run carries the given font size and left inset */
const boxedNode = (sizePx: number, borderLeftSpacePx = 10) => ({
  runs: [{ text: 'chip', sizePx }],
  style: { borderLeftSpacePx },
})

describe('makeNodeRuns border-left spacer', () => {
  it('keeps the spacer short for a degenerate run font size', () => {
    const context = createRenderContext({})
    // sizePx is truthy, so the `|| 16` fallback never fires: a 1e-9px font put
    // the count at ~1.5e10 and '\u00A0'.repeat(count) threw or allocated GBs.
    const runs = makeNodeRuns(context, boxedNode(1e-9))
    const spacer = runText(runs[0])
    expect(spacer.length).toBeGreaterThan(0)
    expect(spacer.length).toBeLessThanOrEqual(1024)
  })

  it('still pads an honest font size to cover the inset', () => {
    const context = createRenderContext({})
    // 40px inset at a 16px font is ceil(40 / 10.4) = 4 nbsp
    const runs = makeNodeRuns(context, boxedNode(16, 40))
    const spacer = runText(runs[0])
    expect(spacer).toBe(NBSP.repeat(4))
  })
})
