import { describe, expect, it } from 'vitest'
import { NO_SCREENSHOT_NOTE, screenshotAllowed } from '../src/renderer/ai/slide-qc'
import { redactionCount } from '../src/renderer/ai/redact-view'
import type { RenderNode } from '@genoffice/pptx-render'

/**
 * A screenshot is the one outbound path no text projection can reach.
 *
 * The renderer tints a withheld run rather than replacing it, so the words are
 * in the bitmap, and there is no mask for that — a picture has no spans to
 * cover. So a slide that withholds anything gets no picture, and the model is
 * told why rather than left to assume it is looking at the rendering.
 */

const plain: { nodes: RenderNode[] } = { nodes: [] as unknown as RenderNode[] }

/** a text node whose run carries the label, which is where the mark rides */
const withheld = {
  nodes: [
    {
      type: 'text',
      text: {
        lines: [
          {
            runs: [
              { text: 'Call ' },
              { text: '13800138000', redact: '客户电话' },
              { text: ' now' },
            ],
          },
        ],
      },
    },
  ] as unknown as RenderNode[],
}

describe('a picture of a slide', () => {
  it('goes out when the slide withholds nothing', () => {
    expect(screenshotAllowed(plain)).toBe(true)
  })

  it('does not go out when the slide withholds a span', () => {
    expect(screenshotAllowed(withheld)).toBe(false)
  })

  it('agrees with the projection about what the slide withholds', () => {
    // two sources of truth about the same slide is how a guard stops matching
    // the thing it guards
    expect(screenshotAllowed(withheld)).toBe(redactionCount(withheld) === 0)
    expect(screenshotAllowed(plain)).toBe(redactionCount(plain) === 0)
  })

  it('says why there is no picture', () => {
    // said, not omitted: a model told nothing assumes it has the rendering
    expect(NO_SCREENSHOT_NOTE).toContain('withheld')
    expect(NO_SCREENSHOT_NOTE).toContain('element inventory')
  })
})
