import { describe, expect, it } from 'vitest'
import { DocProtected, anchorLineSpec } from '../src/renderer/editor/extensions'

describe('anchorLineSpec', () => {
  it('lays the line out with the paragraph format but leaves the page break on the picture', () => {
    const [tag, attrs] = anchorLineSpec({
      pageBreakBefore: true,
      keepNext: true,
      spaceBefore: 240,
      styleId: 'Caption',
    }) as [string, Record<string, string>, unknown]
    expect(tag).toBe('p')
    expect(attrs.class).toContain('doc-anchor-line')
    expect(attrs.class).not.toContain('page-break-before')
    expect(attrs['data-page-break-label']).toBeUndefined()
    expect(attrs['data-para']).toBeUndefined()
    expect(attrs.style).toContain('margin-top')
  })
})

describe('floating textbox anchor line', () => {
  const spec = (anchorLine: Record<string, unknown> | null) =>
    DocProtected.config.renderHTML!.call(
      DocProtected as never,
      {
        node: {
          attrs: {
            blockType: 'passthrough',
            docxIndex: 3,
            label: 'Text box',
            textboxes: [{ paras: [], floating: true, widthPx: 83, heightPx: 62 }],
            anchorLine,
          },
        },
        HTMLAttributes: {},
      } as never,
    ) as unknown as [string, Record<string, string>, ...unknown[]]
  const childTags = (out: unknown[]) =>
    out.slice(2).map((c) => (Array.isArray(c) ? [c[0], (c[1] as Record<string, string>).class] : c))

  it('lays the anchor paragraph out with its own spacing when the format is known', () => {
    const out = spec({ styleId: null, spaceAfter: 200, lineSpacing: 276, lineRule: 'auto' })
    expect(out[1].class).toContain('doc-protected-floating-stray')
    const line = out.slice(2).find((c) => Array.isArray(c) && c[0] === 'p') as
      [string, Record<string, string>] | undefined
    expect(line?.[1].class).toContain('doc-anchor-line')
    expect(line?.[1].class).toContain('doc-anchor-strut')
    expect(line?.[1].style).toContain('margin-bottom')
  })

  it('falls back to the bare strut without a format', () => {
    expect(childTags(spec(null))).toContainEqual(['div', 'doc-anchor-strut'])
  })
})
