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

  it('lifts stray runs by their own face and puts the space-after on the wrapper', () => {
    const out = DocProtected.config.renderHTML!.call(
      DocProtected as never,
      {
        node: {
          attrs: {
            blockType: 'passthrough',
            docxIndex: 3,
            label: 'Text box',
            textboxes: [{ paras: [], floating: true, widthPx: 83, heightPx: 62 }],
            anchorLine: {
              styleId: null,
              spaceAfter: 480,
              lineSpacing: 1.3,
              lineRawTwips: 312,
              lineRule: 'auto',
            },
            strayRuns: [{ text: 'Prepared For: ', font: 'Poppins', bold: true }, { text: 'x' }],
          },
        },
        HTMLAttributes: {},
      } as never,
    ) as unknown as [string, Record<string, string>, ...unknown[]]
    expect(out[1].style).toMatch(/margin-bottom:24(\.0)?pt/)
    const stray = out
      .slice(2)
      .find(
        (c) => Array.isArray(c) && (c[1] as Record<string, string>).class === 'doc-textbox-stray',
      ) as [string, Record<string, string>, ...unknown[]]
    expect(stray[1].style).not.toContain('margin-bottom')
    expect(stray[1].style).toContain('--doc-line-mult:1.3')
    const [poppins, plain] = stray.slice(2) as [string, Record<string, string>][]
    expect(poppins[1].class).toContain('doc-run-lf')
    expect(poppins[1].style).toMatch(/--doc-line-factor:1\.[0-9]+$/)
    expect(plain[1].class ?? '').not.toContain('doc-run-lf')
  })

  it('falls back to the bare strut without a format', () => {
    expect(childTags(spec(null))).toContainEqual(['div', 'doc-anchor-strut'])
  })
})

describe('protected block page break', () => {
  it('emits the page-break-before class on the block wrapper', () => {
    const [, attrs] = DocProtected.config.renderHTML!.call(
      DocProtected as never,
      {
        node: {
          attrs: {
            blockType: 'image',
            docxIndex: 3,
            label: 'Image',
            imageDataUrl: 'data:image/png;base64,AA==',
            pageBreakBefore: true,
          },
        },
        HTMLAttributes: {},
      } as never,
    ) as unknown as [string, Record<string, string>, ...unknown[]]
    expect(attrs.class).toContain('page-break-before')
    expect(attrs['data-page-break-label']).toBeTruthy()
  })
})
