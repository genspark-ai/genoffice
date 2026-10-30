import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
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

describe('overlay picture anchor line', () => {
  const render = (attrs: Record<string, unknown>) =>
    DocProtected.config.renderHTML!.call(
      DocProtected as never,
      {
        node: {
          attrs: {
            blockType: 'image',
            docxIndex: 3,
            label: 'Image',
            imageDataUrl: 'data:image/png;base64,AA==',
            imageWidthPx: 100,
            imageHeightPx: 50,
            imageWrap: 'behind',
            anchorLine: { styleId: null, spaceBefore: 240, lineSpacing: 240, lineRule: 'auto' },
            ...attrs,
          },
        },
        HTMLAttributes: {},
      } as never,
    ) as unknown as [string, Record<string, string>, ...unknown[]]
  const blocks = (out: unknown[]) =>
    out.slice(2).filter((c) => Array.isArray(c) && c[0] === 'p') as [
      string,
      Record<string, string>,
      ...unknown[],
    ][]

  it('keeps leading spaces and tabs on the paragraph-mark line (one line box)', () => {
    const out = render({ imageLeadingText: '  \t' })
    expect(out[1].class).toContain('doc-img-float-line')
    const lines = blocks(out)
    expect(lines).toHaveLength(1)
    expect(lines[0][1].class).toContain('doc-anchor-line')
    const leading = lines[0]
      .slice(2)
      .filter(
        (c) =>
          Array.isArray(c) &&
          (c[1] as Record<string, string> | undefined)?.class === 'doc-image-leading-space',
      )
    expect(leading).toHaveLength(1)
    const brs = lines[0].slice(2).filter((c) => Array.isArray(c) && c[0] === 'br')
    expect(brs).toHaveLength(0)
    const strayLeading = out
      .slice(2)
      .filter(
        (c) =>
          Array.isArray(c) &&
          (c[1] as Record<string, string> | undefined)?.class === 'doc-image-leading-space',
      )
    expect(strayLeading).toHaveLength(0)
  })

  it('keeps the br on the anchor line when there are no leading spaces', () => {
    const line = blocks(render({}))[0]
    expect(line.slice(2).filter((c) => Array.isArray(c) && c[0] === 'br')).toHaveLength(1)
    expect(line.slice(2).some((c) => Array.isArray(c) && c[0] === 'span')).toBe(false)
  })

  it('reserves the anchor line as a flow-root so its space-before stays inside the wrapper', () => {
    const css = readFileSync(
      join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
      'utf8',
    )
    const rule = /\.doc-protected\.doc-img-float\.doc-img-float-line\s*\{([^}]*)\}/.exec(css)
    expect(rule).not.toBeNull()
    expect(rule![1]).toMatch(/height:\s*auto/)
    expect(rule![1]).toMatch(/display:\s*flow-root/)
    const out = render({})
    expect(blocks(out)[0][1].style).toMatch(/margin-top:12(\.0)?pt/)
  })
})
