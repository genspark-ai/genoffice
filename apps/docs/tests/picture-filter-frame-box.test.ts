/**
 * @vitest-environment jsdom
 */
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions, pictureFilterCss } from '../src/renderer/editor/extensions'

/** replay a CSS filter list the way the compositor does (clamped per function) */
function applyFilters(css: string, x: number): number {
  let v = x
  for (const [, fn, arg] of css.matchAll(/(\w+)\(([\d.]+)\)/g)) {
    const a = Number(arg)
    if (fn === 'contrast') v = (v - 0.5) * a + 0.5
    else if (fn === 'brightness') v = v * a
    else if (fn !== 'grayscale') throw new Error(`unexpected filter ${fn}`)
    v = Math.max(0, Math.min(1, v))
  }
  return v
}

/** Word's recolor: contrast (1 + c) about mid-grey, then the brightness offset */
const word = (x: number, bright: number, contrastPct: number): number =>
  Math.max(0, Math.min(1, (x - 0.5) * (1 + contrastPct / 100) + 0.5 + bright))

async function openDoc(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

describe('pictureFilterCss', () => {
  it('reproduces the brightness/contrast line for positive and negative offsets', () => {
    for (const [bright, contrast] of [
      [0.4, 40],
      [-0.3, 40],
      [0.2, -50],
      [-0.6, -20],
      [0, 100],
      [0.3, -100],
    ]) {
      const css = pictureFilterCss({ bright, contrast: contrast / 100 })
      for (const x of [0, 0.1, 0.25, 0.5, 0.75, 0.9, 1]) {
        expect(applyFilters(css, x), `b=${bright} c=${contrast} x=${x}`).toBeCloseTo(
          word(x, bright, contrast),
          2,
        )
      }
    }
  })

  it('maps grayscale and bi-level to a grayscale step around the threshold', () => {
    expect(pictureFilterCss({ grayscale: true })).toBe('filter:grayscale(1)')
    const css = pictureFilterCss({ biLevelThresh: 0.25 })
    expect(css.startsWith('filter:grayscale(1)')).toBe(true)
    expect(applyFilters(css, 0.2)).toBe(0)
    expect(applyFilters(css, 0.3)).toBe(1)
    expect(pictureFilterCss(null)).toBe('')
  })

  it('puts the filter on the rendered picture', async () => {
    const editor = await openDoc(
      IMAGE_PARAGRAPH_XML.replace(
        '<a:blip r:embed="rId10"/>',
        '<a:blip r:embed="rId10"><a:lum bright="40000" contrast="40000"/></a:blip>',
      ),
    )
    const img = editor.view.dom.querySelector('img.doc-protected-img') as HTMLElement
    expect(img.style.filter).toMatch(/^contrast\([\d.]+\) brightness\([\d.]+\)$/)
    editor.destroy()
  })
})

describe('w:framePr frame box rendering', () => {
  it('sizes the paragraph to the frame and floats it beside the following text', async () => {
    const editor = await openDoc(
      '<w:p><w:pPr><w:framePr w:w="1440" w:h="1440" w:hSpace="180" w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:y="1"/>' +
        '<w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>' +
        '<w:p><w:r><w:t>after</w:t></w:r></w:p>',
    )
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(p.style.width).toBe('72pt')
    expect(p.style.minHeight).toBe('72pt')
    expect(p.style.float).toBe('left')
    expect(p.style.marginRight).toBe('9pt')
    editor.destroy()
  })

  it('clips an exact-height frame like Word', async () => {
    const editor = await openDoc(
      '<w:p><w:pPr><w:framePr w:w="2000" w:h="300" w:hRule="exact" w:wrap="none"/></w:pPr>' +
        '<w:r><w:t>framed</w:t></w:r></w:p>',
    )
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(p.style.height).toBe('15pt')
    expect(p.style.overflow).toBe('hidden')
    expect(p.style.float).toBe('')
    editor.destroy()
  })
})
