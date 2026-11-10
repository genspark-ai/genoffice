/**
 * A picture-only paragraph under an auto line multiple is the picture extent
 * plus (m - 1) x the paragraph mark's single line below it (Word probe: Calibri
 * 11 mark, line=259 -> 1.2pt, line=360 -> 6.72pt, independent of the picture
 * height; Verdana 10 line=360 -> 6.1pt). A leading tab the picture cannot share
 * the line with keeps a full text line of its own.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { lineHeightFactor } from '../src/renderer/line-metrics'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

const withPPr = (pPr: string, runs = '') =>
  IMAGE_PARAGRAPH_XML.replace('<w:p>', `<w:p>${pPr}${runs}`)

describe('picture paragraph line multiple', () => {
  it('carries the direct auto multiple and the mark font metrics on the block', async () => {
    const editor = await open(
      withPPr(
        '<w:pPr><w:spacing w:line="360" w:lineRule="auto"/>' +
          '<w:rPr><w:rFonts w:ascii="Verdana"/><w:sz w:val="20"/></w:rPr></w:pPr>',
      ),
    )
    const block = editor.view.dom.querySelector<HTMLElement>('.doc-img-para')!
    expect(block.style.getPropertyValue('--doc-line-mult')).toBe('1.5')
    // a direct rule beats a style-level exact/atLeast pin
    expect(block.style.getPropertyValue('--doc-line-fixed')).toBe('0')
    expect(block.style.getPropertyValue('--doc-line-factor')).not.toBe('')
    expect(block.style.fontSize).toBe('10pt')
    editor.destroy()
  })

  it('a direct fixed rule or explicit single pins the multiple to 1', async () => {
    for (const spacing of ['w:line="300" w:lineRule="exact"', 'w:line="240" w:lineRule="auto"']) {
      const editor = await open(withPPr(`<w:pPr><w:spacing ${spacing}/></w:pPr>`))
      const block = editor.view.dom.querySelector<HTMLElement>('.doc-img-para')!
      expect(block.style.getPropertyValue('--doc-line-mult')).toBe('1')
      editor.destroy()
    }
  })

  it('the mark follows its Latin face; an East Asian face counts only when it is the sole one', async () => {
    const cases: [string, number][] = [
      ['w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="SimSun"', lineHeightFactor('Calibri')],
      ['w:eastAsia="SimSun"', lineHeightFactor('SimSun')],
    ]
    for (const [fonts, expected] of cases) {
      const editor = await open(withPPr(`<w:pPr><w:rPr><w:rFonts ${fonts}/></w:rPr></w:pPr>`))
      const block = editor.view.dom.querySelector<HTMLElement>('.doc-img-para')!
      expect(Number(block.style.getPropertyValue('--doc-line-factor'))).toBe(expected)
      editor.destroy()
    }
  })

  it('a style-level exact/atLeast rule (--doc-line-fixed) pins the multiple to 1 in CSS', () => {
    const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')
    const block = /\.doc-protected\[data-doc-protected='image'\]\.doc-img-para \{([^}]*)\}/.exec(
      css,
    )![1]
    expect(block).toMatch(
      /--doc-pic-mult: calc\(1 \+ \(var\(--doc-line-mult, 1\) - 1\) \* \(1 - var\(--doc-line-fixed, 0\)\)\)/,
    )
    expect(block).toMatch(/--doc-pic-extra: calc\(\(var\(--doc-pic-mult\) - 1\)/)
    const span = /\.doc-image-leading-space \{([^}]*)\}/.exec(css)![1]
    expect(span).toMatch(/var\(--doc-pic-mult, var\(--doc-line-mult, 1\)\)/)
  })

  it('leaves an undeclared paragraph to the document multiple', async () => {
    const editor = await open(IMAGE_PARAGRAPH_XML)
    const block = editor.view.dom.querySelector<HTMLElement>('.doc-img-para')!
    expect(block.style.getPropertyValue('--doc-line-mult')).toBe('')
    editor.destroy()
  })

  it('adds the effectExtent on top of the multiple extra', async () => {
    const editor = await open(
      IMAGE_PARAGRAPH_XML.replace(
        '<wp:extent cx="914400" cy="914400"/>',
        '<wp:extent cx="914400" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="6985"/>',
      ),
    )
    const block = editor.view.dom.querySelector<HTMLElement>('.doc-img-para')!
    expect(block.style.getPropertyValue('--doc-pic-effect-b')).toBe('0.73px')
    // no inline padding: the typed-grid whole-cell padding rule must stay on top
    expect(block.style.paddingBottom).toBe('')
    expect(block.style.paddingTop).toBe('')
    editor.destroy()
  })

  it('renders a leading tab as its own preserved span', async () => {
    for (const tab of ['<w:tab/>', '<w:tab></w:tab>']) {
      const editor = await open(withPPr('', `<w:r>${tab}</w:r>`))
      const span = editor.view.dom.querySelector<HTMLElement>('.doc-image-leading-space')!
      expect(span.textContent).toBe('\t')
      editor.destroy()
    }
  })
})
