/**
 * Typed line grid and in-line objects: Word lays a picture or chart line over
 * whole grid cells with the object centred (a 279pt chart on an 18pt grid holds
 * 288pt), and a raised or lowered run never lifts its line. The render paths
 * expose the object height as --doc-obj-h; docStyleCss turns it into padding.
 */
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  buildDocx,
  CHART_PARAGRAPH_XML,
  CHART_PART_XML,
  CHART_RELS,
  IMAGE_PARAGRAPH_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

afterEach(() => vi.restoreAllMocks())

const PIC = (xfrmAttrs: string) =>
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>' +
  `<pic:spPr><a:xfrm${xfrmAttrs}><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'

const inlineRun = (xfrmAttrs = '') =>
  '<w:r><w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/>' +
  PIC(xfrmAttrs) +
  '</wp:inline></w:drawing></w:r>'

const cellParagraph = (runs: string) =>
  '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p>' +
  runs +
  '</w:p></w:tc></w:tr></w:tbl>'

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

const CAPTION =
  '<w:r><w:t xml:space="preserve">Figure </w:t></w:r>' +
  '<w:fldSimple w:instr=" SEQ Figure \\* ARABIC "><w:r><w:t>1</w:t></w:r></w:fldSimple>'

async function openChart(caption: boolean) {
  const bodyXml = caption
    ? CHART_PARAGRAPH_XML.replace('</w:r></w:p>', `</w:r>${CAPTION}</w:p>`)
    : CHART_PARAGRAPH_XML
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml,
      extraRels: CHART_RELS,
      extraParts: [
        {
          path: 'word/charts/chart1.xml',
          xml: CHART_PART_XML,
          contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
        },
      ],
    }),
  )
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

describe('--doc-obj-h on in-line objects', () => {
  it('run-level pictures carry their height', async () => {
    const editor = await open(cellParagraph(inlineRun()))
    const img = editor.view.dom.querySelector<HTMLImageElement>('img.doc-inline-img')!
    expect(img.style.getPropertyValue('--doc-obj-h')).toBe('100.0px')
    editor.destroy()
  })

  it('quarter-turned pictures keep their inset margins instead', async () => {
    const editor = await open(cellParagraph(inlineRun(' rot="5400000"')))
    const img = editor.view.dom.querySelector<HTMLImageElement>('img.doc-inline-img')!
    expect(img.style.getPropertyValue('--doc-obj-h')).toBe('')
    editor.destroy()
  })

  it('block pictures carry their height, anchored ones do not', async () => {
    const editor = await open(IMAGE_PARAGRAPH_XML)
    const block = editor.view.dom.querySelector<HTMLElement>('.doc-protected-image')!
    expect(block.style.getPropertyValue('--doc-obj-h')).toBe('96.0px')
    editor.commands.updateAttributes('docProtected', { imageWrap: 'topBottom' })
    const anchored = editor.view.dom.querySelector<HTMLElement>('.doc-protected-image')!
    expect(anchored.style.getPropertyValue('--doc-obj-h')).toBe('')
    editor.destroy()
  })

  it('charts carry their height with or without a caption in the paragraph', async () => {
    for (const caption of [false, true]) {
      const editor = await openChart(caption)
      const block = editor.view.dom.querySelector<HTMLElement>('.doc-protected-chart')!
      expect(block.style.getPropertyValue('--doc-obj-h')).toBe('336px')
      expect(!!block.querySelector('.doc-chart-caption')).toBe(caption)
      editor.destroy()
    }
  })
})

const GRID_SECT =
  '<w:sectPr><w:pgSz w:w="11900" w:h="16840"/>' +
  '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'

function parsedWith(sectPrXml: string | null, docDefaults: object = {}): ParsedDocFull {
  return {
    styles: new Map(),
    docDefaults,
    blocks: sectPrXml ? [{ docxIndex: 0, originalXml: `<w:p>${sectPrXml}</w:p>` }] : [],
  } as unknown as ParsedDocFull
}

describe('docStyleCss grid object rules', () => {
  it('pads pictures and charts to whole cells in typed-grid docs only', () => {
    const pad =
      'calc((round(up, var(--doc-obj-h) - var(--doc-grid-pitch,0.0001px) * 0.004, var(--doc-grid-pitch,0.0001px)) - var(--doc-obj-h)) / 2)'
    const css = docStyleCss(parsedWith(GRID_SECT))
    // the chart pads its plot, not the block: caption lines below keep the grid rule
    expect(css).toContain(
      `.doc-page .doc-protected[style*="--doc-obj-h"]:not(.doc-protected-chart), ` +
        `.doc-page .doc-protected-chart[style*="--doc-obj-h"] > .doc-chart { padding-block:${pad} }`,
    )
    expect(css).toContain(
      `img.doc-inline-img[style*="--doc-obj-h"] { vertical-align:bottom; margin-block:${pad} }`,
    )
    expect(docStyleCss(parsedWith(null))).not.toContain('--doc-obj-h')
  })

  it('in-flow picture and chart blocks take the document paragraph spacing', () => {
    const sel =
      '.doc-page :is(.doc-protected[data-doc-protected="image"], .doc-protected-chart):not(.doc-img-float)'
    expect(docStyleCss(parsedWith(null))).toContain(
      `${sel} { margin-top:0.0pt;margin-bottom:0.0pt }`,
    )
    expect(docStyleCss(parsedWith(null, { spaceAfterTwips: 160 }))).toContain(
      `${sel} { margin-top:0.0pt;margin-bottom:8.0pt }`,
    )
  })

  it('raised and lowered runs never lift their line', () => {
    const css = docStyleCss(parsedWith(null))
    const raised =
      ':is(sup.doc-note-ref, span[style*="vertical-align:super"], span[style*="vertical-align: super"], span[style*="vertical-align:sub"], span[style*="vertical-align: sub"])'
    const blocks = '.doc-page :is(p, h1, h2, h3, h4, h5, h6, .doc-li, .doc-textbox-para)'
    // the DOM serialises the mark as `vertical-align: super`; the run's inner
    // spans carry their own line-height rules and must collapse too
    expect(css).toContain(
      `${blocks} ${raised}, ${blocks} ${raised} :is(span, .doc-run-lf) { line-height:0 }`,
    )
    // after the typed-grid span rules and the .doc-run-lf rules so it wins every tie
    const grid = docStyleCss(parsedWith(GRID_SECT))
    const at = grid.indexOf('sup.doc-note-ref')
    expect(grid.indexOf('span { line-height:var(--doc-line-max) }')).toBeLessThan(at)
    expect(grid.lastIndexOf('.doc-run-lf { line-height:')).toBeLessThan(at)
  })
})
