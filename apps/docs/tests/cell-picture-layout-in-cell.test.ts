import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

const ruleBody = (selector: string): string => {
  const esc = selector.replace(/[.[\]'()]/g, '\\$&')
  const m = css.match(new RegExp('(?:^|\\n)' + esc + ' \\{([^}]*)\\}'))
  expect(m, 'rule for ' + selector).toBeTruthy()
  return m![1]
}

/** a wp:anchor wrapNone picture, the shape real templates put in table cells */
function anchoredPicture(cx: number, cy: number, x = 0, y = 0, layoutInCell = '1'): string {
  return (
    '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ' +
    'relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="' +
    layoutInCell +
    '" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="column"><wp:posOffset>' +
    x +
    '</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>' +
    y +
    '</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="' +
    cx +
    '" cy="' +
    cy +
    '"/><wp:wrapNone/><wp:docPr id="1" name="Picture 1"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' +
    cx +
    '" cy="' +
    cy +
    '"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
  )
}

const cell = (w: string, body: string): string =>
  '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/></w:tcPr>' + body + '</w:tc>'

/** three-column guide table: the middle 图示 cell holds the picture (1.39" wide,
 *  the picture is 3.5" x 2.6": wider and taller than its cell) */
function tableXml(picture: string): string {
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="8400" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="800"/><w:gridCol w:w="2000"/><w:gridCol w:w="5600"/></w:tblGrid>' +
    '<w:tr>' +
    cell('800', '<w:p><w:r><w:t>1</w:t></w:r></w:p>') +
    cell('2000', '<w:p>' + picture + '</w:p><w:p><w:r><w:t>caption</w:t></w:r></w:p>') +
    cell('5600', '<w:p><w:r><w:t>content description text next to the picture</w:t></w:r></w:p>') +
    '</w:tr>' +
    '</w:tbl>'
  )
}

async function openDoc(layoutInCell = '1') {
  const source = await buildDocx({
    bodyXml:
      tableXml(anchoredPicture(3200000, 2400000, 0, 0, layoutInCell)) +
      '<w:p><w:r><w:t>after</w:t></w:r></w:p>',
    withImage: true,
  })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const td = editor.view.dom.querySelectorAll('td')[1]!
  return { editor, parsed, source, td }
}

interface CellModel {
  richParas?: Array<{ runs?: Array<{ image?: { inCell?: boolean; wrap?: string } }> }>
}

describe('layoutInCell="1" pictures anchored in table cells (#2480)', () => {
  it('marks the anchored picture as cell-contained in the model', async () => {
    const { editor, parsed, td } = await openDoc()
    const block = parsed.blocks[0] as unknown as { table: { rows: CellModel[][] } }
    const runs = block.table.rows[0]![1]!.richParas?.[0]?.runs ?? []
    const image = runs.find((r) => r.image)?.image
    expect(image?.inCell).toBe(true)
    expect(image?.wrap).toBe('front')
    // the editor node carries the flag so its DOM renderer can confine the box
    expect(td.querySelector('span.doc-inline-img-cell')).not.toBeNull()
    editor.destroy()
  })

  it('renders the picture in a cell-width clipping window that reserves the picture bottom', async () => {
    const { editor, td } = await openDoc()
    const wrapper = td.querySelector<HTMLElement>('span.doc-inline-img-cell')!
    // 251.97px picture extent, top offset 0 -> the wrapper reserves that height
    expect(parseFloat(wrapper.style.height)).toBeCloseTo(251.97, 0)
    expect(wrapper.getAttribute('data-inline-image-anchor')).toBe('1')
    const img = wrapper.querySelector('img')!
    expect(img.style.position).toBe('absolute')
    expect(parseFloat(img.style.left)).toBeCloseTo(0, 1)
    expect(parseFloat(img.style.top)).toBeCloseTo(0, 1)
    expect(img.style.maxWidth).toBe('none')
    // the picture is not a page-positioned overlay any more: no transform /
    // page-relative pin and no anchor outside the cell
    expect(img.style.transform).toBe('')
    expect(img.hasAttribute('data-page-rel-v')).toBe(false)
    expect(wrapper.closest('td')).toBe(td)
    editor.destroy()
  })

  it('grows the reserved height with a positive anchor offset and clamps a negative one', async () => {
    const source = await buildDocx({
      bodyXml: tableXml(anchoredPicture(3200000, 2400000, -100000, 476250)) + '<w:p/>',
      withImage: true,
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const td = editor.view.dom.querySelectorAll('td')[1]!
    const wrapper = td.querySelector<HTMLElement>('span.doc-inline-img-cell')!
    const img = wrapper.querySelector('img')!
    // negative X clamps at the cell start (Word confines the picture to the cell)
    expect(parseFloat(img.style.left)).toBeCloseTo(0, 1)
    // top offset 50px + picture 251.97px: the row must make room for both
    expect(parseFloat(img.style.top)).toBeCloseTo(50, 1)
    expect(parseFloat(wrapper.style.height)).toBeCloseTo(301.97, 0)
    editor.destroy()
  })

  it('keeps layoutInCell="0" pictures on the page-positioned overlay path', async () => {
    const { editor, td } = await openDoc('0')
    expect(td.querySelector('.doc-inline-img-cell')).toBeNull()
    const img = td.querySelector<HTMLElement>('img.doc-inline-img')!
    expect(img.style.position).toBe('absolute')
    expect(img.style.transform).toContain('translate')
    editor.destroy()
  })

  it('clips the wrapper at the cell edge so the picture cannot overprint the next column', () => {
    // jsdom has no layout, so the clipping contract is asserted on the CSS rule:
    // a cell-wide inline-block window with hidden overflow, grown by the inline height
    const body = ruleBody('.doc-inline-img-cell')
    expect(body).toMatch(/\n\s*display: inline-block;/)
    expect(body).toMatch(/\n\s*position: relative;/)
    expect(body).toMatch(/\n\s*width: 100%;/)
    expect(body).toMatch(/\n\s*max-width: 100%;/)
    expect(body).toMatch(/\n\s*overflow: hidden;/)
    expect(body).toMatch(/\n\s*vertical-align: top;/)
  })

  it('saves the untouched table byte-identically', async () => {
    const { editor, parsed, source } = await openDoc()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })
})
