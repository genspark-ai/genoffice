/**
 * Issue #2481: the contract's blanks are legacy VML <v:line> shapes. The
 * parsed model carries them as floating rule images (text paragraphs) and as
 * an anchored decorative rule (empty anchor paragraph); this pins the DOM the
 * editor emits, i.e. a rule drawn at Word's own offsets instead of a caret or
 * a misplaced flow line.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const NS =
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office"'

interface ConnectorOptions {
  id: number
  xEmu: number
  yEmu: number
  widthEmu: number
  vmlStyle: string
}

/** one mc:AlternateContent run: wps connector Choice + VML <v:line> Fallback */
function connectorRun(opts: ConnectorOptions): string {
  const anchor =
    `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="${opts.id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${opts.xEmu}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${opts.yEmu}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${opts.widthEmu}" cy="0"/><wp:effectExtent l="0" t="0" r="19050" b="19050"/>` +
    `<wp:wrapNone/><wp:docPr id="${opts.id}" name="Connector ${opts.id}"/>` +
    `<wp:cNvGraphicFramePr/><a:graphic>` +
    `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp><wps:cNvCnPr/><wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${opts.widthEmu}" cy="0"/></a:xfrm>` +
    `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>` +
    `<a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>` +
    `</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:anchor>`
  const vmlLine =
    `<v:line id="_x0000_s1${opts.id}" o:spt="20" style="position:absolute;left:0pt;${opts.vmlStyle}" ` +
    `filled="f" stroked="t" coordsize="21600,21600"><v:fill on="f" focussize="0,0"/>` +
    `<v:stroke weight="0.5pt" color="#000000 [3213]" miterlimit="8" joinstyle="miter"/>` +
    `<v:imagedata o:title=""/><o:lock v:ext="edit" aspectratio="f"/></v:line>`
  return (
    `<w:r><mc:AlternateContent ${NS}>` +
    `<mc:Choice Requires="wps"><w:drawing>${anchor}</w:drawing></mc:Choice>` +
    `<mc:Fallback><w:pict>${vmlLine}</w:pict></mc:Fallback>` +
    `</mc:AlternateContent></w:r>`
  )
}

const FIRST_LINE = {
  id: 1,
  xEmu: 457200,
  yEmu: 381000,
  widthEmu: 2286000,
  vmlStyle: 'margin-left:36pt;margin-top:30pt;height:0pt;width:180pt;z-index:251661312;',
}
const SECOND_LINE = {
  id: 2,
  xEmu: 457200,
  yEmu: 3175,
  widthEmu: 2286000,
  vmlStyle: 'margin-left:36pt;margin-top:0.25pt;height:0pt;width:180pt;z-index:251662336;',
}
const THIRD_LINE = {
  id: 3,
  xEmu: 2162175,
  yEmu: 320675,
  widthEmu: 3124200,
  vmlStyle: 'margin-left:170.25pt;margin-top:25.25pt;height:0pt;width:246pt;z-index:251660288;',
}

const BODY_XML =
  '<w:p><w:r><w:t xml:space="preserve">甲方：           </w:t></w:r>' +
  connectorRun(FIRST_LINE) +
  '<w:r><w:t xml:space="preserve">以下简称（“甲方”）</w:t></w:r></w:p>' +
  '<w:p>' +
  connectorRun(SECOND_LINE) +
  '</w:p>' +
  '<w:p><w:r><w:t>甲乙双方共同合作的项目为：</w:t></w:r>' +
  connectorRun(THIRD_LINE) +
  '</w:p>'

let element: HTMLDivElement | null = null
let editor: Editor | null = null

afterEach(() => {
  editor?.destroy()
  editor = null
  element?.remove()
  element = null
})

async function render(): Promise<HTMLDivElement> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: BODY_XML }))
  element = document.createElement('div')
  document.body.append(element)
  editor = new Editor({
    element,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return element
}

describe('VML v:line blanks render as rules (issue #2481)', () => {
  it('draws the two text-paragraph blanks as floating rule images at their offsets', async () => {
    const host = await render()
    const images = Array.from(host.querySelectorAll<HTMLImageElement>('img[data-inline-image]'))
    expect(images).toHaveLength(2)

    const first = images[0]!
    const wrapper = first.parentElement!
    expect(wrapper.classList.contains('doc-inline-img-anchor')).toBe(true)
    expect(first.style.position).toBe('absolute')
    expect(first.style.width).toBe('240px')
    expect(first.style.height).toBe('2px')
    expect(first.getAttribute('src')).toMatch(/^data:image\/svg\+xml,/)
    const svg = decodeURIComponent(first.getAttribute('src')!.replace('data:image/svg+xml,', ''))
    expect(svg).toContain('<line')
    expect(svg).toContain('stroke="#000000"')

    // the third blank sits further down the column: 170.25pt from its start
    expect(images[1]!.style.width).toBe('328px')
    const secondWrapper = images[1]!.parentElement!
    expect(secondWrapper.classList.contains('doc-inline-img-anchor')).toBe(true)
  })

  it('draws the empty-anchor blank as an anchored rule at Word’s X/Y offsets', async () => {
    const host = await render()
    const line = host.querySelector<HTMLElement>('.doc-protected-rule-anchored > .doc-rule-line')
    expect(line, 'the anchored decorative rule').toBeTruthy()
    expect(line!.style.left).toBe('36pt')
    expect(line!.style.top).toBe('0.25pt')
    expect(line!.style.width).toBe('240px')
  })

  it('keeps the blank paragraphs editable text, not protected chips', async () => {
    const host = await render()
    expect(host.textContent).toContain('甲方：')
    expect(host.textContent).toContain('以下简称（“甲方”）')
    expect(host.textContent).toContain('甲乙双方共同合作的项目为：')
    // only the empty anchor paragraph is a protected display block
    expect(host.querySelectorAll('[data-doc-protected]')).toHaveLength(1)
  })
})
