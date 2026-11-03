/**
 * A wpg group made only of pictures (a photo over its drop-shadow bitmap)
 * is one run picture whose layers paint as stacked backgrounds, each at its
 * fraction of the box, so the second picture shows instead of only the first.
 */
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { pictureLayerDecls } from '../src/renderer/editor/image-rotation'

const pic = (x: number, y: number, cx: number, cy: number) =>
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="P"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  `<pic:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'

const GROUP =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/><wp:docPr id="9" name="Group 9"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
  '<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:cNvGrpSpPr/>' +
  '<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="1905000" cy="952500"/></a:xfrm></wpg:grpSpPr>' +
  pic(0, 0, 1905000, 952500) +
  pic(190500, 95250, 1524000, 762000) +
  '</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

describe('pictureLayerDecls', () => {
  it('stacks the layers top first and solves percentage positions for the edges', () => {
    const decls = pictureLayerDecls([
      { dataUrl: 'data:a', x: 0, y: 0, w: 1, h: 1 },
      { dataUrl: 'data:b', x: 0.1, y: 0.1, w: 0.8, h: 0.8 },
    ])
    expect(decls).toContain('background-image:url("data:b"),url("data:a")')
    expect(decls).toContain('background-size:80.00% 80.00%,100.00% 100.00%')
    expect(decls).toContain('background-position:50.00% 50.00%,0% 0%')
    expect(pictureLayerDecls(null)).toEqual([])
  })
})

describe('grouped pictures in the editor', () => {
  it('paints both pictures of the group on one run image', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: GROUP, withImage: true }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const img = editor.view.dom.querySelector<HTMLImageElement>('img.doc-inline-img')!
    expect(img.style.width).toBe('200px')
    expect(img.style.height).toBe('100px')
    expect(img.style.backgroundImage.match(/url\(/g)).toHaveLength(2)
    expect(img.style.backgroundSize).toBe('80.00% 80.00%, 100.00% 100.00%')
    expect(img.style.backgroundPosition).toBe('50% 50%, 0% 0%')
    editor.destroy()
  })
})
