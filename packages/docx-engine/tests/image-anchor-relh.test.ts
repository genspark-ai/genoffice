/**
 * wp:positionH relativeFrom="page"/"margin" with a posOffset: the offset
 * measures from that edge, one left margin before the column origin the
 * renderer positions from. Word keeps the page x and the full extent even
 * when the picture is wider than the column.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const PIC =
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6700000" cy="2540000"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic>'

const anchor = (positionH: string, wrap: string) =>
  '<w:r><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  positionH +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>232695</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="6700000" cy="2540000"/>' +
  wrap +
  '<wp:docPr id="1" name="Image 1"/>' +
  PIC +
  '</wp:anchor></w:drawing></w:r>'

const PAGE_H =
  '<wp:positionH relativeFrom="page"><wp:posOffset>444500</wp:posOffset></wp:positionH>'
const COLUMN_H =
  '<wp:positionH relativeFrom="column"><wp:posOffset>444500</wp:posOffset></wp:positionH>'
const PAGE_ALIGN_H = '<wp:positionH relativeFrom="page"><wp:align>left</wp:align></wp:positionH>'

const open = (bodyXml: string) => buildDocx({ bodyXml, withImage: true }).then(parseDocx)

describe('page-relative anchor horizontal offsets', () => {
  it('flags a block-level wrapTopAndBottom picture positioned from the page edge', async () => {
    const doc = await open(`<w:p>${anchor(PAGE_H, '<wp:wrapTopAndBottom/>')}</w:p>`)
    const block = doc.blocks[0]
    expect(block.type).toBe('image')
    expect(block.imageWrap).toBe('topBottom')
    expect(block.imageOffsetXEmu).toBe(444500)
    expect(block.imageRelH).toBe('page')
  })

  it('flags a run-level square-wrapped picture positioned from the page edge', async () => {
    const doc = await open(
      `<w:p>${anchor(PAGE_H, '<wp:wrapSquare wrapText="bothSides"/>')}<w:r><w:t>T</w:t></w:r></w:p>`,
    )
    const image = doc.blocks[0].runs!.find((r) => r.image)!.image!
    expect(image.offsetXEmu).toBe(444500)
    expect(image.relH).toBe('page')
  })

  it('keeps column-relative offsets and page wp:align unflagged', async () => {
    const column = await open(`<w:p>${anchor(COLUMN_H, '<wp:wrapTopAndBottom/>')}</w:p>`)
    expect(column.blocks[0].imageRelH).toBeUndefined()
    const aligned = await open(`<w:p>${anchor(PAGE_ALIGN_H, '<wp:wrapTopAndBottom/>')}</w:p>`)
    expect(aligned.blocks[0].imageOffsetXEmu).toBeUndefined()
    expect(aligned.blocks[0].imageRelH).toBeUndefined()
  })
})
