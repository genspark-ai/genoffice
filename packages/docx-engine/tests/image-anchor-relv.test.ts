/**
 * wp:positionV relativeFrom="page"/"margin" with a posOffset: the offset is a
 * position on the page, not a distance below the anchor paragraph. The parser
 * flags it (run-level and block-level pictures) so the renderer can re-pin the
 * picture on the page its anchor lands on.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const PIC =
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1806363" cy="738967"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic>'

const anchor = (positionV: string, wrap = '<wp:wrapNone/>', behind = '1') =>
  '<w:r><w:drawing>' +
  `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="${behind}" locked="0" layoutInCell="1" allowOverlap="1">` +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  positionV +
  '<wp:extent cx="1806363" cy="738967"/>' +
  wrap +
  '<wp:docPr id="1" name="Image 1"/>' +
  PIC +
  '</wp:anchor></w:drawing></w:r>'

const PAGE_V =
  '<wp:positionV relativeFrom="page"><wp:posOffset>548640</wp:posOffset></wp:positionV>'
const MARGIN_V =
  '<wp:positionV relativeFrom="margin"><wp:posOffset>548640</wp:posOffset></wp:positionV>'
const PARA_V =
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>548640</wp:posOffset></wp:positionV>'
const PAGE_ALIGN_V = '<wp:positionV relativeFrom="page"><wp:align>top</wp:align></wp:positionV>'

const open = (bodyXml: string) => buildDocx({ bodyXml, withImage: true }).then(parseDocx)

describe('page-relative anchor vertical offsets', () => {
  it('flags a run-level behind-text logo anchored from the page top', async () => {
    const doc = await open(`<w:p>${anchor(PAGE_V)}<w:r><w:t>School name</w:t></w:r></w:p>`)
    const image = doc.blocks[0].runs!.find((r) => r.image)!.image!
    expect(image.wrap).toBe('behind')
    expect(image.offsetYEmu).toBe(548640)
    expect(image.relV).toBe('page')
  })

  it('keeps paragraph-relative offsets unflagged', async () => {
    const doc = await open(`<w:p>${anchor(PARA_V)}<w:r><w:t>School name</w:t></w:r></w:p>`)
    const image = doc.blocks[0].runs!.find((r) => r.image)!.image!
    expect(image.offsetYEmu).toBe(548640)
    expect(image.relV).toBeUndefined()
  })

  it('ignores page-relative wp:align (no numeric offset to re-pin)', async () => {
    const doc = await open(`<w:p>${anchor(PAGE_ALIGN_V)}<w:r><w:t>School name</w:t></w:r></w:p>`)
    const image = doc.blocks[0].runs!.find((r) => r.image)!.image!
    expect(image.offsetYEmu).toBeUndefined()
    expect(image.relV).toBeUndefined()
  })

  it('flags a block-level square-wrapped picture anchored from the margin', async () => {
    const doc = await open(
      `<w:p>${anchor(MARGIN_V, '<wp:wrapSquare wrapText="bothSides"/>', '0')}</w:p>`,
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('image')
    expect(block.imageWrap).toBe('square-left')
    expect(block.imageOffsetYEmu).toBe(548640)
    expect(block.imageRelV).toBe('margin')
  })
})
