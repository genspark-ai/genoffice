import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

describe('w:br outside a w:r (Word honors these; tdf108714 family)', () => {
  it('body-level <w:br w:type="page"/> becomes a page-break block', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>Paragraph 1</w:t></w:r></w:p>' +
          '<w:br w:type="page"/>' +
          '<w:p><w:r><w:t>Paragraph 2</w:t></w:r></w:p>',
      }),
    )
    const br = doc.blocks[1]
    expect(br.type).toBe('passthrough')
    expect(br.fieldDisplay?.kind).toBe('pageBreak')
    expect(br.invisibleMarker).toBeUndefined()
  })

  it('body-level <w:br w:type="textWrapping"/> stays invisible (no placeholder chip)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p><w:br w:type="textWrapping"/>',
      }),
    )
    expect(doc.blocks[1].invisibleMarker).toBe(true)
  })

  it('<w:br> as a direct child of w:p becomes a break run', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:br w:type="page"/><w:r><w:t>after</w:t></w:r></w:p>',
      }),
    )
    const text = (doc.blocks[0].runs ?? []).map((r) => r.text).join('')
    expect(text).toBe('\fafter')
  })

  it('page-type w:br next to an anchored VML textbox marks the block (tdf141220)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:pict>' +
          '<v:rect style="position:absolute;width:100pt;height:50pt">' +
          '<v:textbox><w:txbxContent><w:p><w:r><w:t>box text</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
          '</v:rect><w10:wrap type="none"/></w:pict></w:r>' +
          '<w:r><w:br w:type="page"/></w:r></w:p>' +
          '<w:p><w:r><w:t>next page</w:t></w:r></w:p>',
      }),
    )
    const box = doc.blocks[0]
    expect(box.label).toBe('Text box')
    expect(box.fieldDisplay?.kind).toBe('pageBreak')
  })
})
