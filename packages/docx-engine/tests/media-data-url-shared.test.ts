import JSZip from 'jszip'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const INLINE_PIC =
  '<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="P"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="P"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing>'

describe('media data URLs', () => {
  afterEach(() => vi.restoreAllMocks())

  it('a media part referenced by many blocks is read and encoded once', async () => {
    const bodyXml = Array.from({ length: 3 }, () => `<w:p><w:r>${INLINE_PIC}</w:r></w:p>`).join('')
    const bytes = await buildDocx({ bodyXml, withImage: true })
    const probe = await JSZip.loadAsync(bytes)
    const zipObjectProto = Object.getPrototypeOf(probe.file('word/media/image1.png')!) as {
      async: (type: string) => Promise<unknown>
    }
    const asyncSpy = vi.spyOn(zipObjectProto, 'async')
    const doc = await parseDocx(bytes)
    const urls = doc.blocks.map((b) => b.imageDataUrl).filter((u): u is string => !!u)
    expect(urls).toHaveLength(3)
    expect(urls[0]).toMatch(/^data:image\/png;base64,/)
    expect(new Set(urls).size).toBe(1)
    expect(asyncSpy.mock.calls.filter((c) => c[0] === 'base64')).toHaveLength(1)
  })
})
