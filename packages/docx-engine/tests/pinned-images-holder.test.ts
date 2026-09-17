import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, saveDocx, type NewImage } from '../src/index'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'

const pinned = (x: number, y: number, zOrder: number): NewImage => ({
  base64: TINY_PNG_BASE64,
  mime: 'image/png',
  widthPx: 16,
  heightPx: 16,
  wrap: 'behind',
  posOffsetEmu: { x, y, relativeTo: 'page' },
  zOrder,
  paraSpacing: { afterTwips: 0, lineTwips: 20, lineRule: 'exact' },
})

describe('images save block', () => {
  it('writes several anchored pictures into one holder paragraph', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>' }),
    )
    const saved = await saveDocx(parsed, [
      { kind: 'images', images: [pinned(100, 200, 1), pinned(300, 400, 2), pinned(500, 600, 3)] },
      { kind: 'original', docxIndex: 0 },
    ])
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    const paras = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []
    expect(paras).toHaveLength(2)
    expect(paras[0]!.match(/<wp:anchor/g)).toHaveLength(3)
    expect(paras[0]!.match(/<w:pPr>/g)).toHaveLength(1)
    expect(paras[0]).toContain('<wp:posOffset>500</wp:posOffset>')
    expect(paras[0]).toContain('w:line="20" w:lineRule="exact"')
    // the pictures still round-trip as three anchored drawings
    const back = await parseDocx(saved)
    const visible = back.blocks.filter((b) => !b.hidden)
    expect(visible).toHaveLength(2)
    expect((visible[0].originalXml ?? '').match(/<wp:anchor/g)).toHaveLength(3)
  })

  it('degrades to the plain image paragraph for a single picture', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>' }),
    )
    const one = await saveDocx(parsed, [{ kind: 'images', images: [pinned(1, 2, 1)] }])
    const single = await saveDocx(parsed, [{ kind: 'image', image: pinned(1, 2, 1) }])
    const xmlOf = async (b: Uint8Array) =>
      (await (await JSZip.loadAsync(b)).file('word/document.xml')!.async('string')).replace(
        /\s+/g,
        ' ',
      )
    expect(await xmlOf(one)).toBe(await xmlOf(single))
  })
})
