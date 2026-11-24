import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  croppedPixelRect,
  planMediaReduction,
  rasterInfo,
  reducedPixelSize,
  replaceDocxParts,
} from '../src/index'
import { lazyMediaPlaceholder } from '../src/lazy-media'

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml"'
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const INCH = 914400

const pic = (rId: string, fill = '') =>
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="${rId}"/>${fill || '<a:stretch><a:fillRect/></a:stretch>'}</pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic>`
const inline = (rId: string, cx: number, cy: number, fill = '') =>
  `<w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="p"/>${pic(rId, fill)}</wp:inline></w:drawing>`
const para = (inner: string) => `<w:p><w:r>${inner}</w:r></w:p>`
const rels = (map: Record<string, string>) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  Object.entries(map)
    .map(([id, target]) => `<Relationship Id="${id}" Type="${IMAGE_REL}" Target="${target}"/>`)
    .join('') +
  `</Relationships>`

async function pkg(parts: Record<string, string | Uint8Array>): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>`,
  )
  for (const [name, data] of Object.entries(parts)) zip.file(name, data)
  return zip.generateAsync({ type: 'uint8array' })
}
const body = (inner: string) => `<w:document ${NS}><w:body>${inner}</w:body></w:document>`
const BYTES = new Uint8Array([1, 2, 3, 4])

function pngHeader(width: number, height: number, extraChunk = ''): Uint8Array {
  const out = new Uint8Array(8 + 25 + (extraChunk ? 12 : 0) + 12)
  const view = new DataView(out.buffer)
  out.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  view.setUint32(8, 13)
  out.set([0x49, 0x48, 0x44, 0x52], 12)
  view.setUint32(16, width)
  view.setUint32(20, height)
  let at = 33
  if (extraChunk) {
    out.set(
      [...extraChunk].map((c) => c.charCodeAt(0)),
      at + 4,
    )
    at += 12
  }
  out.set([0x49, 0x44, 0x41, 0x54], at + 4)
  return out
}

/** an APP1 segment that is not Exif: the XMP packet cameras and editors write after it */
const XMP_APP1 = (() => {
  const body = [...'http://ns.adobe.com/xap/1.0/\0<x:xmpmeta/>'].map((c) => c.charCodeAt(0))
  return [0xff, 0xe1, (body.length + 2) >> 8, (body.length + 2) & 255, ...body]
})()

function jpegHeader(
  width: number,
  height: number,
  channels: number,
  orientation?: number,
  afterExif: number[] = [],
) {
  const exif = orientation
    ? [
        0xff,
        0xe1,
        0x00,
        0x22,
        0x45,
        0x78,
        0x69,
        0x66,
        0x00,
        0x00,
        0x4d,
        0x4d,
        0x00,
        0x2a,
        0x00,
        0x00,
        0x00,
        0x08,
        0x00,
        0x01,
        0x01,
        0x12,
        0x00,
        0x03,
        0x00,
        0x00,
        0x00,
        0x01,
        0x00,
        orientation,
        0x00,
        0x00,
        0x00,
        0x00,
        0x00,
        0x00,
      ]
    : []
  const sof = [
    0xff,
    0xc0,
    0x00,
    0x11,
    0x08,
    height >> 8,
    height & 255,
    width >> 8,
    width & 255,
    channels,
    ...new Array(9).fill(0),
  ]
  return new Uint8Array([0xff, 0xd8, ...exif, ...afterExif, ...sof])
}

describe('planMediaReduction', () => {
  it('sizes a part by the largest drawing that shows it, across body, header and footer', async () => {
    const bytes = await pkg({
      'word/document.xml': body(
        para(inline('rId1', INCH, INCH)) + para(inline('rId2', INCH, INCH)),
      ),
      'word/_rels/document.xml.rels': rels({ rId1: 'media/a.png', rId2: 'media/b.jpeg' }),
      'word/header1.xml': `<w:hdr ${NS}>${para(inline('rId7', 3 * INCH, 2 * INCH))}</w:hdr>`,
      'word/_rels/header1.xml.rels': rels({ rId7: 'media/a.png' }),
      'word/footer1.xml': `<w:ftr ${NS}>${para(inline('rId1', INCH / 2, 4 * INCH))}</w:ftr>`,
      'word/_rels/footer1.xml.rels': rels({ rId1: 'media/a.png' }),
      'word/media/a.png': BYTES,
      'word/media/b.jpeg': BYTES,
    })
    const plan = await planMediaReduction(bytes, 150)
    expect(plan.kept).toEqual([])
    expect(plan.targets.map(({ data: _data, ...rest }) => rest)).toEqual([
      { partPath: 'word/media/a.png', mime: 'image/png', widthPx: 450, heightPx: 600 },
      { partPath: 'word/media/b.jpeg', mime: 'image/jpeg', widthPx: 150, heightPx: 150 },
    ])
    expect(plan.targets[0].data).toEqual(BYTES)
  })

  it('reads the drawing a picture sits in: a picture inside a text box, and a group', async () => {
    const textBox =
      `<w:drawing><wp:anchor><wp:extent cx="${5 * INCH}" cy="${5 * INCH}"/>` +
      `<a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>` +
      para(inline('rId1', INCH, INCH)) +
      `</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`
    const group =
      `<w:drawing><wp:anchor><wp:extent cx="${2 * INCH}" cy="${3 * INCH}"/>` +
      `<a:graphic><a:graphicData><wpg:wgp>${pic('rId2')}</wpg:wgp></a:graphicData></a:graphic>` +
      `</wp:anchor></w:drawing>`
    const bytes = await pkg({
      'word/document.xml': body(para(textBox) + para(group)),
      'word/_rels/document.xml.rels': rels({ rId1: 'media/a.png', rId2: 'media/b.png' }),
      'word/media/a.png': BYTES,
      'word/media/b.png': BYTES,
    })
    const plan = await planMediaReduction(bytes, 96)
    expect(plan.targets.map((t) => [t.partPath, t.widthPx, t.heightPx])).toEqual([
      ['word/media/a.png', 96, 96],
      ['word/media/b.png', 192, 288],
    ])
  })

  it('keeps the pixels of a cropped picture that the visible part needs', async () => {
    const bytes = await pkg({
      'word/document.xml': body(
        para(inline('rId1', INCH, INCH, '<a:srcRect l="25000" r="25000" t="50000"/><a:stretch/>')),
      ),
      'word/_rels/document.xml.rels': rels({ rId1: 'media/a.png' }),
      'word/media/a.png': BYTES,
    })
    const [target] = (await planMediaReduction(bytes, 96)).targets
    expect([target.widthPx, target.heightPx]).toEqual([192, 192])
  })

  it('leaves alone what it cannot size, and other formats', async () => {
    const vmlGroup = `<w:pict><v:group style="width:100pt;height:100pt"><v:shape style="width:1000;height:1000"><v:imagedata r:id="rId3"/></v:shape></v:group></w:pict>`
    const bytes = await pkg({
      'word/document.xml': body(
        para(inline('rId1', INCH, INCH, '<a:tile sx="50000" sy="50000"/>')) +
          para(inline('rId2', INCH, INCH)) +
          para(vmlGroup) +
          para(inline('rId4', INCH, INCH)) +
          para(inline('rId5', INCH, INCH)),
      ),
      'word/_rels/document.xml.rels': rels({
        rId1: 'media/tiled.png',
        rId2: 'media/drawn.emf',
        rId3: 'media/grouped.png',
        rId4: 'media/charted.png',
        rId5: 'media/unused-elsewhere.png',
        rId9: 'media/orphan.png',
      }),
      'word/charts/chart1.xml': `<c:chartSpace ${NS}><c:spPr><a:blipFill><a:blip r:embed="rId1"/></a:blipFill></c:spPr></c:chartSpace>`,
      'word/charts/_rels/chart1.xml.rels': rels({ rId1: '../media/charted.png' }),
      'word/media/tiled.png': BYTES,
      'word/media/drawn.emf': BYTES,
      'word/media/grouped.png': BYTES,
      'word/media/charted.png': BYTES,
      'word/media/unused-elsewhere.png': BYTES,
      'word/media/orphan.png': BYTES,
    })
    const plan = await planMediaReduction(bytes, 96)
    expect(plan.targets.map((t) => t.partPath)).toEqual(['word/media/unused-elsewhere.png'])
    expect(plan.kept).toEqual([
      { partPath: 'word/media/charted.png', reason: 'unmeasured' },
      { partPath: 'word/media/drawn.emf', reason: 'format' },
      { partPath: 'word/media/grouped.png', reason: 'unmeasured' },
      { partPath: 'word/media/tiled.png', reason: 'unmeasured' },
    ])
  })

  it('sizes a VML picture by its shape and ignores the fallback of a sized choice', async () => {
    const vml = `<w:pict><v:shape style="position:absolute;width:144pt;height:1in"><v:imagedata r:id="rId1"/></v:shape></w:pict>`
    const alternate =
      `<mc:AlternateContent><mc:Choice Requires="wps">${inline('rId2', INCH, INCH)}</mc:Choice>` +
      `<mc:Fallback><w:pict><v:shape style="width:10"><v:imagedata r:id="rId2"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent>`
    const fallbackOnly = `<mc:AlternateContent><mc:Fallback><w:pict><v:shape style="width:72pt;height:72pt"><v:imagedata r:id="rId3"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent>`
    const bytes = await pkg({
      'word/document.xml': body(para(vml) + para(alternate) + para(fallbackOnly)),
      'word/_rels/document.xml.rels': rels({
        rId1: 'media/a.png',
        rId2: 'media/b.png',
        rId3: 'media/c.png',
      }),
      'word/media/a.png': BYTES,
      'word/media/b.png': BYTES,
      'word/media/c.png': BYTES,
    })
    const plan = await planMediaReduction(bytes, 96)
    expect(plan.targets.map((t) => [t.partPath, t.widthPx, t.heightPx])).toEqual([
      ['word/media/a.png', 192, 96],
      ['word/media/b.png', 96, 96],
    ])
    expect(plan.kept).toEqual([{ partPath: 'word/media/c.png', reason: 'unmeasured' }])
  })

  it('names where a lazily served part is read from instead of handing out its placeholder', async () => {
    const hash = 'ab'.repeat(32)
    const bytes = await pkg({
      'word/document.xml': body(para(inline('rId1', INCH, INCH))),
      'word/_rels/document.xml.rels': rels({ rId1: 'media/a.png' }),
      'word/media/a.png': lazyMediaPlaceholder(hash),
    })
    const [target] = (await planMediaReduction(bytes, 96)).targets
    expect(target.data).toBeUndefined()
    expect(target.lazyUrl).toBe(`genoffice-docx-media://${hash}/word/media/a.png`)
  })
})

describe('reducedPixelSize', () => {
  it('scales down to cover the need in both directions and never up', () => {
    expect(
      reducedPixelSize({ width: 2000, height: 1000 }, { widthPx: 500, heightPx: 250 }),
    ).toEqual({ width: 500, height: 250 })
    // a picture shown squashed keeps the pixels its wider direction needs
    expect(
      reducedPixelSize({ width: 2000, height: 1000 }, { widthPx: 500, heightPx: 500 }),
    ).toEqual({ width: 1000, height: 500 })
    expect(
      reducedPixelSize({ width: 400, height: 200 }, { widthPx: 500, heightPx: 250 }),
    ).toBeNull()
    expect(
      reducedPixelSize({ width: 500, height: 250 }, { widthPx: 500, heightPx: 250 }),
    ).toBeNull()
    expect(
      reducedPixelSize({ width: 2000, height: 100 }, { widthPx: 500, heightPx: 250 }),
    ).toBeNull()
  })
})

describe('rasterInfo', () => {
  it('reads the size from the bytes and refuses what a canvas would change', () => {
    expect(rasterInfo(pngHeader(640, 480))).toEqual({
      mime: 'image/png',
      width: 640,
      height: 480,
      resamplable: true,
      output: 'image/png',
    })
    expect(rasterInfo(pngHeader(640, 480, 'acTL'))?.resamplable).toBe(false)
    expect(rasterInfo(jpegHeader(1024, 768, 3))).toEqual({
      mime: 'image/jpeg',
      width: 1024,
      height: 768,
      resamplable: true,
      output: 'image/jpeg',
    })
    expect(rasterInfo(jpegHeader(1024, 768, 3, 6))?.resamplable).toBe(false)
    expect(rasterInfo(jpegHeader(1024, 768, 4))?.resamplable).toBe(false)
    // an XMP segment after the Exif one says nothing about the orientation
    expect(rasterInfo(jpegHeader(1024, 768, 3, 6, XMP_APP1))?.resamplable).toBe(false)
    expect(rasterInfo(jpegHeader(1024, 768, 3, undefined, XMP_APP1))?.resamplable).toBe(true)
    expect(rasterInfo(new Uint8Array([0x47, 0x49, 0x46, 0x38]))).toBeNull()
    expect(rasterInfo(new Uint8Array([0x89, 0x50, 0x4e, 0x47]))).toBeNull()
  })
})

describe('replaceDocxParts', () => {
  it('swaps the named parts and keeps every other part byte for byte', async () => {
    const bytes = await pkg({
      'word/document.xml': body(para(inline('rId1', INCH, INCH))),
      'word/_rels/document.xml.rels': rels({ rId1: 'media/a.png' }),
      'word/media/a.png': BYTES,
      'word/media/b.png': BYTES,
    })
    const out = await replaceDocxParts(bytes, { 'word/media/a.png': new Uint8Array([9]) })
    const before = await JSZip.loadAsync(bytes)
    const after = await JSZip.loadAsync(out)
    expect(Object.keys(after.files)).toEqual(
      Object.keys(before.files).filter((n) => !n.endsWith('/')),
    )
    expect(await after.file('word/media/a.png')!.async('uint8array')).toEqual(new Uint8Array([9]))
    for (const name of ['word/document.xml', 'word/media/b.png', '[Content_Types].xml'])
      expect(await after.file(name)!.async('uint8array')).toEqual(
        await before.file(name)!.async('uint8array'),
      )
    await expect(replaceDocxParts(bytes, { 'word/media/x.png': BYTES })).rejects.toThrow()
  })
})

describe('deleting the cropped areas of pictures', () => {
  const CROP = '<a:srcRect l="25000" r="25000" t="50000"/><a:stretch><a:fillRect/></a:stretch>'
  const cropped = (rId: string, cx = INCH) => para(inline(rId, cx, INCH, CROP))
  const build = (bodyXml: string, extra: Record<string, string | Uint8Array> = {}) =>
    pkg({
      'word/document.xml': body(bodyXml),
      'word/_rels/document.xml.rels': rels({ rId1: 'media/a.png', rId2: 'media/b.png' }),
      'word/media/a.png': BYTES,
      'word/media/b.png': BYTES,
      ...extra,
    })

  it('names the crop of a picture only when asked, and only when every use shares it', async () => {
    const bytes = await build(
      cropped('rId1') + cropped('rId1', 2 * INCH) + para(inline('rId2', INCH, INCH)),
    )
    expect((await planMediaReduction(bytes, 96)).targets[0].crop).toBeUndefined()
    const plan = await planMediaReduction(bytes, 96, { deleteCropped: true })
    expect(plan.targets.map((t) => t.crop)).toEqual([{ l: 0.25, t: 0.5, r: 0.25, b: 0 }, undefined])

    const mixed = await build(cropped('rId1') + para(inline('rId1', INCH, INCH)))
    expect((await planMediaReduction(mixed, 96, { deleteCropped: true })).targets[0].crop).toBe(
      undefined,
    )
    const outset = await build(
      para(inline('rId1', INCH, INCH, '<a:srcRect l="-10000" r="20000"/><a:stretch/>')),
    )
    expect((await planMediaReduction(outset, 96, { deleteCropped: true })).targets[0].crop).toBe(
      undefined,
    )
    const fallback = await build(
      cropped('rId1') +
        para(
          '<mc:AlternateContent><mc:Choice Requires="wps"><w:t>x</w:t></mc:Choice><mc:Fallback>' +
            '<w:pict><v:shape style="width:72pt;height:72pt"><v:imagedata r:id="rId1"/></v:shape></w:pict>' +
            '</mc:Fallback></mc:AlternateContent>',
        ),
    )
    expect((await planMediaReduction(fallback, 96, { deleteCropped: true })).targets[0].crop).toBe(
      undefined,
    )
  })

  it('keeps the stored resolution when no quality is asked for', async () => {
    const [target] = (await planMediaReduction(await build(cropped('rId1')), null)).targets
    expect([target.widthPx, target.heightPx]).toEqual([Infinity, Infinity])
    expect(reducedPixelSize({ width: 4000, height: 3000 }, target)).toBeNull()
  })

  it('finds the visible pixels and the size they are needed at', () => {
    const crop = { l: 0.25, t: 0.5, r: 0.25, b: 0 }
    expect(
      croppedPixelRect(
        { width: 800, height: 600 },
        { widthPx: Infinity, heightPx: Infinity, crop },
      ),
    ).toEqual({ sx: 200, sy: 300, sw: 400, sh: 300, width: 400, height: 300 })
    // the whole picture is needed at 192 x 192, so the visible half at 96 x 96
    expect(
      croppedPixelRect({ width: 800, height: 600 }, { widthPx: 192, heightPx: 192, crop }),
    ).toEqual({ sx: 200, sy: 300, sw: 400, sh: 300, width: 128, height: 96 })
  })

  it('takes the crop off the pictures of a part that was cut down, and off no other', async () => {
    const bytes = await build(cropped('rId1') + cropped('rId2'), {
      'word/header1.xml': body(cropped('rId7')),
      'word/_rels/header1.xml.rels': rels({ rId7: 'media/a.png' }),
    })
    const out = await JSZip.loadAsync(
      await replaceDocxParts(bytes, { 'word/media/a.png': new Uint8Array([9]) }, [
        'word/media/a.png',
      ]),
    )
    const doc = await out.file('word/document.xml')!.async('string')
    expect(doc.match(/<a:srcRect/g)).toHaveLength(1)
    expect(doc).toMatch(/r:embed="rId1"\/><a:stretch>/)
    expect(doc).toMatch(/r:embed="rId2"\/><a:srcRect l="25000" r="25000" t="50000"\/>/)
    expect(await out.file('word/header1.xml')!.async('string')).not.toContain('a:srcRect')
  })
})

describe('pictures stored as another format than PNG or JPEG', () => {
  const bytesOf = (text: string) => Uint8Array.from([...text].map((c) => c.charCodeAt(0)))
  const le16 = (v: number) => [v & 255, v >> 8]
  const gif = (frames: number) =>
    Uint8Array.from([
      ...bytesOf('GIF89a'),
      ...le16(40),
      ...le16(30),
      0x80,
      0,
      0,
      ...[0, 0, 0, 255, 255, 255],
      // a graphic control extension, then each frame: descriptor, code size, one data block
      ...Array.from({ length: frames }, () => [
        0x21,
        0xf9,
        4,
        0,
        0,
        0,
        0,
        0,
        0x2c,
        ...le16(0),
        ...le16(0),
        ...le16(40),
        ...le16(30),
        0,
        2,
        2,
        0x4c,
        0x01,
        0,
      ]).flat(),
      0x3b,
    ])
  const bmp = (w: number, h: number) => {
    const out = new Uint8Array(54)
    out.set(bytesOf('BM'))
    const view = new DataView(out.buffer)
    view.setUint32(14, 40, true)
    view.setInt32(18, w, true)
    view.setInt32(22, h, true)
    return out
  }
  const riff = (chunks: number[]) =>
    Uint8Array.from([...bytesOf('RIFF'), 0, 0, 0, 0, ...bytesOf('WEBP'), ...chunks])
  const vp8 = [...bytesOf('VP8 '), 10, 0, 0, 0, 0, 0, 0, 0x9d, 0x01, 0x2a, ...le16(64), ...le16(48)]
  // 14 bits of width - 1, then 14 of height - 1
  const vp8l = [...bytesOf('VP8L'), 5, 0, 0, 0, 0x2f, 63, 0xc0 | (47 << 6), 47 >> 2, 0]
  const vp8x = (flags: number, then: number[]) => [
    ...bytesOf('VP8X'),
    10,
    0,
    0,
    0,
    flags,
    0,
    0,
    0,
    ...[63, 0, 0],
    ...[47, 0, 0],
    ...then,
  ]

  it('reads what a one-frame GIF, a BMP, a WebP and a TIFF are, and what they are stored as', () => {
    expect(rasterInfo(gif(1))).toEqual({
      mime: 'image/gif',
      width: 40,
      height: 30,
      resamplable: true,
      output: 'image/png',
    })
    expect(rasterInfo(gif(2))?.resamplable).toBe(false)
    expect(rasterInfo(bmp(20, -10))).toMatchObject({ mime: 'image/bmp', width: 20, height: 10 })
    expect(rasterInfo(riff(vp8))).toMatchObject({ width: 64, height: 48, output: 'image/jpeg' })
    expect(rasterInfo(riff(vp8l))).toMatchObject({ width: 64, height: 48, output: 'image/png' })
    expect(rasterInfo(riff(vp8x(0, vp8)))).toMatchObject({
      width: 64,
      height: 48,
      resamplable: true,
      output: 'image/jpeg',
    })
    // alpha keeps a lossy picture a PNG; an animation and an Exif turn are left alone
    expect(rasterInfo(riff(vp8x(0x10, vp8)))?.output).toBe('image/png')
    expect(rasterInfo(riff(vp8x(0x02, vp8)))?.resamplable).toBe(false)
    expect(rasterInfo(riff(vp8x(0x08, vp8)))?.resamplable).toBe(false)
    expect(rasterInfo(bytesOf('II*\0rest'))).toMatchObject({
      mime: 'image/tiff',
      output: 'image/png',
    })
  })

  it('plans them by their extension and leaves metafiles and SVG out', async () => {
    const names = ['a.gif', 'b.bmp', 'c.webp', 'd.tiff', 'e.emf', 'f.svg', 'g.wmf']
    const bytes = await pkg({
      'word/document.xml': body(names.map((_, i) => para(inline(`rId${i}`, INCH, INCH))).join('')),
      'word/_rels/document.xml.rels': rels(
        Object.fromEntries(names.map((name, i) => [`rId${i}`, `media/${name}`])),
      ),
      ...Object.fromEntries(names.map((name) => [`word/media/${name}`, BYTES])),
    })
    const plan = await planMediaReduction(bytes, 96)
    expect(plan.targets.map((t) => [t.partPath, t.mime])).toEqual([
      ['word/media/a.gif', 'image/gif'],
      ['word/media/b.bmp', 'image/bmp'],
      ['word/media/c.webp', 'image/webp'],
      ['word/media/d.tiff', 'image/tiff'],
    ])
    expect(plan.kept.map((k) => k.partPath)).toEqual([
      'word/media/e.emf',
      'word/media/f.svg',
      'word/media/g.wmf',
    ])
  })

  it('moves a converted part to the name of its new format, with its relationships and content type', async () => {
    const bytes = await pkg({
      'word/document.xml': body(
        para(inline('rId1', INCH, INCH)) + para(inline('rId2', INCH, INCH)),
      ),
      'word/_rels/document.xml.rels': rels({ rId1: 'media/a.bmp', rId2: 'media/b.webp' }),
      'word/header1.xml': `<w:hdr ${NS}>${para(inline('rId7', INCH, INCH))}</w:hdr>`,
      'word/_rels/header1.xml.rels': rels({ rId7: '/word/media/a.bmp' }),
      'word/media/a.bmp': BYTES,
      'word/media/a.png': BYTES,
      'word/media/b.webp': BYTES,
    })
    const out = await JSZip.loadAsync(
      await replaceDocxParts(
        bytes,
        { 'word/media/a.bmp': new Uint8Array([9]), 'word/media/b.webp': new Uint8Array([8]) },
        [],
        { 'word/media/a.bmp': 'image/png', 'word/media/b.webp': 'image/jpeg' },
      ),
    )
    expect(out.file('word/media/a.bmp')).toBeNull()
    expect(out.file('word/media/b.webp')).toBeNull()
    // a.png was taken by another picture
    expect(await out.file('word/media/a-2.png')!.async('uint8array')).toEqual(new Uint8Array([9]))
    expect(await out.file('word/media/a.png')!.async('uint8array')).toEqual(BYTES)
    expect(await out.file('word/media/b.jpeg')!.async('uint8array')).toEqual(new Uint8Array([8]))
    const docRels = await out.file('word/_rels/document.xml.rels')!.async('string')
    expect(docRels).toContain('Target="media/a-2.png"')
    expect(docRels).toContain('Target="media/b.jpeg"')
    expect(await out.file('word/_rels/header1.xml.rels')!.async('string')).toContain(
      'Target="/word/media/a-2.png"',
    )
    const types = await out.file('[Content_Types].xml')!.async('string')
    expect(types.match(/Extension="png"/g)).toHaveLength(1)
    expect(types).toContain('<Default Extension="jpeg" ContentType="image/jpeg"/>')
    expect(await out.file('word/document.xml')!.async('uint8array')).toEqual(
      await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('uint8array'),
    )
  })
})
