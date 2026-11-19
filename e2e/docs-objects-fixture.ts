/** Documents with floating shapes and pictures, and how the object specs read them off the page. */
import type { ElectronApplication, Page } from '@playwright/test'
import JSZip from 'jszip'

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

/**
 * A paragraph holding one rectangle in front of the text, size in px. `x` and
 * `y` are offsets in px, or how the shape is aligned, in what `from` names:
 * the column and the paragraph, the paper, or the margins.
 */
export function shapeParagraph(
  id: number,
  x: number | 'left' | 'center' | 'right',
  y: number | 'top' | 'center' | 'bottom',
  w: number,
  h: number,
  from: 'text' | 'page' | 'margin' = 'text',
): string {
  const emu = (px: number) => px * 9525
  const [fromH, fromV] = from === 'text' ? ['column', 'paragraph'] : [from, from]
  return (
    `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
    `relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    (typeof x === 'number'
      ? `<wp:positionH relativeFrom="${fromH}"><wp:posOffset>${emu(x)}</wp:posOffset></wp:positionH>`
      : `<wp:positionH relativeFrom="${fromH}"><wp:align>${x}</wp:align></wp:positionH>`) +
    (typeof y === 'number'
      ? `<wp:positionV relativeFrom="${fromV}"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV>`
      : `<wp:positionV relativeFrom="${fromV}"><wp:align>${y}</wp:align></wp:positionV>`) +
    `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:wrapNone/><wp:docPr id="${id}" name="Rectangle ${id}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>` +
    `<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`
  )
}

/** a run holding one picture with text wrapped square around it, size in px; `x` is an offset in px or an alignment in the column */
export function pictureRun(
  id: number,
  x: number | 'left' | 'center' | 'right',
  y: number,
  w: number,
  h: number,
): string {
  const emu = (px: number) => px * 9525
  return (
    `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    (typeof x === 'number'
      ? `<wp:positionH relativeFrom="column"><wp:posOffset>${emu(x)}</wp:posOffset></wp:positionH>`
      : `<wp:positionH relativeFrom="column"><wp:align>${x}</wp:align></wp:positionH>`) +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`
  )
}

const xfrm = (x: number, y: number, w: number, h: number, attrs: string, space = '') =>
  `<a:xfrm${attrs}><a:off x="${x * 9525}" y="${y * 9525}"/><a:ext cx="${w * 9525}" cy="${h * 9525}"/>${space}</a:xfrm>`

/** a shape inside a group, px in the group's child space; `attrs` go on its a:xfrm (rot, flipH…) */
export function groupShape(
  id: number,
  box: [number, number, number, number],
  fill: string,
  opts: { prst?: string; attrs?: string; text?: string } = {},
): string {
  return (
    `<wps:wsp><wps:cNvPr id="${id}" name="Shape ${id}"/><wps:cNvSpPr/><wps:spPr>${xfrm(...box, opts.attrs ?? '')}` +
    `<a:prstGeom prst="${opts.prst ?? 'rect'}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></wps:spPr>` +
    (opts.text
      ? `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${opts.text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>`
      : '') +
    `<wps:bodyPr/></wps:wsp>`
  )
}

/** a picture inside a group, px in the group's child space */
export function groupPicture(id: number, box: [number, number, number, number]): string {
  return (
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr>${xfrm(...box, '')}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`
  )
}

/** a group inside a group: its child space is its own box */
export function innerGroup(
  id: number,
  box: [number, number, number, number],
  children: string,
  attrs = '',
): string {
  const space = `<a:chOff x="${box[0] * 9525}" y="${box[1] * 9525}"/><a:chExt cx="${box[2] * 9525}" cy="${box[3] * 9525}"/>`
  return (
    `<wpg:grpSp><wpg:cNvPr id="${id}" name="Group ${id}"/><wpg:cNvGrpSpPr/>` +
    `<wpg:grpSpPr>${xfrm(...box, attrs, space)}</wpg:grpSpPr>${children}</wpg:grpSp>`
  )
}

/**
 * A paragraph holding one group in front of the text: `box` is where it
 * stands from the column and the paragraph and its size, px; the child space
 * is the same size. `attrs` go on the group's a:xfrm (rot, flipH…).
 */
export function groupParagraph(
  id: number,
  box: [number, number, number, number],
  children: string,
  attrs = '',
  wrap = '<wp:wrapNone/>',
): string {
  const [x, y, w, h] = box.map((v) => v * 9525)
  return (
    `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
    `relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${w}" cy="${h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `${wrap}<wp:docPr id="${id}" name="Group ${id}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">` +
    `<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:cNvGrpSpPr/>` +
    `<wpg:grpSpPr>${xfrm(0, 0, box[2], box[3], attrs, `<a:chOff x="0" y="0"/><a:chExt cx="${w}" cy="${h}"/>`)}</wpg:grpSpPr>` +
    `${children}</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`
  )
}

const PNG_1PX =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

export async function docxWith(body: string): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
  )
  zip.file('word/media/image1.png', Buffer.from(PNG_1PX, 'base64'))
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${body}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

export const para = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`

export type Rect = { left: number; top: number; width: number; height: number }

export const rects = (page: Page, selector: string): Promise<Rect[]> =>
  page.evaluate(
    (selector) =>
      Array.from(document.querySelectorAll(selector)).map((el) => {
        const r = el.getBoundingClientRect()
        return { left: r.left, top: r.top, width: r.width, height: r.height }
      }),
    selector,
  )

export const boxRects = (page: Page): Promise<Rect[]> => rects(page, '.doc-page .doc-textbox')

export const saveFromMenu = (app: ElectronApplication) =>
  app.evaluate(({ Menu }) => {
    const file = Menu.getApplicationMenu()?.items.find((item) => item.label === 'File')
    file?.submenu?.items.find((candidate) => candidate.label === 'Save')?.click()
  })
