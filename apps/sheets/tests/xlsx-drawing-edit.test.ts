import { describe, expect, it } from 'vitest'

import {
  applyVisualEdits,
  VisualEditError,
} from '@genoffice/xlsx-gateway/gateway/xlsx-drawing-edit'
import type { MutablePackage } from '@genoffice/xlsx-gateway/gateway/xlsx-drawing-add'

const ANCHOR = {
  fromRow: 2,
  fromColumn: 1,
  fromRowOffset: 0,
  fromColumnOffset: 9525,
  toRow: 12,
  toColumn: 7,
  toRowOffset: -9525,
  toColumnOffset: 0,
}

const marker = (prefix: 'from' | 'to', col: number, row: number): string =>
  `<xdr:${prefix}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff>` +
  `<xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${prefix}>`

// Document order: [0] chart graphicFrame, [1] picture, [2] one-cell shape.
const DRAWING =
  '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
  `<xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}` +
  '<xdr:graphicFrame macro=""><a:graphic><a:graphicData>' +
  '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId7"/>' +
  '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>' +
  `<xdr:twoCellAnchor editAs="oneCell">${marker('from', 5, 1)}${marker('to', 9, 9)}` +
  '<xdr:pic><xdr:blipFill><a:blip r:embed="rId8"/></xdr:blipFill></xdr:pic>' +
  '<xdr:clientData/></xdr:twoCellAnchor>' +
  `<xdr:oneCellAnchor>${marker('from', 2, 20)}<xdr:ext cx="914400" cy="914400"/>` +
  '<xdr:sp><xdr:txBody/></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>' +
  '</xdr:wsDr>'

const DRAWING_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId7" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" ' +
  'Target="../charts/chart3.xml"/>' +
  '<Relationship Id="rId8" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
  'Target="../media/image1.png"/>' +
  '</Relationships>'

const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
  '<Override PartName="/xl/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
  '</Types>'

function fakePackage(entries: Map<string, string>): MutablePackage {
  return {
    paths: () => Promise.resolve([...entries.keys()]),
    has: (path) => Promise.resolve(entries.has(path)),
    readText: (path) => {
      const content = entries.get(path)
      if (content === undefined) return Promise.reject(new Error(`missing ${path}`))
      return Promise.resolve(content)
    },
    write: (path, content) => void entries.set(path, content),
    add: (path, content) => void entries.set(path, content),
    addBinary: () => undefined,
    remove: (path) => void entries.delete(path),
  }
}

const PATH = 'xl/drawings/drawing1.xml'

function packageWithDrawing(): Map<string, string> {
  return new Map([
    [PATH, DRAWING],
    ['xl/drawings/_rels/drawing1.xml.rels', DRAWING_RELS],
    ['xl/charts/chart3.xml', '<c:chartSpace/>'],
    ['xl/charts/_rels/chart3.xml.rels', '<Relationships/>'],
    ['xl/media/image1.png', 'png'],
    ['[Content_Types].xml', CONTENT_TYPES],
  ])
}

describe('applyVisualEdits', () => {
  it('removes a picture anchor and leaves the others verbatim', async () => {
    const entries = packageWithDrawing()
    const touched = new Set<string>()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 1, remove: true }],
      touched,
    )
    const xml = entries.get(PATH)!
    expect(xml).not.toContain('<xdr:pic>')
    expect(xml).toContain('<xdr:graphicFrame')
    expect(xml).toContain('<xdr:oneCellAnchor>')
    expect(touched.has(PATH)).toBe(true)
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).not.toContain('rId8')
    expect(entries.has('xl/media/image1.png')).toBe(false)
  })

  it('keeps media shared by another anchor while removing only the unused relationship', async () => {
    const entries = packageWithDrawing()
    entries.set(
      PATH,
      DRAWING.replace(
        '</xdr:wsDr>',
        `<xdr:twoCellAnchor>${marker('from', 10, 1)}${marker('to', 14, 9)}` +
          '<xdr:pic><xdr:blipFill><a:blip r:embed="rId9"/></xdr:blipFill></xdr:pic>' +
          '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>',
      ),
    )
    entries.set(
      'xl/drawings/_rels/drawing1.xml.rels',
      DRAWING_RELS.replace(
        '</Relationships>',
        '<Relationship Id="rId9" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
          'Target="../media/image1.png"/></Relationships>',
      ),
    )

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 1, remove: true }],
      new Set(),
    )

    const rels = entries.get('xl/drawings/_rels/drawing1.xml.rels')!
    expect(rels).not.toContain('rId8')
    expect(rels).toContain('rId9')
    expect(entries.has('xl/media/image1.png')).toBe(true)
  })

  it('keeps media shared by a picture on another sheet drawing', async () => {
    const entries = packageWithDrawing()
    entries.set(
      'xl/drawings/drawing2.xml',
      '<xdr:wsDr><xdr:twoCellAnchor><xdr:pic><a:blip r:embed="rId1"/></xdr:pic>' +
        '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>',
    )
    entries.set(
      'xl/drawings/_rels/drawing2.xml.rels',
      '<Relationships><Relationship Id="rId1" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
        'Target="../media/image1.png"/></Relationships>',
    )

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 1, remove: true }],
      new Set(),
    )

    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).not.toContain('rId8')
    expect(entries.has('xl/media/image1.png')).toBe(true)
  })

  it('moves a two-cell anchor by rewriting both markers', async () => {
    const entries = packageWithDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 1, anchor: ANCHOR }],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain(
      '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>9525</xdr:colOff>' +
        '<xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
    )
    expect(xml).toContain(
      '<xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff>' +
        '<xdr:row>12</xdr:row><xdr:rowOff>-9525</xdr:rowOff></xdr:to>',
    )
    // The editAs attribute and the pic content survive verbatim.
    expect(xml).toContain('editAs="oneCell"')
    expect(xml).toContain('<xdr:pic>')
  })

  it('applies an edge resize that leaves the to marker unchanged', async () => {
    const entries = packageWithDrawing()
    // Picture anchor is from(5,1)→to(9,9); an NW resize moves only `from`.
    await applyVisualEdits(
      fakePackage(entries),
      [
        {
          drawingPath: PATH,
          drawingIndex: 1,
          anchor: {
            fromRow: 3,
            fromColumn: 6,
            fromRowOffset: 0,
            fromColumnOffset: 0,
            toRow: 9,
            toColumn: 9,
            toRowOffset: 0,
            toColumnOffset: 0,
          },
        },
      ],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain(marker('from', 6, 3))
    expect(xml).toContain(marker('to', 9, 9))
  })

  it('rewrites the true frame ext when a resize sends frameSize', async () => {
    const rotated =
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
      `<xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}` +
      '<xdr:sp><xdr:spPr><a:xfrm rot="2700000"><a:off x="10" y="20"/>' +
      '<a:ext cx="100" cy="50"/></a:xfrm></xdr:spPr></xdr:sp>' +
      '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>'
    const entries = new Map([[PATH, rotated]])
    await applyVisualEdits(
      fakePackage(entries),
      [
        {
          drawingPath: PATH,
          drawingIndex: 0,
          anchor: ANCHOR,
          frameSize: { width: 555, height: 333 },
        },
      ],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain('<a:ext cx="555" cy="333"/>')
    // Rotation and offset survive verbatim; only the ext is rewritten.
    expect(xml).toContain('<a:xfrm rot="2700000"><a:off x="10" y="20"/>')
  })

  it('fails closed when frameSize targets an anchor without a frame ext', async () => {
    await expect(
      applyVisualEdits(
        fakePackage(packageWithDrawing()),
        [
          {
            drawingPath: PATH,
            drawingIndex: 1,
            anchor: ANCHOR,
            frameSize: { width: 1, height: 1 },
          },
        ],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
  })

  describe('repaint', () => {
    const shapeDrawing = (spPr: string, style = ''): string =>
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
      `<xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}` +
      `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="s"/></xdr:nvSpPr><xdr:spPr>${spPr}</xdr:spPr>${style}` +
      '<xdr:txBody><a:p><a:r><a:t>x</a:t></a:r></a:p></xdr:txBody></xdr:sp>' +
      '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>'
    const GEOM =
      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    const repaint = async (
      drawing: string,
      edit: { fillColor?: string; lineColor?: string },
    ): Promise<string> => {
      const entries = new Map([[PATH, drawing]])
      await applyVisualEdits(
        fakePackage(entries),
        [{ drawingPath: PATH, drawingIndex: 0, ...edit }],
        new Set(),
      )
      return entries.get(PATH)!
    }

    it('replaces an existing fill and leaves the outline alone', async () => {
      const xml = await repaint(
        shapeDrawing(
          `${GEOM}<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst></a:gradFill>` +
            '<a:ln w="12700"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>',
        ),
        { fillColor: '#0000ff' },
      )
      expect(xml).toContain(
        '</a:prstGeom><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>' +
          '<a:ln w="12700"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></xdr:spPr>',
      )
      expect(xml).not.toContain('gradFill')
    })

    it('writes an explicit noFill over a style fillRef', async () => {
      const style =
        '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>' +
        '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></xdr:style>'
      const xml = await repaint(shapeDrawing(GEOM, style), { fillColor: 'none' })
      expect(xml).toContain('</a:prstGeom><a:noFill/></xdr:spPr>')
      expect(xml).toContain(style)
    })

    it('swaps the outline paint but keeps its width and dash', async () => {
      const xml = await repaint(
        shapeDrawing(
          `${GEOM}<a:noFill/><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:prstDash val="dash"/></a:ln>`,
        ),
        { lineColor: '#123456' },
      )
      expect(xml).toContain(
        '<a:noFill/><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:prstDash val="dash"/></a:ln></xdr:spPr>',
      )
    })

    it('adds an outline where none existed and removes one with none', async () => {
      const added = await repaint(shapeDrawing(GEOM), { lineColor: '#123456' })
      expect(added).toContain(
        '</a:prstGeom><a:ln w="9525"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></xdr:spPr>',
      )
      const removed = await repaint(shapeDrawing(`${GEOM}<a:ln w="9525"/>`), { lineColor: 'none' })
      expect(removed).toContain('</a:prstGeom><a:ln w="9525"><a:noFill/></a:ln></xdr:spPr>')
    })

    it('re-applying the current paint is a no-op, not an error', async () => {
      const drawing = shapeDrawing(`${GEOM}<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>`)
      expect(await repaint(drawing, { fillColor: '#abcdef' })).toBe(drawing)
    })

    it('repaints and moves in one edit', async () => {
      const entries = new Map([[PATH, shapeDrawing(GEOM)]])
      await applyVisualEdits(
        fakePackage(entries),
        [{ drawingPath: PATH, drawingIndex: 0, anchor: ANCHOR, fillColor: '#ABCDEF' }],
        new Set(),
      )
      const xml = entries.get(PATH)!
      expect(xml).toContain('<a:srgbClr val="ABCDEF"/>')
      expect(xml).toContain('<xdr:row>2</xdr:row>')
    })

    it('fails closed on pictures and group shapes', async () => {
      await expect(
        applyVisualEdits(
          fakePackage(packageWithDrawing()),
          [{ drawingPath: PATH, drawingIndex: 1, fillColor: '#ABCDEF' }],
          new Set(),
        ),
      ).rejects.toThrow(VisualEditError)
      const group =
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
        `<xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}` +
        `<xdr:grpSp><xdr:grpSpPr/><xdr:sp><xdr:spPr>${GEOM}</xdr:spPr></xdr:sp></xdr:grpSp>` +
        '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>'
      await expect(repaint(group, { fillColor: '#ABCDEF' })).rejects.toThrow(VisualEditError)
    })
  })

  it('moves a one-cell anchor by rewriting only its from marker', async () => {
    const entries = packageWithDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 2, anchor: ANCHOR }],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain('<xdr:ext cx="914400" cy="914400"/>')
    expect(xml).not.toContain(`${marker('from', 2, 20)}<xdr:ext`)
  })

  it('processes several edits on one part high-index first', async () => {
    const entries = packageWithDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [
        { drawingPath: PATH, drawingIndex: 1, remove: true },
        { drawingPath: PATH, drawingIndex: 2, remove: true },
      ],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).not.toContain('<xdr:pic>')
    expect(xml).not.toContain('<xdr:oneCellAnchor>')
    expect(xml).toContain('<xdr:graphicFrame')
  })

  it('deleting a chart cascades its rel, part, own rels, and override', async () => {
    const entries = packageWithDrawing()
    const touched = new Set<string>()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      touched,
    )
    expect(entries.get(PATH)).not.toContain('<xdr:graphicFrame')
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).not.toContain('rId7')
    // The unrelated image relationship survives verbatim.
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).toContain('rId8')
    expect(entries.has('xl/charts/chart3.xml')).toBe(false)
    expect(entries.has('xl/charts/_rels/chart3.xml.rels')).toBe(false)
    expect(entries.get('[Content_Types].xml')).not.toContain('chart3.xml')
    expect(entries.get('[Content_Types].xml')).toContain('drawing1.xml')
    expect(touched.has('[Content_Types].xml')).toBe(true)
  })

  it('recursively collects chart-owned style, color, embedding, theme, and media parts', async () => {
    const entries = packageWithDrawing()
    entries.set(
      'xl/charts/_rels/chart3.xml.rels',
      '<Relationships>' +
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style3.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors3.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/source3.xlsx"/>' +
        '<Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2011/relationships/themeOverride" Target="../theme/themeOverride3.xml"/>' +
        '</Relationships>',
    )
    entries.set('xl/charts/style3.xml', '<cs:chartStyle/>')
    entries.set(
      'xl/charts/_rels/style3.xml.rels',
      '<Relationships><Relationship Id="rId1" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
        'Target="../media/chart-style.png"/></Relationships>',
    )
    entries.set('xl/charts/colors3.xml', '<cs:colorStyle/>')
    entries.set('xl/embeddings/source3.xlsx', 'embedded')
    entries.set('xl/theme/themeOverride3.xml', '<a:themeOverride/>')
    entries.set('xl/media/chart-style.png', 'png')
    entries.set(
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        '</Types>',
        '<Override PartName="/xl/charts/style3.xml" ContentType="chart-style"/>' +
          '<Override PartName="/xl/charts/colors3.xml" ContentType="chart-colors"/>' +
          '<Override PartName="/xl/theme/themeOverride3.xml" ContentType="theme-override"/>' +
          '</Types>',
      ),
    )

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )

    for (const path of [
      'xl/charts/chart3.xml',
      'xl/charts/_rels/chart3.xml.rels',
      'xl/charts/style3.xml',
      'xl/charts/_rels/style3.xml.rels',
      'xl/charts/colors3.xml',
      'xl/embeddings/source3.xlsx',
      'xl/theme/themeOverride3.xml',
      'xl/media/chart-style.png',
    ]) {
      expect(entries.has(path), path).toBe(false)
    }
    const contentTypes = entries.get('[Content_Types].xml')!
    expect(contentTypes).not.toContain('style3.xml')
    expect(contentTypes).not.toContain('colors3.xml')
    expect(contentTypes).not.toContain('themeOverride3.xml')
  })

  it('retains chart-owned dependencies referenced by another package part', async () => {
    const entries = packageWithDrawing()
    entries.set(
      'xl/charts/_rels/chart3.xml.rels',
      '<Relationships>' +
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style3.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors3.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/source3.xlsx"/>' +
        '</Relationships>',
    )
    entries.set('xl/charts/style3.xml', '<cs:chartStyle/>')
    entries.set(
      'xl/charts/_rels/style3.xml.rels',
      '<Relationships><Relationship Id="rId1" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
        'Target="../media/chart-style.png"/></Relationships>',
    )
    entries.set('xl/charts/colors3.xml', '<cs:colorStyle/>')
    entries.set('xl/embeddings/source3.xlsx', 'embedded')
    entries.set('xl/media/chart-style.png', 'png')
    entries.set('xl/charts/chart4.xml', '<c:chartSpace/>')
    entries.set(
      'xl/charts/_rels/chart4.xml.rels',
      '<Relationships>' +
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style3.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors3.xml"/>' +
        '</Relationships>',
    )

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )

    expect(entries.has('xl/charts/chart3.xml')).toBe(false)
    expect(entries.has('xl/embeddings/source3.xlsx')).toBe(false)
    expect(entries.has('xl/charts/style3.xml')).toBe(true)
    expect(entries.has('xl/charts/_rels/style3.xml.rels')).toBe(true)
    expect(entries.has('xl/charts/colors3.xml')).toBe(true)
    expect(entries.has('xl/media/chart-style.png')).toBe(true)
  })

  it('fails closed and preserves unsupported internal chart relationships', async () => {
    const entries = packageWithDrawing()
    entries.set(
      'xl/charts/_rels/chart3.xml.rels',
      '<Relationships><Relationship Id="rId1" Type="urn:vendor:chart-extension" ' +
        'Target="../custom/chart-extension.xml"/></Relationships>',
    )
    entries.set('xl/custom/chart-extension.xml', '<vendor:extension/>')
    const before = new Map(entries)

    await expect(
      applyVisualEdits(
        fakePackage(entries),
        [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)

    expect(entries).toEqual(before)
  })

  it('keeps a chart target while another anchor still references its relationship', async () => {
    const entries = packageWithDrawing()
    const duplicated = DRAWING.replace(
      '</xdr:wsDr>',
      `<xdr:twoCellAnchor>${marker('from', 0, 30)}${marker('to', 4, 38)}` +
        '<xdr:graphicFrame><a:graphic><a:graphicData>' +
        '<c:chart xmlns:c="c" xmlns:r="r" r:id="rId7"/>' +
        '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>' +
        '</xdr:wsDr>',
    )
    entries.set(PATH, duplicated)
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).toContain('rId7')
    expect(entries.has('xl/charts/chart3.xml')).toBe(true)
  })

  it('cleans the final picture drawing hookup only when it is empty and unambiguous', async () => {
    const entries = new Map([
      [
        PATH,
        '<xdr:wsDr>' +
          `<xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 2, 2)}` +
          '<xdr:pic><a:blip r:embed="rId8"/></xdr:pic>' +
          '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>',
      ],
      [
        'xl/drawings/_rels/drawing1.xml.rels',
        '<Relationships><Relationship Id="rId8" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
          'Target="../media/image1.png"/></Relationships>',
      ],
      ['xl/media/image1.png', 'png'],
      ['xl/worksheets/sheet1.xml', '<worksheet><sheetData/><drawing r:id="rId5"/></worksheet>'],
      [
        'xl/worksheets/_rels/sheet1.xml.rels',
        '<Relationships><Relationship Id="rId5" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" ' +
          'Target="../drawings/drawing1.xml"/></Relationships>',
      ],
      ['[Content_Types].xml', CONTENT_TYPES],
    ])

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )

    expect(entries.has(PATH)).toBe(false)
    expect(entries.has('xl/drawings/_rels/drawing1.xml.rels')).toBe(false)
    expect(entries.has('xl/media/image1.png')).toBe(false)
    expect(entries.get('xl/worksheets/sheet1.xml')).not.toContain('<drawing')
    expect(entries.has('xl/worksheets/_rels/sheet1.xml.rels')).toBe(false)
    expect(entries.get('[Content_Types].xml')).not.toContain('drawing1.xml')
  })

  it('preserves an empty drawing hookup that still has an unsupported relationship', async () => {
    const entries = new Map([
      [
        PATH,
        '<xdr:wsDr>' +
          `<xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 2, 2)}` +
          '<xdr:pic><a:blip r:embed="rId8"/></xdr:pic>' +
          '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>',
      ],
      [
        'xl/drawings/_rels/drawing1.xml.rels',
        '<Relationships><Relationship Id="rId8" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
          'Target="../media/image1.png"/><Relationship Id="rId99" ' +
          'Type="urn:vendor:unsupported" Target="../custom/vendor.xml"/></Relationships>',
      ],
      ['xl/media/image1.png', 'png'],
      ['xl/custom/vendor.xml', 'custom'],
      ['xl/worksheets/sheet1.xml', '<worksheet><sheetData/><drawing r:id="rId5"/></worksheet>'],
      [
        'xl/worksheets/_rels/sheet1.xml.rels',
        '<Relationships><Relationship Id="rId5" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" ' +
          'Target="../drawings/drawing1.xml"/></Relationships>',
      ],
      ['[Content_Types].xml', CONTENT_TYPES],
    ])

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )

    expect(entries.has(PATH)).toBe(true)
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).toContain('rId99')
    expect(entries.has('xl/custom/vendor.xml')).toBe(true)
    expect(entries.get('xl/worksheets/sheet1.xml')).toContain('<drawing')
  })

  it('fails closed on non-chart graphic frames, absolute anchors, and bad indexes', async () => {
    const frame = new Map(packageWithDrawing())
    frame.set(PATH, DRAWING.replace(/<c:chart\b[^>]*\/>/, '<a:tbl/>'))
    await expect(
      applyVisualEdits(
        fakePackage(frame),
        [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
    const absolute = new Map([
      [
        PATH,
        '<xdr:wsDr><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:sp/><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>',
      ],
    ])
    await expect(
      applyVisualEdits(
        fakePackage(absolute),
        [{ drawingPath: PATH, drawingIndex: 0, anchor: ANCHOR }],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
    await expect(
      applyVisualEdits(
        fakePackage(packageWithDrawing()),
        [{ drawingPath: PATH, drawingIndex: 9, remove: true }],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
    await expect(
      applyVisualEdits(
        fakePackage(new Map()),
        [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
    await expect(
      applyVisualEdits(
        fakePackage(packageWithDrawing()),
        [
          { drawingPath: PATH, drawingIndex: 1, remove: true },
          { drawingPath: PATH, drawingIndex: 1, anchor: ANCHOR },
        ],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
  })
})

// openpyxl writes the spreadsheetDrawing namespace as the DEFAULT namespace:
// anchors, markers, pic and graphicFrame all carry no prefix. The sidecar
// counts anchors by local name, so edits must find and rewrite these too.
const DEFAULT_NS_DRAWING =
  '<wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
  '<oneCellAnchor><from><col>3</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from>' +
  '<ext cx="5400000" cy="2700000"/><graphicFrame><a:graphic><a:graphicData>' +
  '<c:chart r:id="rId7"/></a:graphicData></a:graphic></graphicFrame>' +
  '<clientData/></oneCellAnchor>' +
  '<twoCellAnchor><from><col>1</col><colOff>0</colOff><row>2</row><rowOff>0</rowOff></from>' +
  '<to><col>4</col><colOff>0</colOff><row>8</row><rowOff>0</rowOff></to>' +
  '<pic><blipFill><a:blip r:embed="rId8"/></blipFill></pic>' +
  '<clientData/></twoCellAnchor>' +
  '</wsDr>'

function packageWithDefaultNsDrawing(): Map<string, string> {
  return new Map([
    [PATH, DEFAULT_NS_DRAWING],
    ['xl/drawings/_rels/drawing1.xml.rels', DRAWING_RELS],
    ['xl/charts/chart3.xml', '<c:chartSpace/>'],
    ['xl/charts/_rels/chart3.xml.rels', '<Relationships/>'],
    ['xl/media/image1.png', 'png'],
    ['[Content_Types].xml', CONTENT_TYPES],
  ])
}

describe('applyVisualEdits on default-namespace (openpyxl) drawings', () => {
  it('moves an unprefixed two-cell anchor, keeping its markers unprefixed', async () => {
    const entries = packageWithDefaultNsDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 1, anchor: ANCHOR }],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain(
      '<from><col>1</col><colOff>9525</colOff><row>2</row><rowOff>0</rowOff></from>',
    )
    expect(xml).toContain(
      '<to><col>7</col><colOff>0</colOff><row>12</row><rowOff>-9525</rowOff></to>',
    )
    expect(xml).not.toContain('xdr:')
  })

  it('moves an unprefixed one-cell anchor by its from marker only', async () => {
    const entries = packageWithDefaultNsDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, anchor: ANCHOR }],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain(
      '<from><col>1</col><colOff>9525</colOff><row>2</row><rowOff>0</rowOff></from>',
    )
    expect(xml).toContain('<ext cx="5400000" cy="2700000"/>')
  })

  it('removes an unprefixed picture anchor and cascades its image relationship', async () => {
    const entries = packageWithDefaultNsDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 1, remove: true }],
      new Set(),
    )
    expect(entries.get(PATH)!).not.toContain('<pic>')
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).not.toContain('rId8')
    expect(entries.has('xl/media/image1.png')).toBe(false)
  })

  it('removes an unprefixed chart graphic frame and cascades the chart part', async () => {
    const entries = packageWithDefaultNsDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )
    expect(entries.get(PATH)!).not.toContain('<graphicFrame>')
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).not.toContain('rId7')
    expect(entries.has('xl/charts/chart3.xml')).toBe(false)
  })

  it('keeps anchor indexes aligned across mixed prefixed and unprefixed anchors', async () => {
    const entries = packageWithDefaultNsDrawing()
    entries.set(
      PATH,
      DEFAULT_NS_DRAWING.replace(
        '<oneCellAnchor>',
        '<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
          `${marker('from', 0, 0)}${marker('to', 2, 2)}<xdr:sp/><xdr:clientData/>` +
          '</xdr:twoCellAnchor><oneCellAnchor>',
      ),
    )
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 2, remove: true }],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain('<xdr:sp/>')
    expect(xml).toContain('<graphicFrame>')
    expect(xml).not.toContain('<pic>')
  })

  it('cleans the final default-namespace drawing hookup once it is empty', async () => {
    const entries = new Map([
      [
        PATH,
        '<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          '<twoCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from>' +
          '<to><col>2</col><colOff>0</colOff><row>2</row><rowOff>0</rowOff></to>' +
          '<pic><a:blip r:embed="rId8"/></pic><clientData/></twoCellAnchor></wsDr>',
      ],
      [
        'xl/drawings/_rels/drawing1.xml.rels',
        '<Relationships><Relationship Id="rId8" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
          'Target="../media/image1.png"/></Relationships>',
      ],
      ['xl/media/image1.png', 'png'],
      ['xl/worksheets/sheet1.xml', '<worksheet><sheetData/><drawing r:id="rId5"/></worksheet>'],
      [
        'xl/worksheets/_rels/sheet1.xml.rels',
        '<Relationships><Relationship Id="rId5" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" ' +
          'Target="../drawings/drawing1.xml"/></Relationships>',
      ],
      ['[Content_Types].xml', CONTENT_TYPES],
    ])

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )

    expect(entries.has(PATH)).toBe(false)
    expect(entries.get('xl/worksheets/sheet1.xml')).not.toContain('<drawing')
  })
})

describe('applyVisualEdits arrange properties', () => {
  const picture =
    `<xdr:twoCellAnchor editAs="oneCell">${marker('from', 0, 0)}${marker('to', 4, 8)}` +
    '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1" descr="old"/></xdr:nvPicPr>' +
    '<xdr:blipFill><a:blip r:embed="rId8"/></xdr:blipFill>' +
    '<xdr:spPr><a:xfrm rot="600000" flipV="1"><a:off x="0" y="0"/><a:ext cx="100" cy="50"/></a:xfrm></xdr:spPr>' +
    '</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>'
  const shape =
    `<xdr:twoCellAnchor>${marker('from', 5, 1)}${marker('to', 9, 9)}` +
    '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="3" name="Shape 1"><a:hlinkClick r:id="rId5"/></xdr:cNvPr></xdr:nvSpPr>' +
    '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></xdr:spPr>' +
    '</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>'
  const line =
    `<xdr:oneCellAnchor>${marker('from', 2, 20)}<xdr:ext cx="914400" cy="914400"/>` +
    '<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id="4" name="Line 1"/></xdr:nvCxnSpPr>' +
    '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></xdr:spPr>' +
    '</xdr:cxnSp><xdr:clientData/></xdr:oneCellAnchor>'
  const drawing =
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `${picture}\n${shape}\n${line}</xdr:wsDr>`
  const rels =
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>' +
    '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://old.example" TargetMode="External"/>' +
    '</Relationships>'
  const RELS_PATH = 'xl/drawings/_rels/drawing1.xml.rels'
  const entriesWith = (): Map<string, string> =>
    new Map([
      [PATH, drawing],
      [RELS_PATH, rels],
      ['[Content_Types].xml', CONTENT_TYPES],
    ])
  const anchorsOf = (xml: string): string[] =>
    [...xml.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/xdr:\1>/g)].map((m) => m[0])

  it('reorders anchors by zIndex and keeps unlisted ones in place', async () => {
    const entries = entriesWith()
    await applyVisualEdits(
      fakePackage(entries),
      [
        { drawingPath: PATH, drawingIndex: 0, zIndex: 2 },
        { drawingPath: PATH, drawingIndex: 1, zIndex: 0 },
        { drawingPath: PATH, drawingIndex: 2, zIndex: 1 },
      ],
      new Set(),
    )
    const names = anchorsOf(entries.get(PATH)!).map((a) => /name="([^"]+)"/.exec(a)![1])
    expect(names).toEqual(['Shape 1', 'Line 1', 'Picture 1'])
    expect(entries.get(PATH)!.startsWith('<xdr:wsDr')).toBe(true)
    expect(entries.get(PATH)!.endsWith('</xdr:wsDr>')).toBe(true)
  })

  it('reorders around a removal in the same part and keeps inter-anchor markup', async () => {
    const entries = entriesWith()
    entries.set(PATH, drawing.replace(`${shape}\n`, `${shape}<!-- keep -->\n`))
    await applyVisualEdits(
      fakePackage(entries),
      [
        { drawingPath: PATH, drawingIndex: 1, remove: true },
        { drawingPath: PATH, drawingIndex: 0, zIndex: 1 },
        { drawingPath: PATH, drawingIndex: 2, zIndex: 0 },
      ],
      new Set(),
    )
    const xml = entries.get(PATH)!
    const names = anchorsOf(xml).map((a) => /name="([^"]+)"/.exec(a)![1])
    expect(names).toEqual(['Line 1', 'Picture 1'])
    expect(xml).toContain('<!-- keep -->')
  })

  it('only permutes anchors that carry a zIndex and never those inside mc:AlternateContent', async () => {
    const entries = entriesWith()
    const wrapped =
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      `<mc:Choice Requires="sle">${line.replace('Line 1', 'Slicer 1')}</mc:Choice>` +
      `<mc:Fallback>${line.replace('Line 1', 'Fallback 1')}</mc:Fallback></mc:AlternateContent>`
    entries.set(PATH, drawing.replace('</xdr:wsDr>', `${wrapped}</xdr:wsDr>`))
    await applyVisualEdits(
      fakePackage(entries),
      [
        { drawingPath: PATH, drawingIndex: 0, zIndex: 9 },
        { drawingPath: PATH, drawingIndex: 2, zIndex: 0 },
        { drawingPath: PATH, drawingIndex: 3, zIndex: 1 },
      ],
      new Set(),
    )
    const xml = entries.get(PATH)!
    const names = anchorsOf(xml).map((a) => /name="([^"]+)"/.exec(a)![1])
    expect(names).toEqual(['Line 1', 'Shape 1', 'Picture 1', 'Slicer 1', 'Fallback 1'])
    expect(xml).toContain('<mc:Choice Requires="sle"><xdr:oneCellAnchor>')
  })

  it('allocates no hyperlink relationship for a removed anchor', async () => {
    const entries = entriesWith()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true, hyperlink: 'https://stale.example' }],
      new Set(),
    )
    expect(entries.get(RELS_PATH)).not.toContain('stale.example')
  })

  it('writes rotation and flips on the a:xfrm, clearing zero and false', async () => {
    const entries = entriesWith()
    await applyVisualEdits(
      fakePackage(entries),
      [
        { drawingPath: PATH, drawingIndex: 0, rotation: 0, flipV: false, flipH: true },
        { drawingPath: PATH, drawingIndex: 2, rotation: -90 },
      ],
      new Set(),
    )
    const [pic, , ln] = anchorsOf(entries.get(PATH)!)
    expect(pic).toContain('<a:xfrm flipH="1">')
    expect(pic).not.toContain('rot=')
    expect(ln).toContain('<a:xfrm rot="16200000">')
    expect(ln).toContain('<a:ext cx="1" cy="1"/>')
  })

  it('writes and clears cNvPr descr and the anchor editAs', async () => {
    const entries = entriesWith()
    await applyVisualEdits(
      fakePackage(entries),
      [
        { drawingPath: PATH, drawingIndex: 0, altText: '', editAs: 'twoCell' },
        {
          drawingPath: PATH,
          drawingIndex: 1,
          altText: 'Quarterly "chart" <v2>',
          editAs: 'absolute',
        },
      ],
      new Set(),
    )
    const [pic, sp] = anchorsOf(entries.get(PATH)!)
    expect(pic).toMatch(/^<xdr:twoCellAnchor>/)
    expect(pic).toContain('<xdr:cNvPr id="2" name="Picture 1"/>')
    expect(sp).toMatch(/^<xdr:twoCellAnchor editAs="absolute">/)
    expect(sp).toContain('name="Shape 1" descr="Quarterly &quot;chart&quot; &lt;v2&gt;">')
    expect(sp).toContain('<a:hlinkClick r:id="rId5"/>')
  })

  it('rejects editAs on a one-cell anchor', async () => {
    await expect(
      applyVisualEdits(
        fakePackage(entriesWith()),
        [{ drawingPath: PATH, drawingIndex: 2, editAs: 'oneCell' }],
        new Set(),
      ),
    ).rejects.toBeInstanceOf(VisualEditError)
  })

  it('adds a hyperlink relationship and replaces or drops hlinkClick', async () => {
    const entries = entriesWith()
    const touched = new Set<string>()
    await applyVisualEdits(
      fakePackage(entries),
      [
        { drawingPath: PATH, drawingIndex: 0, hyperlink: 'https://new.example/?a=1&b=2' },
        { drawingPath: PATH, drawingIndex: 1, hyperlink: '' },
      ],
      touched,
    )
    const [pic, sp] = anchorsOf(entries.get(PATH)!)
    const relId = /<a:hlinkClick[^>]*r:id="([^"]+)"/.exec(pic!)![1]
    expect(pic).toContain('<xdr:cNvPr id="2" name="Picture 1" descr="old">')
    expect(pic).toContain('</xdr:cNvPr>')
    expect(sp).not.toContain('hlinkClick')
    expect(sp).toContain('<xdr:cNvPr id="3" name="Shape 1"/>')
    const relsXml = entries.get(RELS_PATH)!
    expect(relsXml).toContain(
      `Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://new.example/?a=1&amp;b=2" TargetMode="External"`,
    )
    expect(relsXml).not.toContain('rId5')
    expect(touched.has(RELS_PATH)).toBe(true)
  })
})
