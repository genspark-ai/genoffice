import { describe, expect, it } from 'vitest'

import { applyVisualEdits, VisualEditError } from '../src/gateway/xlsx-drawing-edit'
import type { MutablePackage } from '../src/gateway/xlsx-drawing-add'

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
