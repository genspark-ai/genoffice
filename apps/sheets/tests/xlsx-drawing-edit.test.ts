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
  '<xdr:pic><xdr:blipFill/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>' +
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

  it('fails closed when another anchor still references the deleted chart', async () => {
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
    await expect(
      applyVisualEdits(
        fakePackage(entries),
        [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
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
