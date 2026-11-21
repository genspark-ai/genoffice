import { describe, expect, it } from 'vitest'
import { generateTableModelXml, parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

// a table whose first cell paragraph carries a paragraph style and a bookmark
// anchor its REF fields point at (genoffice#2803, genoffice#2804)
const CELL_ANCHOR_TABLE_XML =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr>' +
  '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
  '<w:bookmarkStart w:id="7" w:name="KeepAnchor"/><w:bookmarkEnd w:id="7"/>' +
  '<w:r><w:t>Styled</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="Compact"/></w:pPr>' +
  '<w:bookmarkStart w:id="8" w:name="_Toc99001"/><w:bookmarkEnd w:id="8"/>' +
  '<w:r><w:t>Toc</w:t></w:r></w:p>' +
  '</w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Plain</w:t></w:r></w:p></w:tc>' +
  '</w:tr></w:tbl>'

describe('table cell paragraph styles and bookmark anchors', () => {
  it('parses the cell paragraph pStyle and bookmark names into the model', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: CELL_ANCHOR_TABLE_XML }))
    const rich = doc.blocks[0].table!.rows[0][0].richParas!
    expect(rich[0].styleId).toBe('Heading1')
    expect(rich[0].bookmarks).toEqual(['KeepAnchor'])
    expect(rich[1].styleId).toBe('Compact')
    expect(rich[1].hiddenBookmarks).toEqual(['_Toc99001'])
    expect(doc.blocks[0].table!.rows[0][1].richParas![0].styleId).toBeUndefined()
  })

  it('a bookmark inside a text box in the cell is not lifted onto the cell paragraph', async () => {
    const boxed =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
      '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData>' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:txbx><w:txbxContent>' +
      '<w:p><w:bookmarkStart w:id="9" w:name="Inner"/><w:bookmarkEnd w:id="9"/><w:r><w:t>box</w:t></w:r></w:p>' +
      '</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' +
      '<w:r><w:t>Host</w:t></w:r></w:p>' +
      '</w:tc></w:tr></w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: boxed }))
    expect(doc.blocks[0].table!.rows[0][0].richParas![0].bookmarks).toBeUndefined()
  })

  it('a regenerated cell re-emits the pStyle and the bookmark anchors', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: CELL_ANCHOR_TABLE_XML }))
    const xml = generateTableModelXml(doc.blocks[0].table!)
    expect(xml).toContain('<w:pStyle w:val="Heading1"/>')
    expect(xml).toContain('<w:pStyle w:val="Compact"/>')
    expect(xml).toContain('w:name="KeepAnchor"')
    expect(xml).toContain('w:name="_Toc99001"')
    expect(xml.match(/<w:bookmarkStart /g)).toHaveLength(2)
    expect(xml.match(/<w:bookmarkEnd /g)).toHaveLength(2)
  })

  it('a supplied allocator mints the rebuilt anchors ids; without one the ids stay self-unique', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: CELL_ANCHOR_TABLE_XML }))
    const allocated: string[] = []
    const xml = generateTableModelXml(doc.blocks[0].table!, undefined, {
      allocateBookmarkId: (name) => {
        allocated.push(name)
        return 500 + allocated.length
      },
    })
    expect(xml).toContain('<w:bookmarkStart w:id="501" w:name="KeepAnchor"/>')
    expect(xml).toContain('<w:bookmarkStart w:id="502" w:name="_Toc99001"/>')
    // the process-counter fallback still emits ids unique among themselves
    const plain = generateTableModelXml(doc.blocks[0].table!)
    const ids = [...plain.matchAll(/<w:bookmarkStart w:id="(\d+)"/g)].map((m) => m[1])
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('the anchors survive a save and reload (round trip)', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: CELL_ANCHOR_TABLE_XML }))
    const xml = generateTableModelXml(doc.blocks[0].table!)
    const saved = await saveDocx(doc, [{ kind: 'xml', xml }])
    const again = await parseDocx(saved)
    const rich = again.blocks[0].table!.rows[0][0].richParas!
    expect(rich[0].styleId).toBe('Heading1')
    expect(rich[0].bookmarks).toEqual(['KeepAnchor'])
    expect(rich[1].hiddenBookmarks).toEqual(['_Toc99001'])
  })
})
