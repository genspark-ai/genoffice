import { describe, expect, it } from 'vitest'

import { applySheetNotes, type SheetNote } from '../src/gateway/xlsx-notes'

/**
 * A picture painted onto a note.
 *
 * Excel shows a picture in a note by painting it onto the note's shape as a
 * picture fill, so the bytes live in `xl/media` and the shape reaches them
 * through the VML part's own relationships — a different id space from the
 * worksheet's.
 */

const SHEET_PATH = 'xl/worksheets/sheet1.xml'
const VML_PATH = 'xl/drawings/vmlDrawing1.vml'
const VML_RELS_PATH = 'xl/drawings/_rels/vmlDrawing1.vml.rels'

/// A 1×1 PNG: enough bytes for the media part. The picture's size travels
/// separately, so the bytes themselves are not what is measured.
const PNG_1PX = '89504e470d0a1a0a'

function pkgOf(files: Map<string, string>, binary: Map<string, Uint8Array>) {
  return {
    paths: async () => [...files.keys(), ...binary.keys()],
    has: async (path: string) => files.has(path) || binary.has(path),
    readText: async (path: string) => {
      const content = files.get(path)
      if (content === undefined) throw new Error(`missing part: ${path}`)
      return content
    },
    write: (path: string, content: string) => void files.set(path, content),
    add: (path: string, content: string) => void files.set(path, content),
    addBinary: (path: string, bytes: Uint8Array) => void binary.set(path, bytes),
    remove: (path: string) => void files.delete(path),
  }
}

/// A workbook with one sheet, a legacy VML drawing and a note on A1.
function fixture(): Map<string, string> {
  return new Map<string, string>([
    [
      '[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
        '</Types>',
    ],
    [
      SHEET_PATH,
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheetData/><legacyDrawing r:id="rId2"/></worksheet>',
    ],
    [
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>' +
        '</Relationships>',
    ],
    [
      VML_PATH,
      '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
        '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
        '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>' +
        '</v:shapetype>' +
        '<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute">' +
        '<x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>' +
        '</xml>',
    ],
  ])
}

async function run(notes: readonly SheetNote[], files = fixture()) {
  const binary = new Map<string, Uint8Array>()
  const touched = new Set<string>()
  await applySheetNotes(pkgOf(files, binary), SHEET_PATH, notes, touched)
  return { files, binary, touched }
}

const NOTE: SheetNote = { row: 0, column: 0, author: 'Ada', text: 'the pedigree' }

const PICTURE = {
  mediaType: 'image/png' as const,
  base64: PNG_1PX,
  widthPx: 128,
  heightPx: 106,
  name: 'pedigree.png',
}

describe('putting a picture into a note', () => {
  it("paints the picture onto the note's shape", async () => {
    const { files } = await run([{ ...NOTE, picture: PICTURE }])
    const vml = files.get(VML_PATH)!
    expect(vml).toContain('type="frame"')
    expect(vml).toContain('o:title="pedigree.png"')
    expect(vml).toContain('<x:Row>0</x:Row><x:Column>0</x:Column>')
    // the yellow fill is painted behind the picture and would hide it
    const shape = vml
      .split('<v:shape ')
      .slice(1)
      .find((part) => part.includes('type="frame"'))
    expect(shape).not.toContain('fillcolor')
  })

  it("writes the media and reaches it through the VML part's own relationships", async () => {
    const { files, binary } = await run([{ ...NOTE, picture: PICTURE }])
    expect([...binary.keys()]).toEqual(['xl/media/image1.png'])
    const rels = files.get(VML_RELS_PATH)!
    expect(rels).toContain('Target="../media/image1.png"')
    expect(rels).toContain('relationships/image')
    // the relationship belongs to the VML part's id space, not the worksheet's
    const sheetRels = files.get('xl/worksheets/_rels/sheet1.xml.rels')!
    expect(sheetRels).not.toContain('media/image1.png')
  })

  it("gives the shape the picture's own size", async () => {
    // 128×106 px at 96 px per inch is 96×79.5 pt: a fill is stretched to the
    // frame it paints, so the frame has to be the picture's size
    const { files } = await run([{ ...NOTE, picture: PICTURE }])
    expect(files.get(VML_PATH)).toContain('width:96pt;height:79.5pt')
  })

  it('takes the picture away when the note says so', async () => {
    const { files } = await run([{ ...NOTE, picture: null }])
    const vml = files.get(VML_PATH)!
    expect(vml).not.toContain('type="frame"')
    expect(vml).not.toContain('<v:imagedata')
  })

  it('leaves a note that says nothing about its picture alone', async () => {
    const { files } = await run([NOTE])
    expect(files.get(VML_PATH)).not.toContain('type="frame"')
  })
})

describe('a picture a note was already carrying', () => {
  /// A workbook Excel wrote where A1's note carries a photograph: the picture
  /// lives in `xl/media` and the shape paints it as a frame fill that reaches
  /// it through the VML part's own rels.
  function workbookWithPictureNote(): Map<string, string> {
    const files = fixture()
    files.set(
      VML_PATH,
      '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
        '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
        '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>' +
        '</v:shapetype>' +
        '<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;width:96pt;height:79.5pt;visibility:hidden" o:insetmode="auto">' +
        '<v:fill o:relid="rId1" o:title="photo.png" recolor="t" rotate="t" type="frame"/>' +
        '<x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>' +
        '</xml>',
    )
    files.set(
      VML_RELS_PATH,
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>' +
        '</Relationships>',
    )
    files.set('xl/media/image1.png', PNG_1PX)
    return files
  }

  it('keeps the picture the note was carrying', async () => {
    const { files } = await run([NOTE], workbookWithPictureNote())
    expect(files.get(VML_PATH)).toContain(
      '<v:fill o:relid="rId1" o:title="photo.png" recolor="t" rotate="t" type="frame"/>',
    )
  })

  it('keeps the size the picture was shown at', async () => {
    // a fill is stretched to the frame it paints, so the frame has to stay the
    // picture's size — the text box's 108×60 would squash the photograph
    const { files } = await run([NOTE], workbookWithPictureNote())
    const withPicture = files
      .get(VML_PATH)!
      .split('<v:shape ')
      .slice(1)
      .find((part) => part.includes('type="frame"'))
    expect(withPicture).toContain('width:96pt;height:79.5pt')
    expect(withPicture).not.toContain('width:108pt')
  })

  it('leaves the VML relationships and the media part alone', async () => {
    const { files } = await run([NOTE], workbookWithPictureNote())
    // the picture is still reached through the relationship it came in with
    expect(files.get(VML_RELS_PATH)).toContain('Id="rId1"')
  })
})
