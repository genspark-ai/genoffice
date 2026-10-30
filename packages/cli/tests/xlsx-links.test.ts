import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'

const sidecar = Boolean(xlsxSidecarPath())

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

const STYLES =
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

/// A tiny source workbook: Sheet1 holds A1=10, A2=20, A3=30.
async function writeSource(path: string): Promise<void> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    XML_DECL +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    XML_DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    XML_DECL +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    XML_DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
  )
  zip.file('xl/styles.xml', XML_DECL + STYLES)
  zip.file(
    'xl/worksheets/sheet1.xml',
    XML_DECL +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A3"/><sheetData><row r="1"><c r="A1"><v>10</v></c></row><row r="2"><c r="A2"><v>20</v></c></row><row r="3"><c r="A3"><v>30</v></c></row></sheetData></worksheet>',
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

/// The dependent workbook: link [1] points at a file that does not exist,
/// link [2] at `source.xlsx` (a relative path). Formulas reference both.
async function writeDependent(path: string): Promise<void> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    XML_DECL +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/><Override PartName="/xl/externalLinks/externalLink2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    XML_DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    XML_DECL +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><externalReferences><externalReference r:id="rId4"/><externalReference r:id="rId5"/></externalReferences></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    XML_DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink2.xml"/></Relationships>',
  )
  zip.file('xl/styles.xml', XML_DECL + STYLES)
  for (const n of [1, 2]) {
    zip.file(
      `xl/externalLinks/externalLink${n}.xml`,
      XML_DECL +
        `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>1</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>`,
    )
    zip.file(
      `xl/externalLinks/_rels/externalLink${n}.xml.rels`,
      XML_DECL +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${n === 1 ? 'gone.xlsx' : 'source.xlsx'}" TargetMode="External"/></Relationships>`,
    )
  }
  zip.file(
    'xl/worksheets/sheet1.xml',
    XML_DECL +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1"><f>[2]Sheet1!A1*2</f><v>84</v></c><c r="B1"><f>MATCH(20,[2]Sheet1!A1:A3,0)</f><v>2</v></c></row><row r="2"><c r="A2" t="str"><f>[1]Sheet1!A1&amp;"!"</f><v>gone!</v></c></row></sheetData></worksheet>',
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

describe('genoffice sheet links', () => {
  it.skipIf(!sidecar)('lists the links with resolved sources and reachability', async () => {
    const dir = tempDir()
    await writeSource(join(dir, 'source.xlsx'))
    const book = join(dir, 'book.xlsx')
    await writeDependent(book)
    const r = await run(['sheet', 'links', book, '--json'])
    expect(r.code).toBe(0)
    const links = r.json().detail.links
    expect(links).toHaveLength(2)
    expect(links[0]).toMatchObject({ index: 1, reachable: false, rawTarget: 'gone.xlsx' })
    expect(links[1]).toMatchObject({
      index: 2,
      reachable: true,
      sheetNames: ['Sheet1'],
    })
    expect(links[1].source.endsWith('source.xlsx')).toBe(true)
  })

  it.skipIf(!sidecar)(
    'update writes live cached values and keeps every formula spelling',
    async () => {
      const dir = tempDir()
      await writeSource(join(dir, 'source.xlsx'))
      const book = join(dir, 'book.xlsx')
      await writeDependent(book)
      const r = await run(['sheet', 'links', book, '--update', '--json'])
      expect(r.code).toBe(0)
      expect(r.json().detail).toMatchObject({ reachable: 1, broken: 1 })

      const sheet = (await JSZip.loadAsync(readFileSync(book))).file('xl/worksheets/sheet1.xml')!
      const xml = await sheet.async('string')
      // live values cached, formula spellings intact
      expect(xml).toMatch(/<f>\[2\]Sheet1!A1\*2<\/f><v>20<\/v>/)
      expect(xml).toMatch(/<f>MATCH\(20,\[2\]Sheet1!A1:A3,0\)<\/f><v>2<\/v>/)
      // the broken link's cell keeps its cached value
      expect(xml).toMatch(/<f>\[1\]Sheet1!A1&amp;"!"<\/f><v>gone!<\/v>/)
    },
  )

  it.skipIf(!sidecar)(
    'break renumbers surviving links and turns broken formulas into literals',
    async () => {
      const dir = tempDir()
      await writeSource(join(dir, 'source.xlsx'))
      const book = join(dir, 'book.xlsx')
      await writeDependent(book)
      const r = await run(['sheet', 'links', book, '--break', '1', '--json'])
      expect(r.code).toBe(0)

      const zip = await JSZip.loadAsync(readFileSync(book))
      const workbook = await zip.file('xl/workbook.xml')!.async('string')
      expect(workbook).toContain('<externalReference r:id="rId5"/>')
      expect(workbook).not.toContain('rId4')
      expect([...workbook.matchAll(/<externalReference\b/g)]).toHaveLength(1)
      expect(Object.keys(zip.files).includes('xl/externalLinks/externalLink1.xml')).toBe(false)
      const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
      expect(contentTypes).not.toContain('externalLink1.xml')

      const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      // [2] renumbered to [1]; the broken link's formula is a literal
      expect(xml).toContain('<f>[1]Sheet1!A1*2</f>')
      expect(xml).not.toContain('[2]Sheet1')
      expect(xml).not.toContain('<f>[1]Sheet1!A1&amp;"!"</f>')
      expect(xml).toContain('gone!')
    },
  )

  it('writes --break through --out/--force and leaves the source untouched', async () => {
    const dir = tempDir()
    const book = join(dir, 'book.xlsx')
    await writeDependent(book)
    const before = readFileSync(book)
    const copy = join(dir, 'copy.xlsx')
    const r = await run(['sheet', 'links', book, '--break', '1', '--out', copy, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().output_path).toBe(copy)
    // in-place stays the default: --out leaves the input as it was
    expect(readFileSync(book).equals(before)).toBe(true)
    const zip = await JSZip.loadAsync(readFileSync(copy))
    expect(await zip.file('xl/workbook.xml')!.async('string')).not.toContain('rId4')
    // an existing --out is refused without --force, like sheet apply
    const refused = await run(['sheet', 'links', book, '--break', '1', '--out', copy, '--json'])
    expect(refused.code).toBe(2)
    expect(refused.json().message).toContain('use --force')
    expect(
      (await run(['sheet', 'links', book, '--break', '1', '--out', copy, '--force'])).code,
    ).toBe(0)
  })

  it('counts a duplicated --break index once', async () => {
    const dir = tempDir()
    const book = join(dir, 'book.xlsx')
    await writeDependent(book)
    const r = await run(['sheet', 'links', book, '--break', '1,1', '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail).toMatchObject({ broken: [1], removedLinks: 1 })
  })
})

describe('breaking links with references outside the cell formulas', () => {
  it('refuses a defined name that points at the link being broken', async () => {
    const dir = tempDir()
    const book = join(dir, 'book.xlsx')
    await writeDependent(book)
    // point a defined name at link [1] (the one about to be broken)
    const zip = await JSZip.loadAsync(readFileSync(book))
    const workbook = (await zip.file('xl/workbook.xml')!.async('string')).replace(
      '<sheets>',
      '<definedNames><definedName name="Total">[1]Sheet1!$A$1</definedName></definedNames><sheets>',
    )
    zip.file('xl/workbook.xml', workbook)
    writeFileSync(book, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))

    const r = await run(['sheet', 'links', book, '--break', '1', '--json'])
    expect(r.code).not.toBe(0)
    expect(r.json().message).toMatch(/defined name/)
    // refused before writing: the link is still there
    const after = await JSZip.loadAsync(readFileSync(book))
    expect(await after.file('xl/workbook.xml')!.async('string')).toContain('rId4')
  })
})

describe('update writes the live value even when it is empty or spells TRUE', () => {
  /// A one-sheet source workbook whose A1 is the given cell XML.
  async function writeSourceWithA1(path: string, cell: string): Promise<void> {
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      XML_DECL +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    )
    zip.file(
      '_rels/.rels',
      XML_DECL +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    )
    zip.file(
      'xl/workbook.xml',
      XML_DECL +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      XML_DECL +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    )
    zip.file('xl/styles.xml', XML_DECL + STYLES)
    zip.file(
      'xl/worksheets/sheet1.xml',
      XML_DECL +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">' +
        cell +
        '</row></sheetData></worksheet>',
    )
    writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
  }

  /// A dependent whose A1 reads link [1]'s A1 and caches the given cell XML.
  async function writeDependent(path: string, cachedCell: string): Promise<void> {
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      XML_DECL +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>',
    )
    zip.file(
      '_rels/.rels',
      XML_DECL +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    )
    zip.file(
      'xl/workbook.xml',
      XML_DECL +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><externalReferences><externalReference r:id="rId4"/></externalReferences></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      XML_DECL +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>',
    )
    zip.file('xl/styles.xml', XML_DECL + STYLES)
    zip.file(
      'xl/worksheets/sheet1.xml',
      XML_DECL +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">' +
        cachedCell +
        '</row></sheetData></worksheet>',
    )
    zip.file(
      'xl/externalLinks/externalLink1.xml',
      XML_DECL +
        '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames></externalBook></externalLink>',
    )
    zip.file(
      'xl/externalLinks/_rels/externalLink1.xml.rels',
      XML_DECL +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="source.xlsx" TargetMode="External"/></Relationships>',
    )
    writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
  }

  it.skipIf(!sidecar)('clears a stale cache when the live value is the empty string', async () => {
    const dir = tempDir()
    // an empty source cell: the live value is the empty string
    await writeSourceWithA1(join(dir, 'source.xlsx'), '')
    const book = join(dir, 'book.xlsx')
    await writeDependent(book, '<c r="A1"><f>[1]Sheet1!A1</f><v>999</v></c>')
    const r = await run(['sheet', 'links', book, '--update', '--json'])
    expect(r.code).toBe(0)
    const xml = await (
      await JSZip.loadAsync(readFileSync(book))
    )
      .file('xl/worksheets/sheet1.xml')!
      .async('string')
    // the formula stays, the stale 999 is gone (the empty result has no cache)
    expect(xml).toContain('<f>[1]Sheet1!A1</f>')
    expect(xml).not.toContain('<v>999</v>')
  })

  it.skipIf(!sidecar)('keeps the text TRUE as text instead of caching a boolean', async () => {
    const dir = tempDir()
    await writeSourceWithA1(
      join(dir, 'source.xlsx'),
      '<c r="A1" t="inlineStr"><is><t>TRUE</t></is></c>',
    )
    const book = join(dir, 'book.xlsx')
    await writeDependent(book, '<c r="A1" t="str"><f>[1]Sheet1!A1</f><v>TRUE</v></c>')
    const r = await run(['sheet', 'links', book, '--update', '--json'])
    expect(r.code).toBe(0)
    const xml = await (
      await JSZip.loadAsync(readFileSync(book))
    )
      .file('xl/worksheets/sheet1.xml')!
      .async('string')
    // the cached type came from the file's own cache, not from the TRUE spelling
    expect(xml).toMatch(/<c r="A1" t="str"><f>\[1\]Sheet1!A1<\/f><v>TRUE<\/v><\/c>/)
    expect(xml).not.toMatch(/t="b"/)
  })
})
