import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  customPropertyKind,
  parseDocx,
  parsePropertyChanges,
  saveDocx,
  type PropertyChanges,
  type SaveBlock,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const CORE_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
  'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
  'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
  '<dc:title>\u6d4b\u8bd5\u6587\u6863</dc:title><dc:subject>\u6848\u53f7 2026-114</dc:subject><dc:creator>\u4f5c\u8005\u7532</dc:creator>' +
  '<cp:keywords>\u4fdd\u5bc6</cp:keywords><cp:lastModifiedBy>\u4f5c\u8005\u7532</cp:lastModifiedBy><cp:revision>3</cp:revision>' +
  '<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:created>' +
  '<dcterms:modified xsi:type="dcterms:W3CDTF">2020-01-02T00:00:00Z</dcterms:modified>' +
  '<cp:lastPrinted>2020-01-03T00:00:00Z</cp:lastPrinted>' +
  '</cp:coreProperties>'

const APP_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
  'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
  '<Template>Normal.dotm</Template><Company>\u67d0\u67d0\u5f8b\u5e08\u4e8b\u52a1\u6240</Company><Application>Microsoft Office Word</Application>' +
  '<AppVersion>16.0000</AppVersion>' +
  '<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Filing</vt:lpstr></vt:vector></TitlesOfParts>' +
  '</Properties>'

const CUSTOM_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ' +
  'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="MatterId">' +
  '<vt:lpwstr>M-2026-114</vt:lpwstr></property>' +
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Reviewer">' +
  '<vt:lpwstr>\u4e59\u5f8b\u5e08</vt:lpwstr></property>' +
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Pages"><vt:i4>42</vt:i4></property>' +
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="Final"><vt:bool>true</vt:bool></property>' +
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="Filed">' +
  '<vt:filetime>2026-01-02T00:00:00Z</vt:filetime></property>' +
  '</Properties>'

const BODY = '<w:p><w:r><w:t>\u6b63\u6587</w:t></w:r></w:p>'

const docxWithProps = () =>
  buildDocx({
    bodyXml: BODY,
    extraParts: [
      {
        path: 'docProps/core.xml',
        xml: CORE_XML,
        contentType: 'application/vnd.openxmlformats-package.core-properties+xml',
      },
      {
        path: 'docProps/app.xml',
        xml: APP_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
      },
      {
        path: 'docProps/custom.xml',
        xml: CUSTOM_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.custom-properties+xml',
      },
    ],
  })

const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))

const partOf = async (bytes: Uint8Array, path: string): Promise<string> => {
  const file = (await JSZip.loadAsync(bytes)).file(path)
  return file ? file.async('string') : ''
}

describe('document properties', () => {
  it('reads the core, app and custom properties of a package', async () => {
    const doc = await parseDocx(await docxWithProps())
    expect(doc.properties.core).toEqual({
      title: '\u6d4b\u8bd5\u6587\u6863',
      subject: '\u6848\u53f7 2026-114',
      creator: '\u4f5c\u8005\u7532',
      keywords: '\u4fdd\u5bc6',
      lastModifiedBy: '\u4f5c\u8005\u7532',
      revision: '3',
      created: '2020-01-01T00:00:00Z',
      modified: '2020-01-02T00:00:00Z',
    })
    expect(doc.properties.app).toEqual({
      template: 'Normal.dotm',
      company: '\u67d0\u67d0\u5f8b\u5e08\u4e8b\u52a1\u6240',
      application: 'Microsoft Office Word',
      appVersion: '16.0000',
    })
    expect(doc.properties.custom).toEqual({
      MatterId: 'M-2026-114',
      Reviewer: '\u4e59\u5f8b\u5e08',
      Pages: '42',
      Final: 'true',
      Filed: '2026-01-02T00:00:00Z',
    })
    // additive type map: the vt element each custom value is stored as
    expect(doc.properties.customTypes).toEqual({
      MatterId: 'lpwstr',
      Reviewer: 'lpwstr',
      Pages: 'i4',
      Final: 'bool',
      Filed: 'filetime',
    })
  })

  it('keeps imported Word-style application metadata byte-identical on a content save', async () => {
    const input = await docxWithProps()
    const before = input.slice()
    const doc = await parseDocx(input)
    const saved = await saveDocx(
      doc,
      [
        ...originalOrder(doc),
        { kind: 'generated', block: { type: 'paragraph', runs: [{ text: 'Added content' }] } },
      ],
      { savedAt: '2026-07-28T08:30:00Z' },
    )
    expect(await partOf(saved, 'docProps/app.xml')).toBe(await partOf(input, 'docProps/app.xml'))
    expect((await parseDocx(saved)).properties.app.appVersion).toBe('16.0000')
    expect(input).toEqual(before)
  })

  it('a document without docProps has no properties but still opens', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    expect(doc.properties).toEqual({ core: {}, app: {}, custom: {}, customTypes: {} })
  })

  it('sets a property and leaves every other one byte-identical', async () => {
    const doc = await parseDocx(await docxWithProps())
    const out = await saveDocx(doc, originalOrder(doc), {
      savedAt: '2026-07-28T08:30:00Z',
      properties: {
        core: { title: 'Motion to Dismiss' },
        app: { company: 'Barnes & Cole LLP' },
        custom: { MatterId: 'M-2026-999' },
      },
    })
    const core = await partOf(out, 'docProps/core.xml')
    expect(core).toContain('<dc:title>Motion to Dismiss</dc:title>')
    expect(core).toContain('<dc:subject>\u6848\u53f7 2026-114</dc:subject>')
    expect(core).toContain('<cp:keywords>\u4fdd\u5bc6</cp:keywords>')
    expect(core).toContain(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:created>',
    )
    const app = await partOf(out, 'docProps/app.xml')
    expect(app).toContain('<Company>Barnes &amp; Cole LLP</Company>')
    expect(app).toContain('<Template>Normal.dotm</Template>')
    expect(app).toContain('<Application>Microsoft Office Word</Application>')
    const custom = await partOf(out, 'docProps/custom.xml')
    expect(custom).toContain('<vt:lpwstr>M-2026-999</vt:lpwstr>')
    expect(custom).toContain('<vt:lpwstr>\u4e59\u5f8b\u5e08</vt:lpwstr>')
    expect(custom).toContain('name="MatterId"')
  })

  it('removes the property whose value is null', async () => {
    const doc = await parseDocx(await docxWithProps())
    const out = await saveDocx(doc, originalOrder(doc), {
      properties: {
        core: { creator: null, lastModifiedBy: null },
        app: { company: null },
        custom: { Reviewer: null },
      },
    })
    const core = await partOf(out, 'docProps/core.xml')
    expect(core).not.toContain('dc:creator')
    expect(core).not.toContain('lastModifiedBy')
    expect(core).toContain('<dc:title>\u6d4b\u8bd5\u6587\u6863</dc:title>')
    expect(core).toContain('<cp:keywords>\u4fdd\u5bc6</cp:keywords>')
    const app = await partOf(out, 'docProps/app.xml')
    expect(app).not.toContain('Company')
    expect(app).toContain('<Template>Normal.dotm</Template>')
    const custom = await partOf(out, 'docProps/custom.xml')
    expect(custom).not.toContain('Reviewer')
    expect(custom).toContain('M-2026-114')
  })

  it('clears every property of the three parts', async () => {
    const doc = await parseDocx(await docxWithProps())
    const out = await saveDocx(doc, originalOrder(doc), { properties: { clear: 'all' } })
    const core = await partOf(out, 'docProps/core.xml')
    expect(core).not.toContain('dc:title')
    expect(core).not.toContain('dc:creator')
    expect(core).not.toContain('cp:keywords')
    expect(core).not.toContain('lastPrinted')
    expect(core).toContain('<cp:coreProperties')
    const app = await partOf(out, 'docProps/app.xml')
    expect(app).not.toContain('Company')
    expect(app).not.toContain('Template')
    expect(app).not.toContain('TitlesOfParts')
    expect(app).toContain('<Properties')
    const custom = await partOf(out, 'docProps/custom.xml')
    expect(custom).not.toContain('MatterId')
    expect(custom).not.toContain('Reviewer')
    expect(custom).toContain('<Properties')
  })

  it('clears every property and still sets the values given in the same op', async () => {
    const doc = await parseDocx(await docxWithProps())
    const out = await saveDocx(doc, originalOrder(doc), {
      properties: { clear: 'all', core: { title: 'Motion to Dismiss' } },
    })
    const core = await partOf(out, 'docProps/core.xml')
    expect(core).not.toContain('dc:creator')
    expect(core).not.toContain('cp:keywords')
    expect(core).toContain('<dc:title>Motion to Dismiss</dc:title>')
  })

  it('a properties-only save writes the file instead of returning the original bytes', async () => {
    const bytes = await docxWithProps()
    const doc = await parseDocx(bytes)
    const out = await saveDocx(doc, originalOrder(doc), {
      properties: { core: { title: 'Motion to Dismiss' } },
    })
    expect(out).not.toEqual(bytes)
    expect(await partOf(out, 'docProps/core.xml')).toContain('Motion to Dismiss')
  })

  it('creates the docProps part of a document that has none', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const out = await saveDocx(doc, originalOrder(doc), {
      properties: { core: { title: '\u65b0\u5efa\u6587\u6863' } },
    })
    const zip = await JSZip.loadAsync(out)
    const core = await zip.file('docProps/core.xml')!.async('string')
    expect(core).toContain('<dc:title>\u65b0\u5efa\u6587\u6863</dc:title>')
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
      'PartName="/docProps/core.xml"',
    )
    expect(await zip.file('_rels/.rels')!.async('string')).toContain('Target="docProps/core.xml"')
  })

  it('a document without docProps still saves', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const out = await saveDocx(doc, originalOrder(doc), { watermark: '\u8349\u7a3f' })
    expect((await JSZip.loadAsync(out)).file('docProps/core.xml')).toBeNull()
  })
})

describe('typed custom properties', () => {
  const saveWith = async (properties: PropertyChanges): Promise<string> => {
    const doc = await parseDocx(await docxWithProps())
    const out = await saveDocx(doc, originalOrder(doc), { properties })
    return partOf(out, 'docProps/custom.xml')
  }

  it('keeps the vt type of an existing property for a compatible value', async () => {
    const doc = await parseDocx(await docxWithProps())
    const out = await saveDocx(doc, originalOrder(doc), {
      properties: {
        custom: { Pages: 7, Final: false, Filed: '2026-10-01T00:00:00Z' },
      },
    })
    const custom = await partOf(out, 'docProps/custom.xml')
    expect(custom).toContain('<vt:i4>7</vt:i4>')
    expect(custom).toContain('<vt:bool>false</vt:bool>')
    expect(custom).toContain('<vt:filetime>2026-10-01T00:00:00Z</vt:filetime>')
    const reparsed = await parseDocx(out)
    expect(reparsed.properties.customTypes).toMatchObject({
      Pages: 'i4',
      Final: 'bool',
      Filed: 'filetime',
    })
    expect(reparsed.properties.custom).toMatchObject({
      Pages: '7',
      Final: 'false',
      Filed: '2026-10-01T00:00:00Z',
    })
  })

  it('accepts numeric text for an existing integer property', async () => {
    const custom = await saveWith({ custom: { Pages: '7' } })
    expect(custom).toContain('<vt:i4>7</vt:i4>')
    expect(custom).not.toContain('<vt:lpwstr>7</vt:lpwstr>')
  })

  it('falls back to lpwstr when the value does not fit the stored type', async () => {
    const doc = await parseDocx(await docxWithProps())
    const out = await saveDocx(doc, originalOrder(doc), {
      properties: { custom: { Pages: 'seven', Final: 'yes', Filed: 'later' } },
    })
    const custom = await partOf(out, 'docProps/custom.xml')
    expect(custom).toContain('<vt:lpwstr>seven</vt:lpwstr>')
    expect(custom).toContain('<vt:lpwstr>yes</vt:lpwstr>')
    expect(custom).toContain('<vt:lpwstr>later</vt:lpwstr>')
    expect((await parseDocx(out)).properties.customTypes).toMatchObject({
      Pages: 'lpwstr',
      Final: 'lpwstr',
      Filed: 'lpwstr',
    })
  })

  it('types a new custom property from its value', async () => {
    const custom = await saveWith({
      custom: { Count: 42, Approx: 1.5, Big: 5000000000, Approved: true, Label: 'x' },
    })
    expect(custom).toContain('<vt:i4>42</vt:i4>')
    expect(custom).toContain('<vt:r8>1.5</vt:r8>')
    expect(custom).toContain('<vt:i8>5000000000</vt:i8>')
    expect(custom).toContain('<vt:bool>true</vt:bool>')
    expect(custom).toContain('<vt:lpwstr>x</vt:lpwstr>')
  })

  it('leaves untouched typed properties byte-identical', async () => {
    const before = await partOf(await docxWithProps(), 'docProps/custom.xml')
    const after = await saveWith({ custom: { MatterId: 'M-2026-999' } })
    for (const name of ['Pages', 'Final', 'Filed']) {
      const block = (xml: string): string =>
        new RegExp(`<property\\b[^>]*name="${name}"[^>]*>[\\s\\S]*?</property>`).exec(xml)?.[0] ??
        ''
      expect(block(after)).toBe(block(before))
      expect(block(after)).not.toBe('')
    }
  })

  it('stores a value as the kind asked for, whatever its type was', async () => {
    const custom = await saveWith({
      custom: {
        Due: '2026-11-17',
        Budget: '1250.5',
        Copies: '3',
        Signed: 'true',
        Code: '42',
        // an existing text property turned into a number, an existing number into text
        MatterId: '114',
        Pages: '42',
        // the stored type holds while the kind is the same
        Filed: '2026-12-01',
      },
      customKinds: {
        Due: 'date',
        Budget: 'number',
        Copies: 'number',
        Signed: 'bool',
        Code: 'text',
        MatterId: 'number',
        Pages: 'text',
        Filed: 'date',
      },
    })
    const value = (name: string): string =>
      new RegExp(`<property\\b[^>]*name="${name}"[^>]*>([\\s\\S]*?)</property>`).exec(
        custom,
      )?.[1] ?? ''
    expect(value('Due')).toBe('<vt:filetime>2026-11-17T00:00:00Z</vt:filetime>')
    expect(value('Budget')).toBe('<vt:r8>1250.5</vt:r8>')
    expect(value('Copies')).toBe('<vt:i4>3</vt:i4>')
    expect(value('Signed')).toBe('<vt:bool>true</vt:bool>')
    expect(value('Code')).toBe('<vt:lpwstr>42</vt:lpwstr>')
    expect(value('MatterId')).toBe('<vt:i4>114</vt:i4>')
    expect(value('Pages')).toBe('<vt:lpwstr>42</vt:lpwstr>')
    expect(value('Filed')).toBe('<vt:filetime>2026-12-01T00:00:00Z</vt:filetime>')
    // a changed property keeps its pid; the new ones follow the largest
    expect(custom).toMatch(/pid="2" name="MatterId"/)
    expect(custom).toMatch(/pid="7" name="Due"/)
  })

  it('a value that does not fit the kind asked for is kept as text', async () => {
    const custom = await saveWith({
      custom: { Due: 'next week', Copies: 'three' },
      customKinds: { Due: 'date', Copies: 'number' },
    })
    expect(custom).toContain('<vt:lpwstr>next week</vt:lpwstr>')
    expect(custom).toContain('<vt:lpwstr>three</vt:lpwstr>')
  })

  it('reads a stored vt type as one of the four kinds of the Custom tab', () => {
    expect(customPropertyKind('lpwstr')).toBe('text')
    expect(customPropertyKind('i8')).toBe('number')
    expect(customPropertyKind('r8')).toBe('number')
    expect(customPropertyKind('bool')).toBe('bool')
    expect(customPropertyKind('filetime')).toBe('date')
    expect(customPropertyKind('date')).toBe('date')
    expect(customPropertyKind('vector')).toBeNull()
    expect(customPropertyKind(undefined)).toBeNull()
  })

  it('creates the custom part with its content type and relationship when the package has none', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const out = await saveDocx(doc, originalOrder(doc), {
      properties: {
        custom: { Client: 'Example Ltd', Due: '2026-11-17' },
        customKinds: { Due: 'date' },
      },
    })
    const custom = await partOf(out, 'docProps/custom.xml')
    expect(custom).toContain(
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Client"><vt:lpwstr>Example Ltd</vt:lpwstr></property>',
    )
    expect(custom).toContain('pid="3" name="Due"><vt:filetime>2026-11-17T00:00:00Z</vt:filetime>')
    expect(await partOf(out, '[Content_Types].xml')).toContain(
      '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>',
    )
    expect(await partOf(out, '_rels/.rels')).toMatch(
      /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/custom-properties" Target="docProps\/custom\.xml"/,
    )
    const reparsed = await parseDocx(out)
    expect(reparsed.properties.custom).toEqual({
      Client: 'Example Ltd',
      Due: '2026-11-17T00:00:00Z',
    })
    expect(reparsed.properties.customTypes).toEqual({ Client: 'lpwstr', Due: 'filetime' })
  })

  it('a change to one custom property leaves core.xml, app.xml and the other properties byte for byte', async () => {
    const source = await docxWithProps()
    const doc = await parseDocx(source)
    const out = await saveDocx(doc, originalOrder(doc), {
      properties: { custom: { Reviewer: null, Client: 'Example Ltd' } },
    })
    const before = await partOf(source, 'docProps/custom.xml')
    const after = await partOf(out, 'docProps/custom.xml')
    const removed = /<property\b[^>]*name="Reviewer"[^>]*>[\s\S]*?<\/property>/.exec(before)![0]
    const added =
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7" name="Client"><vt:lpwstr>Example Ltd</vt:lpwstr></property>'
    expect(after).toBe(
      before.replace(removed, '').replace('</Properties>', added + '</Properties>'),
    )
    expect(await partOf(out, 'docProps/app.xml')).toBe(await partOf(source, 'docProps/app.xml'))
  })

  it('removes a typed property with null', async () => {
    const custom = await saveWith({ custom: { Pages: null } })
    expect(custom).not.toContain('name="Pages"')
    expect(custom).toContain('name="Final"')
  })

  it('parses typed values in the op and rejects non-scalar ones', () => {
    expect(parsePropertyChanges({ custom: { Pages: 7, Approved: false } })).toEqual({
      custom: { Pages: 7, Approved: false },
    })
    expect(() => parsePropertyChanges({ custom: { X: { nested: true } } })).toThrow(
      'must be a string, number, boolean or null',
    )
    expect(() => parsePropertyChanges({ core: { title: 7 } })).toThrow('must be a string or null')
  })
})
