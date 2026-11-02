import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'

async function partOf(path: string, name: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  const file = zip.file(name)
  return file ? file.async('string') : ''
}

/** a .docx with the three docProps parts filled in */
async function docxWithProps(path: string): Promise<string> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
      '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'word/document.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body><w:p><w:r><w:t>\u6b63\u6587</w:t></w:r></w:p></w:body></w:document>',
  )
  zip.file(
    'docProps/core.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
      ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
      ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>\u65e7\u6807\u9898</dc:title><dc:creator>\u4f5c\u8005\u7532</dc:creator>' +
      '<dc:subject>\u6848\u53f7 2026-114</dc:subject></cp:coreProperties>',
  )
  zip.file(
    'docProps/app.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
      '<Template>Normal.dotm</Template><Company>\u67d0\u67d0\u5f8b\u5e08\u4e8b\u52a1\u6240</Company></Properties>',
  )
  zip.file(
    'docProps/custom.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"' +
      ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="MatterId">' +
      '<vt:lpwstr>M-2026-114</vt:lpwstr></property>' +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Pages">' +
      '<vt:i4>42</vt:i4></property>' +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Final">' +
      '<vt:bool>true</vt:bool></property>' +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="Filed">' +
      '<vt:filetime>2026-01-02T00:00:00Z</vt:filetime></property></Properties>',
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
  return path
}

describe('genoffice docs (document properties)', () => {
  it('reads the core, app and custom properties of a document', async () => {
    const dir = tempDir()
    const file = await docxWithProps(join(dir, 'motion.docx'))
    const read = await run(['docs', 'read', file, '--properties', '--json'])
    expect(read.code).toBe(0)
    expect(read.json().detail.properties).toEqual({
      core: {
        title: '\u65e7\u6807\u9898',
        creator: '\u4f5c\u8005\u7532',
        subject: '\u6848\u53f7 2026-114',
      },
      app: { template: 'Normal.dotm', company: '\u67d0\u67d0\u5f8b\u5e08\u4e8b\u52a1\u6240' },
      custom: {
        MatterId: 'M-2026-114',
        Pages: '42',
        Final: 'true',
        Filed: '2026-01-02T00:00:00Z',
      },
      customTypes: {
        MatterId: 'lpwstr',
        Pages: 'i4',
        Final: 'bool',
        Filed: 'filetime',
      },
    })
  })

  it('sets and removes properties with a set_properties op', async () => {
    const dir = tempDir()
    const file = await docxWithProps(join(dir, 'motion.docx'))
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        {
          op: 'set_properties',
          core: { title: 'Motion to Dismiss', creator: null },
          app: { company: null },
          custom: { MatterId: 'M-2026-999' },
        },
      ]),
    )
    const applied = await run(['docs', 'apply', file, '--ops', ops, '--json'])
    expect(applied.code).toBe(0)
    const core = await partOf(file, 'docProps/core.xml')
    expect(core).toContain('<dc:title>Motion to Dismiss</dc:title>')
    expect(core).not.toContain('dc:creator')
    expect(core).toContain('<dc:subject>\u6848\u53f7 2026-114</dc:subject>')
    const app = await partOf(file, 'docProps/app.xml')
    expect(app).not.toContain('Company')
    expect(app).toContain('<Template>Normal.dotm</Template>')
    const custom = await partOf(file, 'docProps/custom.xml')
    expect(custom).toContain('<vt:lpwstr>M-2026-999</vt:lpwstr>')
    const read = await run(['docs', 'read', file, '--properties', '--json'])
    expect(read.json().detail.properties.core.title).toBe('Motion to Dismiss')
    expect(read.json().detail.properties.core.creator).toBeUndefined()
    expect(read.json().detail.properties.custom).toMatchObject({
      MatterId: 'M-2026-999',
      Pages: '42',
    })
    expect(read.json().detail.properties.customTypes).toMatchObject({
      MatterId: 'lpwstr',
      Pages: 'i4',
      Final: 'bool',
      Filed: 'filetime',
    })
  })

  it('a properties-only apply writes the file', async () => {
    const dir = tempDir()
    const file = await docxWithProps(join(dir, 'motion.docx'))
    const before = readFileSync(file)
    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'set_properties', core: { title: 'Motion' } }]))
    const applied = await run(['docs', 'apply', file, '--ops', ops, '--json'])
    expect(applied.code).toBe(0)
    expect(readFileSync(file)).not.toEqual(before)
    expect(await partOf(file, 'docProps/core.xml')).toContain('<dc:title>Motion</dc:title>')
  })

  it('clears every property with clear: "all"', async () => {
    const dir = tempDir()
    const file = await docxWithProps(join(dir, 'motion.docx'))
    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'set_properties', clear: 'all' }]))
    const applied = await run(['docs', 'apply', file, '--ops', ops, '--json'])
    expect(applied.code).toBe(0)
    expect(await partOf(file, 'docProps/core.xml')).not.toContain('dc:title')
    expect(await partOf(file, 'docProps/app.xml')).not.toContain('Company')
    expect(await partOf(file, 'docProps/custom.xml')).not.toContain('MatterId')
  })

  it('a later clear: "all" wipes the values of an earlier op in the same batch', async () => {
    const dir = tempDir()
    const file = await docxWithProps(join(dir, 'motion.docx'))
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'set_properties', core: { title: 'Motion' }, custom: { MatterId: 'M-1' } },
        { op: 'set_properties', clear: 'all', core: { subject: 'Filed' } },
      ]),
    )
    const applied = await run(['docs', 'apply', file, '--ops', ops, '--json'])
    expect(applied.code).toBe(0)
    const core = await partOf(file, 'docProps/core.xml')
    expect(core).not.toContain('Motion')
    expect(core).toContain('<dc:subject>Filed</dc:subject>')
    expect(await partOf(file, 'docProps/custom.xml')).not.toContain('MatterId')
  })

  it('rejects an unknown property and leaves the file alone', async () => {
    const dir = tempDir()
    const file = await docxWithProps(join(dir, 'motion.docx'))
    const before = readFileSync(file)
    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'set_properties', core: { nope: 'x' } }]))
    const applied = await run(['docs', 'apply', file, '--ops', ops, '--json'])
    expect(applied.code).toBe(1)
    expect(applied.json().message).toContain('unknown core property')
    expect(readFileSync(file)).toEqual(before)
  })

  it('lists the properties in info --json too', async () => {
    const dir = tempDir()
    const file = await docxWithProps(join(dir, 'motion.docx'))
    const info = await run(['info', file, '--json'])
    expect(info.code).toBe(0)
    expect(info.json().detail.properties.core.title).toBe('\u65e7\u6807\u9898')
    expect(info.json().detail.properties.app.company).toBe(
      '\u67d0\u67d0\u5f8b\u5e08\u4e8b\u52a1\u6240',
    )
    expect(info.json().detail.properties.customTypes).toMatchObject({
      Pages: 'i4',
      Final: 'bool',
      Filed: 'filetime',
    })
  })

  const propertyBlock = (xml: string, name: string): string =>
    new RegExp(`<property\\b[^>]*name="${name}"[^>]*>[\\s\\S]*?</property>`).exec(xml)?.[0] ?? ''

  it('keeps the vt type of an existing typed property when setting it', async () => {
    const dir = tempDir()
    const file = await docxWithProps(join(dir, 'motion.docx'))
    const before = await partOf(file, 'docProps/custom.xml')
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([{ op: 'set_properties', custom: { Pages: 7, Final: false } }]),
    )
    const applied = await run(['docs', 'apply', file, '--ops', ops, '--json'])
    expect(applied.code).toBe(0)
    const custom = await partOf(file, 'docProps/custom.xml')
    expect(custom).toContain('<vt:i4>7</vt:i4>')
    expect(custom).toContain('<vt:bool>false</vt:bool>')
    // untouched typed entries stay byte-identical
    for (const name of ['MatterId', 'Filed']) {
      expect(propertyBlock(custom, name)).toBe(propertyBlock(before, name))
      expect(propertyBlock(custom, name)).not.toBe('')
    }
    const read = await run(['docs', 'read', file, '--properties', '--json'])
    expect(read.json().detail.properties.custom).toMatchObject({ Pages: '7', Final: 'false' })
    expect(read.json().detail.properties.customTypes).toMatchObject({
      Pages: 'i4',
      Final: 'bool',
    })
  })

  it('accepts number and boolean custom values for new properties', async () => {
    const dir = tempDir()
    const file = await docxWithProps(join(dir, 'motion.docx'))
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        {
          op: 'set_properties',
          custom: { Score: 7, Approved: true, Rate: 1.5, Big: 5000000000, Note: 'x' },
        },
      ]),
    )
    expect((await run(['docs', 'apply', file, '--ops', ops, '--json'])).code).toBe(0)
    const custom = await partOf(file, 'docProps/custom.xml')
    expect(custom).toContain('<vt:i4>7</vt:i4>')
    expect(custom).toContain('<vt:bool>true</vt:bool>')
    expect(custom).toContain('<vt:r8>1.5</vt:r8>')
    expect(custom).toContain('<vt:i8>5000000000</vt:i8>')
    expect(custom).toContain('<vt:lpwstr>x</vt:lpwstr>')
    const read = await run(['docs', 'read', file, '--properties', '--json'])
    expect(read.json().detail.properties.customTypes).toMatchObject({
      Score: 'i4',
      Approved: 'bool',
      Rate: 'r8',
      Big: 'i8',
      Note: 'lpwstr',
    })
  })

  it('falls back to lpwstr when an incompatible value meets a typed property', async () => {
    const dir = tempDir()
    const file = await docxWithProps(join(dir, 'motion.docx'))
    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'set_properties', custom: { Pages: 'seven' } }]))
    expect((await run(['docs', 'apply', file, '--ops', ops, '--json'])).code).toBe(0)
    const custom = await partOf(file, 'docProps/custom.xml')
    expect(custom).toContain('<vt:lpwstr>seven</vt:lpwstr>')
    expect(custom).not.toContain('<vt:i4>seven</vt:i4>')
    const read = await run(['docs', 'read', file, '--properties', '--json'])
    expect(read.json().detail.properties.customTypes.Pages).toBe('lpwstr')
  })

  it('rejects a non-scalar custom value and leaves the file alone', async () => {
    const dir = tempDir()
    const file = await docxWithProps(join(dir, 'motion.docx'))
    const before = readFileSync(file)
    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'set_properties', custom: { Nested: { a: 1 } } }]))
    const applied = await run(['docs', 'apply', file, '--ops', ops, '--json'])
    expect(applied.code).toBe(1)
    expect(applied.json().message).toContain('must be a string, number, boolean or null')
    expect(readFileSync(file)).toEqual(before)
  })

  // clear: "all" keeps the three parts (empty roots): Word-compatible, and a
  // later set_properties still lands without re-creating content types/rels.
  it('clear: "all" leaves the empty docProps parts in place', async () => {
    const dir = tempDir()
    const file = await docxWithProps(join(dir, 'motion.docx'))
    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'set_properties', clear: 'all' }]))
    expect((await run(['docs', 'apply', file, '--ops', ops, '--json'])).code).toBe(0)
    const zip = await JSZip.loadAsync(readFileSync(file))
    for (const part of ['docProps/core.xml', 'docProps/app.xml', 'docProps/custom.xml']) {
      expect(zip.file(part), part).not.toBeNull()
    }
    const custom = await partOf(file, 'docProps/custom.xml')
    expect(custom).toContain('<Properties')
    expect(custom).not.toContain('<property')
    expect(await partOf(file, 'docProps/core.xml')).toContain('<cp:coreProperties')
  })

  it('create --type docx from markdown writes docProps core and app parts', async () => {
    const dir = tempDir()
    const md = join(dir, 'brief.md')
    writeFileSync(md, '# Title\n\nBody text.')
    const out = join(dir, 'created.docx')
    const created = await run(['create', '--type', 'docx', '--from', md, '--out', out, '--json'])
    expect(created.code).toBe(0)
    const zip = await JSZip.loadAsync(readFileSync(out))
    const core = await zip.file('docProps/core.xml')!.async('string')
    expect(core).toMatch(
      /<dcterms:created xsi:type="dcterms:W3CDTF">\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z<\/dcterms:created>/,
    )
    expect(core).toMatch(
      /<dcterms:modified xsi:type="dcterms:W3CDTF">\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z<\/dcterms:modified>/,
    )
    expect(core).not.toContain('<dc:creator')
    const app = await zip.file('docProps/app.xml')!.async('string')
    expect(app).toContain('<Application>GenOffice</Application>')
    expect(app).toContain('<Template>Normal.dotm</Template>')
    expect(app).toMatch(/<AppVersion>[^<]+<\/AppVersion>/)
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
      'PartName="/docProps/core.xml"',
    )
    const read = await run(['docs', 'read', out, '--properties', '--json'])
    expect(read.code).toBe(0)
    expect(read.json().detail.properties.core.created).toMatch(/Z$/)
    expect(read.json().detail.properties.app.application).toBe('GenOffice')
    expect(read.json().detail.properties.custom).toEqual({})
    expect(read.json().detail.properties.customTypes).toEqual({})
  })

  it('create --type docx from an HTML fragment writes docProps too', async () => {
    const dir = tempDir()
    const html = join(dir, 'fragment.html')
    writeFileSync(html, '<h1>Title</h1><p>Body text.</p>')
    const out = join(dir, 'created.docx')
    expect(
      (await run(['create', '--type', 'docx', '--from', html, '--out', out, '--json'])).code,
    ).toBe(0)
    const zip = await JSZip.loadAsync(readFileSync(out))
    expect(zip.file('docProps/core.xml')).not.toBeNull()
    expect(await zip.file('docProps/app.xml')!.async('string')).toContain(
      '<Application>GenOffice</Application>',
    )
  })

  it('the docs guide documents typed custom values', async () => {
    const guide = await run(['guide', 'docs', 'set_properties', '--json'])
    expect(guide.code).toBe(0)
    const custom = guide.json().detail.schema.properties.custom
    expect(custom.additionalProperties.type).toEqual(['string', 'number', 'boolean', 'null'])
  })
})
