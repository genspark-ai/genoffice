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
      '<vt:lpwstr>M-2026-114</vt:lpwstr></property></Properties>',
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
      custom: { MatterId: 'M-2026-114' },
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
    expect(read.json().detail.properties.custom).toEqual({ MatterId: 'M-2026-999' })
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
  })
})
