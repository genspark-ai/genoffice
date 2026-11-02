/**
 * genoffice docs apply remove_parts: a pre-filing metadata scrub drops customXml
 * stores and web-extension parts (plus their package wiring), lists the parts
 * on --dry-run, refuses when a removed store backs a bound content control
 * unless dataBinding: "strip" unbinds it, and leaves every other part alone.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const STORE_ITEM_ID = '{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}'

const CONTENT_TYPES_XML =
  XML_DECL +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/customXml/item1.xml" ContentType="application/xml"/>' +
  '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>' +
  '<Override PartName="/word/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>' +
  '<Override PartName="/word/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.taskpanes+xml"/>' +
  '</Types>'

const ROOT_RELS_XML =
  XML_DECL +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>'

const DOCUMENT_RELS_XML =
  XML_DECL +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextensions/webextension1.xml"/>' +
  '</Relationships>'

const CUSTOM_XML_RELS_XML =
  XML_DECL +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>' +
  '</Relationships>'

const WEB_EXTENSION_RELS_XML =
  XML_DECL +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextensionTaskpanes" Target="taskpanes.xml"/>' +
  '</Relationships>'

const CUSTOM_XML_ITEM_PROPS_XML =
  XML_DECL +
  `<ds:datastoreItem ds:itemID="${STORE_ITEM_ID}" ` +
  'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">' +
  '<ds:schemaRefs/></ds:datastoreItem>'

const WEB_EXTENSION_XML =
  XML_DECL +
  '<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{11111111-2222-3333-4444-555555555555}">' +
  '<we:reference id="Example.Addin" version="1.0.0.0" store="developer" storeType="application"/></we:webextension>'

const TASKPANES_XML =
  XML_DECL +
  '<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"/>'

const DATA_BINDING_SDT_XML =
  '<w:sdt><w:sdtPr><w:id w:val="42"/><w:tag w:val="ClientName"/>' +
  `<w:dataBinding w:xpath="/ns0:matter/ns0:client" w:storeItemID="${STORE_ITEM_ID}"/></w:sdtPr>` +
  '<w:sdtContent><w:p><w:r><w:t>Acme</w:t></w:r></w:p></w:sdtContent></w:sdt>'

function documentXml(withDataBinding: boolean): string {
  return (
    XML_DECL +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    '<w:p><w:r><w:t>Body</w:t></w:r></w:p>' +
    (withDataBinding ? DATA_BINDING_SDT_XML : '') +
    '<w:p><w:r><w:t>Tail</w:t></w:r></w:p>' +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>'
  )
}

/** A .docx carrying a customXml store (+ rels) and a web-extension family (+ rels). */
async function fixtureDocx(path: string, opts: { dataBinding?: boolean } = {}): Promise<string> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES_XML)
  zip.file('_rels/.rels', ROOT_RELS_XML)
  zip.file('word/document.xml', documentXml(opts.dataBinding === true))
  zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS_XML)
  zip.file(
    'word/styles.xml',
    XML_DECL + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
  )
  zip.file(
    'customXml/item1.xml',
    XML_DECL + '<matter xmlns="urn:example:matter"><client>Acme</client></matter>',
  )
  zip.file('customXml/itemProps1.xml', CUSTOM_XML_ITEM_PROPS_XML)
  zip.file('customXml/_rels/item1.xml.rels', CUSTOM_XML_RELS_XML)
  zip.file('word/webextensions/webextension1.xml', WEB_EXTENSION_XML)
  zip.file('word/webextensions/taskpanes.xml', TASKPANES_XML)
  zip.file('word/webextensions/_rels/webextension1.xml.rels', WEB_EXTENSION_RELS_XML)
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
  return path
}

function writeOps(dir: string, ops: unknown[]): string {
  const opsPath = join(dir, 'ops.json')
  writeFileSync(opsPath, JSON.stringify(ops))
  return opsPath
}

async function zipOf(path: string): Promise<JSZip> {
  return JSZip.loadAsync(readFileSync(path))
}

async function partText(path: string, name: string): Promise<string> {
  const zip = await zipOf(path)
  const file = zip.file(name)
  return file ? file.async('string') : ''
}

function partNames(zip: JSZip): string[] {
  return Object.keys(zip.files).filter((name) => !zip.files[name]!.dir)
}

describe('genoffice docs apply remove_parts', () => {
  it('lists the parts on --dry-run and writes nothing', async () => {
    const dir = tempDir()
    const file = await fixtureDocx(join(dir, 'motion.docx'), { dataBinding: true })
    const ops = writeOps(dir, [
      { op: 'remove_parts', kinds: ['customXml', 'webExtensions'], dataBinding: 'strip' },
    ])
    const before = readFileSync(file)
    const dry = await run(['docs', 'apply', file, '--ops', ops, '--dry-run', '--json'])
    expect(dry.code).toBe(0)
    const plan = (dry.json().detail.plan as string[]).join('\n')
    expect(plan).toContain('customXml/item1.xml')
    expect(plan).toContain('customXml/itemProps1.xml')
    expect(plan).toContain('word/webextensions/webextension1.xml')
    expect(plan).toContain('dataBinding: strip')
    expect(readFileSync(file)).toEqual(before)
  })

  it('strips both part families, their rels and content types, keeping other parts byte-identical', async () => {
    const dir = tempDir()
    const file = await fixtureDocx(join(dir, 'motion.docx'))
    const out = join(dir, 'scrubbed.docx')
    const ops = writeOps(dir, [{ op: 'remove_parts', kinds: ['customXml', 'webExtensions'] }])
    const result = await run(['docs', 'apply', file, '--ops', ops, '--out', out, '--json'])
    expect(result.code).toBe(0)
    expect(result.json().detail.results[0].output).toContain('customXml/item1.xml')
    const zip = await zipOf(out)
    const names = partNames(zip)
    expect(names.some((name) => name.startsWith('customXml/'))).toBe(false)
    expect(names.some((name) => name.startsWith('word/webextensions/'))).toBe(false)
    expect(names).toContain('word/document.xml')
    expect(names).toContain('word/styles.xml')
    // the body was not rewritten: same bytes as the source part
    expect(await partText(out, 'word/document.xml')).toBe(await partText(file, 'word/document.xml'))
    expect(await partText(out, 'word/styles.xml')).toBe(await partText(file, 'word/styles.xml'))
    const contentTypes = await partText(out, '[Content_Types].xml')
    expect(contentTypes).not.toContain('customXml')
    expect(contentTypes).not.toContain('webextension')
    const rels = await partText(out, 'word/_rels/document.xml.rels')
    expect(rels).not.toContain('customXml')
    expect(rels).not.toContain('webextensions')
  })

  it('refuses by default when a content control is bound to a removed customXml store', async () => {
    const dir = tempDir()
    const file = await fixtureDocx(join(dir, 'motion.docx'), { dataBinding: true })
    const out = join(dir, 'scrubbed.docx')
    const ops = writeOps(dir, [{ op: 'remove_parts', kinds: ['customXml'] }])
    const result = await run(['docs', 'apply', file, '--ops', ops, '--out', out, '--json'])
    expect(result.code).toBe(1)
    const body = result.json()
    expect(body.error).toBe('op_rejected')
    expect(body.message).toContain('ClientName')
    expect(body.message).toContain(STORE_ITEM_ID)
    expect(existsSync(out)).toBe(false)
    expect(readFileSync(file)).toEqual(readFileSync(file))
  })

  it('strips the binding when asked, keeping the control and its text', async () => {
    const dir = tempDir()
    const file = await fixtureDocx(join(dir, 'motion.docx'), { dataBinding: true })
    const out = join(dir, 'scrubbed.docx')
    const ops = writeOps(dir, [{ op: 'remove_parts', kinds: ['customXml'], dataBinding: 'strip' }])
    const result = await run(['docs', 'apply', file, '--ops', ops, '--out', out, '--json'])
    expect(result.code).toBe(0)
    expect(result.json().detail.results[0].output).toContain('unbind 1 control')
    const body = await partText(out, 'word/document.xml')
    expect(body).toContain('<w:tag w:val="ClientName"/>')
    expect(body).toContain('Acme')
    expect(body).not.toContain('<w:dataBinding')
    const names = partNames(await zipOf(out))
    expect(names.some((name) => name.startsWith('customXml/'))).toBe(false)
  })

  it('reports unknown explicit paths and is idempotent on scrubbed output', async () => {
    const dir = tempDir()
    const file = await fixtureDocx(join(dir, 'motion.docx'))
    const out = join(dir, 'scrubbed.docx')
    const ops = writeOps(dir, [
      { op: 'remove_parts', kinds: ['customXml'], paths: ['customXml/missing.xml'] },
    ])
    const first = await run(['docs', 'apply', file, '--ops', ops, '--out', out, '--json'])
    expect(first.code).toBe(0)
    expect(first.json().detail.results[0].output).toContain(
      'ignored unknown paths: customXml/missing.xml',
    )
    const out2 = join(dir, 'again.docx')
    const second = await run(['docs', 'apply', out, '--ops', ops, '--out', out2, '--json'])
    expect(second.code).toBe(0)
    // nothing left to remove: the save returns the input bytes unchanged
    expect(readFileSync(out2)).toEqual(readFileSync(out))
    const dry = await run(['docs', 'apply', out, '--ops', ops, '--dry-run', '--json'])
    expect((dry.json().detail.plan as string[]).join('\n')).toContain('nothing (already absent)')
  })

  it('rejects a malformed remove_parts op', async () => {
    const dir = tempDir()
    const file = await fixtureDocx(join(dir, 'motion.docx'))
    const missing = await run([
      'docs',
      'apply',
      file,
      '--ops',
      writeOps(dir, [{ op: 'remove_parts' }]),
      '--dry-run',
      '--json',
    ])
    expect(missing.code).toBe(1)
    expect(missing.json().message).toContain('needs kinds')
    const badKind = await run([
      'docs',
      'apply',
      file,
      '--ops',
      writeOps(dir, [{ op: 'remove_parts', kinds: ['docProps'] }]),
      '--dry-run',
      '--json',
    ])
    expect(badKind.code).toBe(1)
    expect(badKind.json().message).toContain('unknown kind')
  })
})
