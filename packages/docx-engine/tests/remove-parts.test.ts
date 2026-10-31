/**
 * remove_parts final-package pass: customXml stores and web-extension parts are
 * dropped with their relationships, content-type entries and empty folders,
 * every other part keeps its bytes, and content controls bound to a removed
 * customXml store refuse the save unless dataBinding: "strip" unbinds them.
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, planRemoveParts, saveDocx } from '../src/index'
import type { RemovePartsOptions } from '../src/index'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const STORE_ITEM_ID = '{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}'

const CONTENT_TYPES_XML =
  XML_DECL +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Default Extension="png" ContentType="image/png"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
  '<Override PartName="/customXml/item1.xml" ContentType="application/xml"/>' +
  '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>' +
  '<Override PartName="/word/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>' +
  '<Override PartName="/word/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.taskpanes+xml"/>' +
  '</Types>'

const ROOT_RELS_XML =
  XML_DECL +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
  '</Relationships>'

const DOCUMENT_RELS_XML =
  XML_DECL +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextensions/webextension1.xml"/>' +
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
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

const CORE_PROPS_XML =
  XML_DECL +
  '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
  'xmlns:dc="http://purl.org/dc/elements/1.1/">' +
  '<dc:title>Matter memo</dc:title></cp:coreProperties>'

const STYLES_XML =
  XML_DECL +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '</w:styles>'

const CUSTOM_XML_ITEM_XML =
  XML_DECL + '<matter xmlns="urn:example:matter"><client>Acme</client></matter>'

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
  '<w:dataBinding w:prefixMappings="xmlns:ns0=&#39;urn:example:matter&#39;" ' +
  `w:xpath="/ns0:matter/ns0:client" w:storeItemID="${STORE_ITEM_ID}"/></w:sdtPr>` +
  '<w:sdtContent><w:p><w:r><w:t>Acme</w:t></w:r></w:p></w:sdtContent></w:sdt>'

function documentXml(withDataBinding: boolean): string {
  return (
    XML_DECL +
    `<w:document ${W_NS}><w:body>` +
    '<w:p><w:r><w:t>Body</w:t></w:r></w:p>' +
    (withDataBinding ? DATA_BINDING_SDT_XML : '') +
    '<w:p><w:r><w:t>Tail</w:t></w:r></w:p>' +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>'
  )
}

/** A .docx carrying a customXml store (both parts + rels), a web-extension part
 *  (+ taskpanes + rels), a lone png (explicit-path case) and a body control
 *  optionally bound to the customXml store. */
async function fixtureDocx(opts: { dataBinding?: boolean } = {}): Promise<Uint8Array> {
  const zip = new JSZip()
  // explicit folder entries, as some producers write them: they must not survive
  zip.folder('customXml')
  zip.folder('customXml/_rels')
  zip.folder('word/webextensions')
  zip.file('[Content_Types].xml', CONTENT_TYPES_XML)
  zip.file('_rels/.rels', ROOT_RELS_XML)
  zip.file('word/document.xml', documentXml(opts.dataBinding === true))
  zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS_XML)
  zip.file('word/styles.xml', STYLES_XML)
  zip.file('docProps/core.xml', CORE_PROPS_XML)
  zip.file('customXml/item1.xml', CUSTOM_XML_ITEM_XML)
  zip.file('customXml/itemProps1.xml', CUSTOM_XML_ITEM_PROPS_XML)
  zip.file('customXml/_rels/item1.xml.rels', CUSTOM_XML_RELS_XML)
  zip.file('word/webextensions/webextension1.xml', WEB_EXTENSION_XML)
  zip.file('word/webextensions/taskpanes.xml', TASKPANES_XML)
  zip.file('word/webextensions/_rels/webextension1.xml.rels', WEB_EXTENSION_RELS_XML)
  zip.file('word/media/only.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47]))
  zip.forEach((_path, entry) => {
    entry.date = new Date(Date.UTC(2026, 0, 1))
  })
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

async function bytesOf(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(bytes)
  const out = new Map<string, Uint8Array>()
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue
    out.set(name, await entry.async('uint8array'))
  }
  return out
}

async function textOf(bytes: Uint8Array, name: string): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  const file = zip.file(name)
  return file ? file.async('string') : ''
}

/** Save an untouched document body with the requested remove_parts pass. */
async function saveFixture(
  bytes: Uint8Array,
  removeParts: RemovePartsOptions,
): Promise<Uint8Array> {
  const parsed = await parseDocx(bytes)
  const blocks = parsed.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
  return saveDocx(parsed, blocks, { removeParts })
}

function entryAt(bytes: Map<string, Uint8Array>, name: string): Uint8Array | undefined {
  return bytes.get(name)
}

describe('remove_parts package scrub', () => {
  it('plans the customXml family with its rels and the bound content control', async () => {
    const bytes = await fixtureDocx({ dataBinding: true })
    const plan = await planRemoveParts(bytes, { kinds: ['customXml'] })
    expect(plan.removedPaths).toEqual([
      'customXml/_rels/item1.xml.rels',
      'customXml/item1.xml',
      'customXml/itemProps1.xml',
    ])
    expect(plan.unknownPaths).toEqual([])
    expect(plan.dataBindings).toEqual([
      {
        part: 'word/document.xml',
        tag: 'ClientName',
        controlId: '42',
        storeItemID: STORE_ITEM_ID,
        itemPropsPath: 'customXml/itemProps1.xml',
      },
    ])
  })

  it('refuses by default when a removed customXml store backs a control', async () => {
    const bytes = await fixtureDocx({ dataBinding: true })
    await expect(saveFixture(bytes, { kinds: ['customXml'] })).rejects.toThrow(
      /ClientName.*w:id 42.*word\/document\.xml.*storeItemID \{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE\}/,
    )
    // nothing was written: the fixture bytes are still the only artifact
    expect((await textOf(bytes, 'customXml/item1.xml')).includes('Acme')).toBe(true)
  })

  it('strips dataBinding when asked, keeping the control and its text', async () => {
    const bytes = await fixtureDocx({ dataBinding: true })
    const saved = await saveFixture(bytes, { kinds: ['customXml'], dataBinding: 'strip' })
    const body = await textOf(saved, 'word/document.xml')
    expect(body).toContain('<w:sdt>')
    expect(body).toContain('<w:tag w:val="ClientName"/>')
    expect(body).toContain('<w:id w:val="42"/>')
    expect(body).toContain('Acme')
    expect(body).not.toContain('<w:dataBinding')
    const entries = await bytesOf(saved)
    expect(entryAt(entries, 'customXml/item1.xml')).toBeUndefined()
    expect(entryAt(entries, 'customXml/itemProps1.xml')).toBeUndefined()
    expect(entryAt(entries, 'customXml/_rels/item1.xml.rels')).toBeUndefined()
    // the relationship to the store and its content-type overrides are gone too
    const rels = await textOf(saved, 'word/_rels/document.xml.rels')
    expect(rels).not.toContain('customXml')
    expect(rels).toContain('styles.xml')
    const contentTypes = await textOf(saved, '[Content_Types].xml')
    expect(contentTypes).not.toContain('customXml')
    expect(contentTypes).toContain('PartName="/word/document.xml"')
  })

  it('removes customXml and web-extension families plus explicit paths, leaving every other part byte-identical', async () => {
    const bytes = await fixtureDocx()
    const before = await bytesOf(bytes)
    const saved = await saveFixture(bytes, {
      kinds: ['customXml', 'webExtensions'],
      paths: ['word/media/only.png'],
    })
    const after = await bytesOf(saved)
    for (const name of [
      'customXml/item1.xml',
      'customXml/itemProps1.xml',
      'customXml/_rels/item1.xml.rels',
      'word/webextensions/webextension1.xml',
      'word/webextensions/taskpanes.xml',
      'word/webextensions/_rels/webextension1.xml.rels',
      'word/media/only.png',
    ]) {
      expect(after.has(name), name).toBe(false)
    }
    // the removed families' folder entries are gone; folders with remaining parts stay
    const zip = await JSZip.loadAsync(saved)
    const folders = Object.keys(zip.files).filter((name) => name.endsWith('/'))
    expect(folders).not.toContain('customXml/')
    expect(folders).not.toContain('customXml/_rels/')
    expect(folders).not.toContain('word/webextensions/')
    expect(folders).not.toContain('word/webextensions/_rels/')
    expect(folders).toContain('word/')
    // untouched parts keep their exact bytes
    for (const name of ['word/styles.xml', 'docProps/core.xml', '_rels/.rels']) {
      expect(Buffer.from(after.get(name)!)).toEqual(Buffer.from(before.get(name)!))
    }
    // the body is unchanged (no control was bound in this fixture)
    expect(Buffer.from(after.get('word/document.xml')!)).toEqual(
      Buffer.from(before.get('word/document.xml')!),
    )
    // png was only used by the removed part: its Default goes; xml/rels stay
    const contentTypes = await textOf(saved, '[Content_Types].xml')
    expect(contentTypes).not.toContain('Extension="png"')
    expect(contentTypes).toContain('Extension="xml"')
    expect(contentTypes).toContain('Extension="rels"')
    expect(contentTypes).not.toContain('webextension')
  })

  it('reports unknown explicit paths and still applies the kinds', async () => {
    const bytes = await fixtureDocx()
    const plan = await planRemoveParts(bytes, {
      kinds: ['customXml'],
      paths: ['customXml/missing.xml', '[Content_Types].xml', 'docProps/core.xml'],
    })
    expect(plan.unknownPaths).toEqual(['customXml/missing.xml', '[Content_Types].xml'])
    expect(plan.removedPaths).toContain('customXml/item1.xml')
    const saved = await saveFixture(bytes, {
      kinds: ['customXml'],
      paths: ['docProps/core.xml'],
    })
    const entries = await bytesOf(saved)
    expect(entries.has('docProps/core.xml')).toBe(false)
    const rootRels = await textOf(saved, '_rels/.rels')
    expect(rootRels).not.toContain('core-properties')
    expect(rootRels).toContain('word/document.xml')
  })

  it('is idempotent: a second pass over scrubbed output returns the input bytes', async () => {
    const bytes = await fixtureDocx()
    const options: RemovePartsOptions = { kinds: ['customXml', 'webExtensions'] }
    const saved = await saveFixture(bytes, options)
    const plan = await planRemoveParts(saved, options)
    expect(plan.removedPaths).toEqual([])
    expect(await saveFixture(saved, options)).toEqual(saved)
  })
})
