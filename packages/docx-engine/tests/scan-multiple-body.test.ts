import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { scanBody } from '../src/scan'
import { buildDocx } from './helpers/build-docx'

describe('scanBody with multiple <w:body> elements', () => {
  it('keeps scanning past </w:body> into later sibling bodies (POI MultipleBodyBug)', () => {
    const xml =
      '<w:document>' +
      '<w:body><w:p><w:r><w:t>BODY 1</w:t></w:r></w:p></w:body>' +
      '<w:body><w:p><w:r><w:t>BODY 2</w:t></w:r></w:p></w:body>' +
      '<w:body><w:p><w:r><w:t>BODY 3</w:t></w:r></w:p></w:body>' +
      '</w:document>'
    const scan = scanBody(xml)
    expect(scan.elements.map((e) => e.name)).toEqual(['w:p', 'w:p', 'w:p'])
    expect(xml.slice(scan.elements[2].start, scan.elements[2].end)).toContain('BODY 3')
    expect(scan.innerEnd).toBe(scan.elements[2].end)
  })

  it('single body still terminates at its closing tag', () => {
    const xml = '<w:document><w:body><w:p/></w:body></w:document>'
    expect(scanBody(xml).elements.map((e) => e.name)).toEqual(['w:p'])
  })
})

describe('opaque body regions', () => {
  const opaque =
    '<!-- <w:p><w:r><w:t>PHANTOM</w:t></w:r></w:p> </w:body> -->' +
    '<![CDATA[<w:p><w:r><w:t>CDATA</w:t></w:r></w:p> </w:body>]]>' +
    '<?probe <w:p><w:r><w:t>PI</w:t></w:r></w:p>?>'
  const before = '<w:p><w:r><w:t>before</w:t></w:r></w:p>'
  const after = '<w:p><w:r><w:t>after</w:t></w:r></w:p>'

  it('does not scan tags or body terminators inside comments, CDATA, or processing instructions', () => {
    const xml = `<w:document><w:body>${opaque}${before}</w:body></w:document>`
    const scan = scanBody(xml)

    expect(scan.elements).toHaveLength(1)
    expect(scan.elements[0].name).toBe('w:p')
    expect(xml.slice(scan.elements[0].start, scan.elements[0].end)).toBe(before)
    expect(scan.innerStart).toBe(xml.indexOf(before))
    expect(scan.innerEnd).toBe(xml.indexOf(before) + before.length)
  })

  it('preserves opaque regions when splicing an edited body', async () => {
    const source = await buildDocx({ bodyXml: before + opaque + after })
    const doc = await parseDocx(source)
    expect(
      doc.blocks.filter((block) => !block.hidden).map((block) => block.runs?.[0]?.text),
    ).toEqual(['before', 'after'])

    const saved = await saveDocx(doc, [
      { kind: 'original', docxIndex: 0 },
      {
        kind: 'generated',
        block: { type: 'paragraph', runs: [{ text: 'edited' }] },
      },
    ])
    const reparsed = await parseDocx(saved)

    expect(reparsed.internal.documentXml).toContain(opaque)
    expect(
      reparsed.blocks.filter((block) => !block.hidden).map((block) => block.runs?.[0]?.text),
    ).toEqual(['before', 'edited'])
  })

  it('discovers real sibling bodies past body-looking opaque text', () => {
    const xml =
      '<w:document><!-- <w:body><w:p>fake</w:p></w:body> -->' +
      '<w:body><w:p>one</w:p><!-- </w:body> --></w:body>' +
      '<![CDATA[<w:body><w:p>fake</w:p></w:body>]]>' +
      '<w:body><w:p>two</w:p></w:body></w:document>'
    const scan = scanBody(xml)

    expect(scan.elements.map((element) => element.name)).toEqual(['w:p', 'w:p'])
    expect(xml.slice(scan.elements[0].start, scan.elements[0].end)).toContain('one')
    expect(xml.slice(scan.elements[1].start, scan.elements[1].end)).toContain('two')
    expect(scan.bodyContentStart).toBe(xml.indexOf('>', xml.indexOf('<w:body><w:p>one')) + 1)
    expect(scan.bodyContentEnd).toBe(xml.lastIndexOf('</w:body>'))
    expect(scan.opaqueRegions.map((region) => xml.slice(region.start, region.end))).toContain(
      '<![CDATA[<w:body><w:p>fake</w:p></w:body>]]>',
    )
  })

  it('retains opaque ranges nested inside a body element', () => {
    const xml =
      '<w:document><w:body><w:p><w:r><w:t>x</w:t></w:r>' +
      '<!-- <w:p>fake</w:p> --><w:r><w:t>y</w:t></w:r></w:p></w:body></w:document>'
    const scan = scanBody(xml)

    expect(scan.opaqueRegions).toHaveLength(1)
    expect(scan.opaqueRegions[0]!.start).toBeGreaterThan(scan.elements[0]!.start)
    expect(scan.opaqueRegions[0]!.end).toBeLessThan(scan.elements[0]!.end)
  })

  it('keeps nested opaque markup when regenerating its source element', async () => {
    const nested = '<w:p><w:r><w:t>middle</w:t></w:r><!--NESTED--></w:p>'
    const source = await buildDocx({ bodyXml: before + '<!--TOP-->' + nested + after })
    const doc = await parseDocx(source)
    const saved = await saveDocx(doc, [
      { kind: 'original', docxIndex: 0 },
      {
        kind: 'generated',
        block: { type: 'paragraph', runs: [{ text: 'edited' }] },
        docxIndex: 1,
      },
      { kind: 'original', docxIndex: 2 },
    ])
    const reparsed = await parseDocx(saved)
    const xml = reparsed.internal.documentXml

    expect(xml).toContain('<!--TOP-->')
    expect(xml).toContain('<!--NESTED-->')
    expect(/<w:p>[^]*<!--NESTED-->[^]*<\/w:p>/.test(xml)).toBe(true)
  })

  it('CDATA inside w:t is run text, not opaque decoration: no duplication or hoisting', async () => {
    const cdataPara = '<w:p><w:r><w:t><![CDATA[cdata text]]></w:t></w:r></w:p>'
    const source = await buildDocx({ bodyXml: before + cdataPara + after })
    const doc = await parseDocx(source)
    expect(
      doc.blocks.filter((block) => !block.hidden).map((block) => block.runs?.[0]?.text),
    ).toEqual(['before', 'cdata text', 'after'])

    // regenerating the paragraph: the CDATA payload must not be re-appended
    const edited = await saveDocx(doc, [
      { kind: 'original', docxIndex: 0 },
      {
        kind: 'generated',
        block: { type: 'paragraph', runs: [{ text: 'edited' }] },
        docxIndex: 1,
      },
      { kind: 'original', docxIndex: 2 },
    ])
    const editedDoc = await parseDocx(edited)
    expect(editedDoc.internal.documentXml).not.toContain('<![CDATA[cdata text]]>')
    expect(
      editedDoc.blocks.filter((block) => !block.hidden).map((block) => block.runs?.[0]?.text),
    ).toEqual(['before', 'edited', 'after'])

    // deleting the paragraph: its text must not be hoisted to the start of the body
    const deleted = await saveDocx(doc, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'original', docxIndex: 2 },
    ])
    const deletedDoc = await parseDocx(deleted)
    expect(deletedDoc.internal.documentXml).not.toContain('cdata text')
    expect(
      deletedDoc.blocks.filter((block) => !block.hidden).map((block) => block.runs?.[0]?.text),
    ).toEqual(['before', 'after'])
  })

  it('an untouched paragraph keeps its CDATA text bytes verbatim', async () => {
    const cdataPara = '<w:p><w:r><w:t><![CDATA[cdata text]]></w:t></w:r></w:p>'
    const source = await buildDocx({ bodyXml: cdataPara })
    const doc = await parseDocx(source)
    const saved = await saveDocx(doc, [{ kind: 'original', docxIndex: 0 }])
    const xml = (await parseDocx(saved)).internal.documentXml
    expect(xml.split('<![CDATA[cdata text]]>')).toHaveLength(2)
  })

  it('a comment inside an sdt shell is emitted once when its paragraph is regenerated', async () => {
    const sdt =
      '<w:sdt><w:sdtPr><!--IN-SDTPR--><w:alias w:val="Ctl"/></w:sdtPr>' +
      '<w:sdtContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:sdtContent><!--IN-TAIL--></w:sdt>'
    const source = await buildDocx({ bodyXml: before + sdt + after })
    const doc = await parseDocx(source)
    const block = doc.blocks[1]!
    expect(block.sdtShell?.openXml).toContain('<!--IN-SDTPR-->')
    expect(block.sdtShell?.closeXml).toContain('<!--IN-TAIL-->')

    const saved = await saveDocx(doc, [
      { kind: 'original', docxIndex: 0 },
      {
        kind: 'generated',
        block: { type: 'paragraph', runs: [{ text: 'edited' }], sdtShell: block.sdtShell },
        docxIndex: 1,
      },
      { kind: 'original', docxIndex: 2 },
    ])
    const xml = (await parseDocx(saved)).internal.documentXml
    expect(xml.split('<!--IN-SDTPR-->')).toHaveLength(2)
    expect(xml.split('<!--IN-TAIL-->')).toHaveLength(2)
    expect(xml).toContain('>edited</w:t>')
    expect(xml).not.toContain('inside')
  })

  it('does not remove comment markers from opaque payloads', async () => {
    const bodyXml =
      '<w:p><w:r><w:commentRangeStart w:id="0"/><w:commentReference w:id="0"/></w:r>' +
      '<!-- <w:commentReference w:id="1"/> <w:fldChar w:fldCharType="separate"/> -->' +
      '</w:p>'
    const source = await buildDocx({ bodyXml })
    const doc = await parseDocx(source)
    const saved = await saveDocx(doc, [{ kind: 'original', docxIndex: 0 }], { comments: [] })
    const xml = (await parseDocx(saved)).internal.documentXml

    expect(xml).not.toContain('w:commentRangeStart w:id="0"')
    expect(xml).not.toContain('w:commentReference w:id="0"')
    expect(xml).toContain(
      '<!-- <w:commentReference w:id="1"/> <w:fldChar w:fldCharType="separate"/> -->',
    )
    expect(xml).toContain('w:fldCharType="separate"')
  })
})
