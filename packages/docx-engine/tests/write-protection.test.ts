/**
 * w:writeProtection (password to modify) and w:removePersonalInformation
 * (anonymize authors on save) — parse / patch round-trips and the save-time
 * author scrub.
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { hashProtectionPassword, parseDocx, saveDocx, verifyProtectionPassword } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

function originalBlocks(parsed: Awaited<ReturnType<typeof parseDocx>>) {
  return parsed.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
}

describe('writeProtection (password to modify)', () => {
  it('round-trips: saved credentials verify the right password and land in settings.xml', async () => {
    const creds = await hashProtectionPassword('m0dify', 1000)
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }),
    )
    expect(parsed.writeProtection).toBeNull()
    const restriction = await hashProtectionPassword('restrict', 1000)
    const saved = await saveDocx(parsed, originalBlocks(parsed), {
      writeProtection: creds,
      protection: { edit: 'trackedChanges', enforced: true, ...restriction },
    })
    const reparsed = await parseDocx(saved)
    expect(reparsed.writeProtection).toMatchObject({ spinCount: 1000, algorithmSid: 14 })
    expect(reparsed.writeProtection?.hash).toBe(creds.hash)
    expect(await verifyProtectionPassword('m0dify', reparsed.writeProtection!)).toBe(true)
    expect(await verifyProtectionPassword('nope', reparsed.writeProtection!)).toBe(false)
    expect(reparsed.protection?.edit).toBe('trackedChanges')

    // CT_Settings sequence: writeProtection must come before documentProtection
    const settings = await (await JSZip.loadAsync(saved)).file('word/settings.xml')!.async('string')
    expect(settings.indexOf('<w:writeProtection')).toBeGreaterThan(-1)
    expect(settings.indexOf('<w:writeProtection')).toBeLessThan(
      settings.indexOf('<w:documentProtection'),
    )
  })

  it('null removes an existing writeProtection', async () => {
    const creds = await hashProtectionPassword('m0dify', 1000)
    const first = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }),
    )
    const withWp = await parseDocx(
      await saveDocx(first, originalBlocks(first), { writeProtection: creds }),
    )
    expect(withWp.writeProtection?.hash).toBe(creds.hash)
    const removed = await parseDocx(
      await saveDocx(withWp, originalBlocks(withWp), { writeProtection: null }),
    )
    expect(removed.writeProtection).toBeNull()
  })
})

describe('removePersonalInformation', () => {
  const bodyWithRevision =
    '<w:p><w:ins w:id="1" w:author="张三" w:date="2026-01-01T00:00:00Z">' +
    '<w:r><w:t>inserted</w:t></w:r></w:ins></w:p>'

  const commentsPart = {
    path: 'word/comments.xml',
    xml:
      `${XML_DECL}<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      '<w:comment w:id="0" w:author="张三" w:initials="ZS"><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:comment>' +
      '</w:comments>',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
  }
  const commentsRel =
    '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'

  const corePart = {
    path: 'docProps/core.xml',
    xml:
      `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">' +
      '<dc:creator>张三</dc:creator><cp:lastModifiedBy>李四</cp:lastModifiedBy>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">2026-01-01T00:00:00Z</dcterms:modified>' +
      '</cp:coreProperties>',
    contentType: 'application/vnd.openxmlformats-package.core-properties+xml',
  }

  it('flag round-trips and scrubs revision/comment authors and core props on save', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: bodyWithRevision,
        extraRels: commentsRel,
        extraParts: [commentsPart, corePart],
      }),
    )
    expect(parsed.removePersonalInfo).toBe(false)
    const saved = await saveDocx(parsed, originalBlocks(parsed), { removePersonalInfo: true })
    const reparsed = await parseDocx(saved)
    expect(reparsed.removePersonalInfo).toBe(true)

    const zip = await JSZip.loadAsync(saved)
    const docXml = await zip.file('word/document.xml')!.async('string')
    expect(docXml).toContain('w:author="Author"')
    expect(docXml).not.toContain('张三')
    // comments.xml is scrubbed even though comments were not edited
    const commentsXml = await zip.file('word/comments.xml')!.async('string')
    expect(commentsXml).toContain('w:author="Author"')
    expect(commentsXml).toContain('w:initials="A"')
    expect(commentsXml).not.toContain('张三')
    const coreXml = await zip.file('docProps/core.xml')!.async('string')
    expect(coreXml).toContain('<dc:creator></dc:creator>')
    expect(coreXml).toContain('<cp:lastModifiedBy></cp:lastModifiedBy>')
  })

  it('scrubs all known metadata after copy-through assembly without touching document data', async () => {
    const wml = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    const extraParts = [
      {
        path: 'word/comments.xml',
        xml:
          `${XML_DECL}<c:comments xmlns:c='${wml}'>` +
          "<c:comment c:id='comment-7' c:author='Comment Person' c:initials='CP' c:date='2024-03-01T02:03:04Z'>" +
          '<c:p><c:r><c:t>Comment Person remains as comment text</c:t></c:r></c:p>' +
          '</c:comment></c:comments>',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
      },
      {
        path: 'word/header1.xml',
        xml:
          `${XML_DECL}<h:hdr xmlns:h="${wml}">` +
          '<h:p><h:ins h:id="header-rev" h:author="Header Person" h:date="2024-03-02T02:03:04Z">' +
          '<h:r><h:t>Header content</h:t></h:r></h:ins></h:p></h:hdr>',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
      },
      {
        path: 'word/footer1.xml',
        xml:
          `${XML_DECL}<f:ftr xmlns:f='${wml}'>` +
          "<f:p><f:del f:id='footer-rev' f:author='Footer Person' f:date='2024-03-03T02:03:04Z'>" +
          '<f:r><f:delText>Footer content</f:delText></f:r></f:del></f:p></f:ftr>',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
      },
      {
        path: 'word/footnotes.xml',
        xml:
          `${XML_DECL}<n:footnotes xmlns:n="${wml}">` +
          '<n:footnote n:id="2"><n:p><n:ins n:id="foot-rev" n:author="Foot Person" n:date="2024-03-04T02:03:04Z">' +
          '<n:r><n:t>Footnote content</n:t></n:r></n:ins></n:p></n:footnote></n:footnotes>',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
      },
      {
        path: 'word/endnotes.xml',
        xml:
          `${XML_DECL}<n:endnotes xmlns:n='${wml}'>` +
          "<n:endnote n:id='3'><n:p><n:ins n:id='end-rev' n:author='End Person' n:date='2024-03-05T02:03:04Z'>" +
          '<n:r><n:t>Endnote content</n:t></n:r></n:ins></n:p></n:endnote></n:endnotes>',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
      },
      {
        path: 'word/glossary/document.xml',
        xml:
          `${XML_DECL}<g:document xmlns:g="${wml}"><g:body>` +
          '<g:p><g:ins g:id="glossary-rev" g:author="Glossary Person" g:date="2024-03-06T02:03:04Z">' +
          '<g:r><g:t>Glossary content</g:t></g:r></g:ins></g:p>' +
          '</g:body></g:document>',
        contentType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml',
      },
      {
        path: 'docProps/core.xml',
        xml:
          `${XML_DECL}<meta:coreProperties xmlns:meta='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' ` +
          "xmlns:d='http://purl.org/dc/elements/1.1/' xmlns:t='http://purl.org/dc/terms/'>" +
          "<d:creator role='writer'>Core Person</d:creator>" +
          '<meta:lastModifiedBy>Last Person</meta:lastModifiedBy>' +
          '<d:title>Preserved title</d:title>' +
          '<t:created>2024-01-01T00:00:00Z</t:created>' +
          '</meta:coreProperties>',
        contentType: 'application/vnd.openxmlformats-package.core-properties+xml',
      },
      {
        path: 'docProps/app.xml',
        xml:
          `${XML_DECL}<ep:Properties xmlns:ep='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'>` +
          '<ep:Manager>Manager Person</ep:Manager><ep:Company>Example Org</ep:Company>' +
          '<ep:Application>Preserved App</ep:Application></ep:Properties>',
        contentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
      },
      {
        path: 'word/people.xml',
        xml:
          `${XML_DECL}<p:people xmlns:p="http://schemas.microsoft.com/office/word/2012/wordml">` +
          '<p:person p:author="People Person" p:initials="PP" p:personId="person-keep">' +
          "<p:presenceInfo p:providerId='Example Org' p:userId='person@example.com'/>" +
          '</p:person></p:people>',
        contentType: 'application/vnd.ms-word.person+xml',
      },
      {
        path: 'customXml/item1.xml',
        xml:
          `${XML_DECL}<custom xmlns:w="${wml}" author="Custom Person">` +
          '<w:ins w:id="custom-rev" w:author="Custom Person"><value>Keep me</value></w:ins>' +
          '</custom>',
        contentType: 'application/xml',
      },
      {
        path: 'docProps/custom.xml',
        xml:
          `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">` +
          '<property name="Company"><value>Custom Org</value></property></Properties>',
        contentType: 'application/vnd.openxmlformats-officedocument.custom-properties+xml',
      },
    ]
    const customXml = extraParts.find((part) => part.path === 'customXml/item1.xml')!.xml
    const customProps = extraParts.find((part) => part.path === 'docProps/custom.xml')!.xml
    const binary = Buffer.from([0, 255, 17, 34, 51])
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:p xmlns:x='${wml}'><w:ins x:id='body-rev' x:author='Body Person' x:date='2024-02-01T02:03:04Z'>` +
          '<w:r><w:t>Body Person remains as content; author="Visible Person"; w:author=\'Visible Qualified\'</w:t></w:r></w:ins></w:p>',
        extraRels:
          '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
          '<Relationship Id="rId41" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
          '<Relationship Id="rId42" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>',
        sectPrExtra:
          '<w:headerReference w:type="default" r:id="rId40"/><w:footerReference w:type="default" r:id="rId41"/>',
        extraParts,
        binaryParts: [
          {
            path: 'word/media/privacy-proof.bin',
            base64: binary.toString('base64'),
            extension: 'bin',
            contentType: 'application/octet-stream',
          },
        ],
      }),
    )

    const saved = await saveDocx(parsed, originalBlocks(parsed), {
      removePersonalInfo: true,
      savedAt: '2026-01-01T00:00:00Z',
    })
    const zip = await JSZip.loadAsync(saved)
    const read = (path: string) => zip.file(path)!.async('string')

    const documentXml = await read('word/document.xml')
    expect(documentXml).toContain("x:author='Author'")
    expect(documentXml).toContain("x:id='body-rev'")
    expect(documentXml).toContain("x:date='2024-02-01T02:03:04Z'")
    expect(documentXml).toContain('Body Person remains as content')
    expect(documentXml).toContain('author="Visible Person"')
    expect(documentXml).toContain("w:author='Visible Qualified'")

    for (const [path, removed, kept, id, date] of [
      [
        'word/comments.xml',
        'Comment Person',
        'Comment Person remains as comment text',
        'comment-7',
        '2024-03-01T02:03:04Z',
      ],
      ['word/header1.xml', 'Header Person', 'Header content', 'header-rev', '2024-03-02T02:03:04Z'],
      ['word/footer1.xml', 'Footer Person', 'Footer content', 'footer-rev', '2024-03-03T02:03:04Z'],
      ['word/footnotes.xml', 'Foot Person', 'Footnote content', 'foot-rev', '2024-03-04T02:03:04Z'],
      ['word/endnotes.xml', 'End Person', 'Endnote content', 'end-rev', '2024-03-05T02:03:04Z'],
      [
        'word/glossary/document.xml',
        'Glossary Person',
        'Glossary content',
        'glossary-rev',
        '2024-03-06T02:03:04Z',
      ],
    ] as const) {
      const xml = await read(path)
      expect(xml).not.toContain(`${removed}'`)
      expect(xml).not.toContain(`${removed}"`)
      expect(xml).toContain(kept)
      expect(xml).toContain('author=')
      expect(xml).toContain(id)
      expect(xml).toContain(date)
    }
    const comments = await read('word/comments.xml')
    expect(comments).toContain("c:author='Author'")
    expect(comments).toContain("c:initials='A'")
    expect(comments).toContain("c:id='comment-7'")
    expect(comments).toContain("c:date='2024-03-01T02:03:04Z'")

    const core = await read('docProps/core.xml')
    expect(core).toContain("<d:creator role='writer'></d:creator>")
    expect(core).toContain('<meta:lastModifiedBy></meta:lastModifiedBy>')
    expect(core).toContain('<d:title>Preserved title</d:title>')
    expect(core).toContain('<t:created>2024-01-01T00:00:00Z</t:created>')

    const app = await read('docProps/app.xml')
    expect(app).toContain('<ep:Manager></ep:Manager>')
    expect(app).toContain('<ep:Company></ep:Company>')
    expect(app).toContain('<ep:Application>Preserved App</ep:Application>')

    const people = await read('word/people.xml')
    expect(people).toContain('<p:people')
    expect(people).not.toContain('<p:person ')
    expect(people).not.toContain('People Person')
    expect(people).not.toContain('person@example.com')
    expect(people).not.toContain('person-keep')

    expect(await read('customXml/item1.xml')).toBe(customXml)
    expect(await read('docProps/custom.xml')).toBe(customProps)
    expect(await zip.file('word/media/privacy-proof.bin')!.async('nodebuffer')).toEqual(binary)
  })

  it('a document that already carries the flag scrubs an otherwise unchanged save', async () => {
    const settingsPart = {
      path: 'word/settings.xml',
      xml:
        `${XML_DECL}<s:settings xmlns:s='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>` +
        "<s:removePersonalInformation s:val='true'/></s:settings>",
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
    }
    const settingsRel =
      '<Relationship Id="rId41" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: bodyWithRevision + '<w:p><w:r><w:t>tail</w:t></w:r></w:p>',
        extraRels: settingsRel,
        extraParts: [settingsPart],
      }),
    )
    expect(parsed.removePersonalInfo).toBe(true)
    const saved = await saveDocx(parsed, originalBlocks(parsed), {})
    const docXml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(docXml).toContain('w:author="Author"')
    expect(docXml).not.toContain('张三')
  })

  it('false removes the settings flag and leaves authors alone', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: bodyWithRevision }))
    const withFlag = await parseDocx(
      await saveDocx(parsed, originalBlocks(parsed), { removePersonalInfo: true }),
    )
    expect(withFlag.removePersonalInfo).toBe(true)
    const cleared = await saveDocx(withFlag, originalBlocks(withFlag), {
      removePersonalInfo: false,
    })
    const reparsed = await parseDocx(cleared)
    expect(reparsed.removePersonalInfo).toBe(false)
  })
})
