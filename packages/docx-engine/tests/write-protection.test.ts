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

  it('a document that already carries the flag keeps scrubbing on later saves', async () => {
    const settingsPart = {
      path: 'word/settings.xml',
      xml:
        `${XML_DECL}<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        '<w:removePersonalInformation/></w:settings>',
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
    // an ordinary content edit (drop the tail paragraph), no privacy option passed
    const blocks = originalBlocks(parsed).slice(0, -1)
    const saved = await saveDocx(parsed, blocks, {})
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
