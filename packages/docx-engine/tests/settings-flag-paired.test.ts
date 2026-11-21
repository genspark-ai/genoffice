import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { parseDocx, saveDocx } from '../src/index'
import type { DocProtection } from '../src/types'
import { buildDocx } from './helpers/build-docx'

const BODY = '<w:p><w:r><w:t>body</w:t></w:r></w:p>'
const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const settingsPart = (inner: string) => `${XML_DECL}<w:settings ${W_NS}>${inner}</w:settings>`

async function settingsOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/settings.xml')!.async('string')
}

const count = (xml: string, tag: string): number =>
  (xml.match(new RegExp(`<${tag}`, 'g')) ?? []).length

async function saveWithFlags(settingsInner: string, options: Record<string, boolean | number>) {
  const source = await buildDocx({
    bodyXml: BODY,
    extraParts: [
      { path: 'word/settings.xml', xml: settingsPart(settingsInner), contentType: SETTINGS_TYPE },
    ],
  })
  const parsed = await parseDocx(source)
  const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], options)
  return settingsOf(saved)
}

async function saveWithProtection(settingsInner: string, protection: DocProtection | null) {
  const source = await buildDocx({
    bodyXml: BODY,
    extraParts: [
      { path: 'word/settings.xml', xml: settingsPart(settingsInner), contentType: SETTINGS_TYPE },
    ],
  })
  const parsed = await parseDocx(source)
  const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], { protection })
  return settingsOf(saved)
}

describe('settings flags written as an element pair', () => {
  // A paired <w:mirrorMargins></w:mirrorMargins> is legal XML, and producers
  // do emit it. The removal matched only the self-closing form, so the paired
  // element survived and switching the flag ON then appended a second one:
  // the saved part held both spellings of a zero-or-one element, which is
  // schema-invalid, and switching it OFF did nothing at all.
  it('leaves exactly one w:mirrorMargins when it is switched on', async () => {
    const xml = await saveWithFlags('<w:mirrorMargins></w:mirrorMargins>', { mirrorMargins: true })
    expect(count(xml, 'w:mirrorMargins')).toBe(1)
    expect(xml).not.toContain('</w:mirrorMargins>')
  })

  it('removes a paired w:mirrorMargins when it is switched off', async () => {
    const xml = await saveWithFlags('<w:mirrorMargins></w:mirrorMargins>', { mirrorMargins: false })
    expect(count(xml, 'w:mirrorMargins')).toBe(0)
    expect(xml).not.toContain('</w:mirrorMargins>')
  })

  it('leaves exactly one w:evenAndOddHeaders when it is switched on', async () => {
    const xml = await saveWithFlags('<w:evenAndOddHeaders></w:evenAndOddHeaders>', {
      evenAndOddHeaders: true,
    })
    expect(count(xml, 'w:evenAndOddHeaders')).toBe(1)
    expect(xml).not.toContain('</w:evenAndOddHeaders>')
  })

  it('removes a paired w:evenAndOddHeaders when it is switched off', async () => {
    const xml = await saveWithFlags('<w:evenAndOddHeaders></w:evenAndOddHeaders>', {
      evenAndOddHeaders: false,
    })
    expect(count(xml, 'w:evenAndOddHeaders')).toBe(0)
    expect(xml).not.toContain('</w:evenAndOddHeaders>')
  })

  it('still handles the self-closing spelling unchanged', async () => {
    expect(
      count(await saveWithFlags('<w:mirrorMargins/>', { mirrorMargins: true }), 'w:mirrorMargins'),
    ).toBe(1)
    expect(
      count(await saveWithFlags('<w:mirrorMargins/>', { mirrorMargins: false }), 'w:mirrorMargins'),
    ).toBe(0)
    expect(
      count(
        await saveWithFlags('<w:evenAndOddHeaders/>', { evenAndOddHeaders: true }),
        'w:evenAndOddHeaders',
      ),
    ).toBe(1)
    expect(
      count(
        await saveWithFlags('<w:evenAndOddHeaders/>', { evenAndOddHeaders: false }),
        'w:evenAndOddHeaders',
      ),
    ).toBe(0)
  })

  // w:documentProtection is empty-content and zero-or-one, so the paired
  // spelling is legal and producers do emit it. Removing only the self-closing
  // form left the paired element in place, so unprotecting a document did
  // nothing and setting protection appended a second, schema-invalid element.
  it('removes a paired w:documentProtection when protection is cleared', async () => {
    const xml = await saveWithProtection(
      '<w:documentProtection w:edit="readOnly" w:enforcement="1"></w:documentProtection>',
      null,
    )
    expect(count(xml, 'w:documentProtection')).toBe(0)
    expect(xml).not.toContain('</w:documentProtection>')
  })

  it('leaves exactly one w:documentProtection when it is set', async () => {
    const xml = await saveWithProtection(
      '<w:documentProtection w:edit="readOnly" w:enforcement="1"></w:documentProtection>',
      { edit: 'comments', enforced: true },
    )
    expect(count(xml, 'w:documentProtection')).toBe(1)
    expect(xml).not.toContain('</w:documentProtection>')
  })

  it('writes w:autoHyphenation once when switched on, and takes it out when switched off', async () => {
    const on = await saveWithFlags('<w:defaultTabStop w:val="720"/>', { autoHyphenation: true })
    expect(count(on, 'w:autoHyphenation')).toBe(1)
    // CT_Settings order: the flag follows the default tab stop
    expect(on.indexOf('<w:defaultTabStop')).toBeLessThan(on.indexOf('<w:autoHyphenation'))
    const reparsed = await saveWithFlags('<w:autoHyphenation/>', { autoHyphenation: true })
    expect(count(reparsed, 'w:autoHyphenation')).toBe(1)
    const off = await saveWithFlags('<w:autoHyphenation/>', { autoHyphenation: false })
    expect(count(off, 'w:autoHyphenation')).toBe(0)
  })
})

describe('the gutter position', () => {
  it('writes w:gutterAtTop once in schema order, reads it back and takes it out', async () => {
    const on = await saveWithFlags('<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/>', {
      gutterAtTop: true,
    })
    expect(count(on, 'w:gutterAtTop')).toBe(1)
    expect(on.indexOf('<w:gutterAtTop')).toBeLessThan(on.indexOf('<w:defaultTabStop'))
    const again = await saveWithFlags('<w:gutterAtTop></w:gutterAtTop>', { gutterAtTop: true })
    expect(count(again, 'w:gutterAtTop')).toBe(1)
    const off = await saveWithFlags('<w:gutterAtTop/>', { gutterAtTop: false })
    expect(count(off, 'w:gutterAtTop')).toBe(0)
  })
})

describe('the default tab stop', () => {
  it('replaces the one the file has, where it stands', async () => {
    const xml = await saveWithFlags(
      '<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>',
      { defaultTabStopTwips: 360 },
    )
    expect(count(xml, 'w:defaultTabStop')).toBe(1)
    expect(xml).toContain(
      '<w:zoom w:percent="100"/><w:defaultTabStop w:val="360"/><w:characterSpacingControl',
    )
  })

  it('is added in schema order to a part that has none, and read back', async () => {
    const source = await buildDocx({
      bodyXml: BODY,
      extraParts: [
        {
          path: 'word/settings.xml',
          xml: settingsPart('<w:zoom w:percent="100"/><w:compat/>'),
          contentType: SETTINGS_TYPE,
        },
      ],
    })
    const saved = await saveDocx(await parseDocx(source), [{ kind: 'original', docxIndex: 0 }], {
      defaultTabStopTwips: 1080,
    })
    expect(await settingsOf(saved)).toContain(
      '<w:zoom w:percent="100"/><w:defaultTabStop w:val="1080"/><w:compat/>',
    )
    expect((await parseDocx(saved)).defaultTabStopTwips).toBe(1080)
  })

  it('writes w:characterSpacingControl in place, replacing the value the file had', async () => {
    const on = await saveWithFlags('<w:defaultTabStop w:val="720"/><w:compat/>', {
      compressPunctuation: true,
    })
    expect(on).toContain('<w:characterSpacingControl w:val="compressPunctuation"/>')
    expect(on.indexOf('<w:defaultTabStop')).toBeLessThan(on.indexOf('<w:characterSpacingControl'))
    expect(on.indexOf('<w:characterSpacingControl')).toBeLessThan(on.indexOf('<w:compat'))
    const off = await saveWithFlags(
      '<w:characterSpacingControl w:val="compressPunctuationAndJapaneseKana"/>',
      { compressPunctuation: false },
    )
    expect(count(off, 'w:characterSpacingControl')).toBe(1)
    expect(off).toContain('<w:characterSpacingControl w:val="doNotCompress"/>')
    const reopened = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        extraParts: [{ path: 'word/settings.xml', xml: on, contentType: SETTINGS_TYPE }],
      }),
    )
    expect(reopened.compressPunctuation).toBe(true)
  })

  it('writes and removes w:strictFirstAndLastChars, and reads it back', async () => {
    const on = await saveWithFlags('<w:defaultTabStop w:val="720"/><w:compat/>', {
      strictFirstAndLastChars: true,
    })
    expect(count(on, 'w:strictFirstAndLastChars')).toBe(1)
    expect(on.indexOf('<w:strictFirstAndLastChars')).toBeLessThan(on.indexOf('<w:compat'))
    const reopened = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        extraParts: [{ path: 'word/settings.xml', xml: on, contentType: SETTINGS_TYPE }],
      }),
    )
    expect(reopened.strictFirstAndLastChars).toBe(true)
    const off = await saveWithFlags('<w:strictFirstAndLastChars/>', {
      strictFirstAndLastChars: false,
    })
    expect(count(off, 'w:strictFirstAndLastChars')).toBe(0)
  })
})
