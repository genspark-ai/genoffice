import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { parseDocx, saveDocx } from '../src/index'
import type { SaveOptions } from '../src/patch'
import { buildDocx } from './helpers/build-docx'

const BODY = '<w:p><w:r><w:t>body</w:t></w:r></w:p>'
const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const source = (inner: string) =>
  buildDocx({
    bodyXml: BODY,
    extraParts: [
      {
        path: 'word/settings.xml',
        xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W_NS}>${inner}</w:settings>`,
        contentType: SETTINGS_TYPE,
      },
    ],
  })

async function save(inner: string, hyphenation: SaveOptions['hyphenation']) {
  const parsed = await parseDocx(await source(inner))
  const bytes = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], { hyphenation })
  const zip = await JSZip.loadAsync(bytes)
  return {
    xml: await zip.file('word/settings.xml')!.async('string'),
    reparsed: await parseDocx(bytes),
  }
}

describe('hyphenation settings', () => {
  it('reads the caps flag, the consecutive limit and the zone', async () => {
    const parsed = await parseDocx(
      await source(
        '<w:autoHyphenation/><w:consecutiveHyphenLimit w:val="2"/><w:hyphenationZone w:val="425"/><w:doNotHyphenateCaps/>',
      ),
    )
    expect(parsed.autoHyphenation).toBe(true)
    expect(parsed.doNotHyphenateCaps).toBe(true)
    expect(parsed.consecutiveHyphenLimit).toBe(2)
    expect(parsed.hyphenationZoneTwips).toBe(425)
  })

  it('leaves them undefined in a file that does not carry them', async () => {
    const parsed = await parseDocx(await source('<w:autoHyphenation/>'))
    expect(parsed.doNotHyphenateCaps).toBeUndefined()
    expect(parsed.consecutiveHyphenLimit).toBeUndefined()
    expect(parsed.hyphenationZoneTwips).toBeUndefined()
  })

  it('writes the three elements in schema order', async () => {
    const { xml, reparsed } = await save('<w:defaultTabStop w:val="720"/><w:showEnvelope/>', {
      caps: false,
      consecutiveLimit: 3,
      zoneTwips: 500,
    })
    expect(xml).toMatch(
      /<w:defaultTabStop w:val="720"\/><w:consecutiveHyphenLimit w:val="3"\/><w:hyphenationZone w:val="500"\/><w:doNotHyphenateCaps\/><w:showEnvelope\/>/,
    )
    expect(reparsed.doNotHyphenateCaps).toBe(true)
    expect(reparsed.consecutiveHyphenLimit).toBe(3)
    expect(reparsed.hyphenationZoneTwips).toBe(500)
  })

  it('removes the caps flag and the limit when they return to the defaults', async () => {
    const { xml } = await save(
      '<w:consecutiveHyphenLimit w:val="2"></w:consecutiveHyphenLimit><w:hyphenationZone w:val="425"/><w:doNotHyphenateCaps/>',
      { caps: true, consecutiveLimit: 0 },
    )
    expect(xml).not.toContain('consecutiveHyphenLimit')
    expect(xml).not.toContain('doNotHyphenateCaps')
    expect(xml).toContain('<w:hyphenationZone w:val="425"/>')
  })

  it('replaces an existing value instead of adding a second element', async () => {
    const { xml } = await save('<w:hyphenationZone w:val="425"/>', { zoneTwips: 200 })
    expect(xml.match(/<w:hyphenationZone/g)).toHaveLength(1)
    expect(xml).toContain('<w:hyphenationZone w:val="200"/>')
  })

  it('touches nothing when no field is given', async () => {
    const inner = '<w:hyphenationZone w:val="425"/><w:doNotHyphenateCaps/>'
    const { xml } = await save(inner, undefined)
    expect(xml).toContain(inner)
  })
})
