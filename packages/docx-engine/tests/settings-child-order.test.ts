/**
 * CT_Settings child order in the saved word/settings.xml. The type is an
 * xsd:sequence, so a child written at the wrong position makes Word offer to
 * repair the part. Every apply* used to insert right after the w:settings open
 * tag, so the element written last took the first slot and a save touching
 * several settings came out in the reverse of the sequence.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { parseDocx, saveDocx, type SaveOptions } from '../src/index'
import { splitXmlChildren } from '../src/generate'
import { buildDocx } from './helpers/build-docx'

const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const BODY = '<w:p><w:r><w:t>body</w:t></w:r></w:p>'

const settingsPart = (inner: string, prefixed = true) =>
  prefixed
    ? `${XML_DECL}<w:settings ${W_NS}>${inner}</w:settings>`
    : `${XML_DECL}<settings xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${inner}</settings>`

async function saveWith(
  settingsInner: string,
  options: SaveOptions,
  prefixed = true,
): Promise<string> {
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml: BODY,
      extraParts: [
        {
          path: 'word/settings.xml',
          xml: settingsPart(settingsInner, prefixed),
          contentType: SETTINGS_TYPE,
        },
      ],
    }),
  )
  const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], options)
  return (await JSZip.loadAsync(saved)).file('word/settings.xml')!.async('string')
}

/** local names of the settings children, in emitted order */
async function childNames(options: SaveOptions, settingsInner = ''): Promise<string[]> {
  const xml = await saveWith(settingsInner, options)
  return childNamesOf(xml)
}

function childNamesOf(xml: string): string[] {
  const close =
    xml.indexOf('</w:settings>') === -1 ? xml.indexOf('</settings>') : xml.indexOf('</w:settings>')
  const open =
    xml.indexOf('<w:settings') === -1 ? xml.indexOf('<settings') : xml.indexOf('<w:settings')
  return splitXmlChildren(xml.slice(xml.indexOf('>', open) + 1, close)).map((c) => c.name)
}

const ALL_OPTIONS: SaveOptions = {
  pageColor: 'FF0000',
  protection: { edit: 'readOnly', enforced: true },
  writeProtection: { recommended: true },
  removePersonalInfo: true,
  evenAndOddHeaders: true,
  mirrorMargins: true,
}

describe('CT_Settings child order', () => {
  it('emits the six settings in schema sequence order', async () => {
    expect(await childNames(ALL_OPTIONS)).toEqual([
      'w:writeProtection',
      'w:removePersonalInformation',
      'w:displayBackgroundShape',
      'w:mirrorMargins',
      'w:documentProtection',
      'w:evenAndOddHeaders',
    ])
  })

  it('lands a new child among the children the part already had', async () => {
    // zoom (2nd) < mirrorMargins (15th) < documentProtection (36th) <
    // defaultTabStop (39th) < evenAndOddHeaders (48th) < compat (87th), so
    // neither "insert at the root" nor "append at the end" is good enough here.
    const existing =
      '<w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/>' +
      '<w:compat><w:compatSetting w:name="compatibilityMode" w:val="15"/></w:compat>'
    const xml = await saveWith(existing, {
      protection: { edit: 'readOnly', enforced: true },
      evenAndOddHeaders: true,
      mirrorMargins: true,
    })
    expect(childNamesOf(xml)).toEqual([
      'w:zoom',
      'w:mirrorMargins',
      'w:documentProtection',
      'w:defaultTabStop',
      'w:evenAndOddHeaders',
      'w:compat',
    ])
    // the untouched children keep their attributes
    expect(xml).toContain('<w:defaultTabStop w:val="420"/>')
    expect(xml).toContain('<w:compatSetting w:name="compatibilityMode" w:val="15"/>')
  })

  it('replaces an element already in the part at its own position', async () => {
    const names = await childNames(
      { protection: { edit: 'trackedChanges', enforced: true } },
      '<w:zoom w:percent="100"/><w:documentProtection w:edit="readOnly" w:enforcement="0"/>',
    )
    expect(names).toEqual(['w:zoom', 'w:documentProtection'])
    const xml = await saveWith('<w:documentProtection w:edit="readOnly" w:enforcement="0"/>', {
      protection: { edit: 'trackedChanges', enforced: true },
    })
    expect(xml).toContain('w:edit="trackedChanges"')
    expect(xml).not.toContain('w:edit="readOnly"')
  })

  it('orders a settings part that uses a default namespace instead of the w: prefix', async () => {
    // removePersonalInformation is the one writer that follows the part's own
    // prefix; defaultTabStop (39th) must still end up after it (4th)
    const xml = await saveWith('<defaultTabStop w:val="420"/>', { removePersonalInfo: true }, false)
    expect(xml).toContain('<removePersonalInformation/>')
    expect(xml.indexOf('<removePersonalInformation')).toBeLessThan(xml.indexOf('<defaultTabStop'))
  })

  it('leaves the part in schema order when settings are switched off', async () => {
    const existing = '<w:mirrorMargins/><w:evenAndOddHeaders/>'
    const names = await childNames({ mirrorMargins: false, evenAndOddHeaders: false }, existing)
    expect(names).toEqual([])
  })
})
