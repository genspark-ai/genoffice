import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { parseDocx, saveDocx, type SaveOptions } from '../src/index'
import { splitXmlChildren } from '../src/generate'
import { buildDocx } from './helpers/build-docx'

const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const BODY = '<w:p><w:r><w:t>body</w:t></w:r></w:p>'
const ORIGINAL = [{ kind: 'original' as const, docxIndex: 0 }]

const docWithSettings = (inner: string) =>
  buildDocx({
    bodyXml: BODY,
    extraParts: [
      {
        path: 'word/settings.xml',
        xml: `${XML_DECL}<w:settings ${W_NS}>${inner}</w:settings>`,
        contentType: SETTINGS_TYPE,
      },
    ],
  })

async function settingsOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/settings.xml')!.async('string')
}

const childNamesOf = (xml: string): string[] =>
  splitXmlChildren(/<w:settings[^>]*>([\s\S]*)<\/w:settings>/.exec(xml)![1]!).map((c) => c.name)

async function resave(inner: string, options: SaveOptions): Promise<string> {
  const parsed = await parseDocx(await docWithSettings(inner))
  return settingsOf(await saveDocx(parsed, ORIGINAL, options))
}

describe('settings.xml w:trackRevisions', () => {
  it('parses the flag (w:val="0" counts as off)', async () => {
    expect((await parseDocx(await docWithSettings('<w:trackRevisions/>'))).trackRevisions).toBe(
      true,
    )
    expect(
      (await parseDocx(await docWithSettings('<w:trackRevisions w:val="0"/>'))).trackRevisions,
    ).toBeUndefined()
    expect((await parseDocx(await docWithSettings(''))).trackRevisions).toBeUndefined()
  })

  it('removes the flag when tracking was turned off, keeping the move/format sub-flags', async () => {
    const xml = await resave(
      '<w:zoom w:percent="100"/><w:trackRevisions/><w:doNotTrackMoves/><w:doNotTrackFormatting/><w:defaultTabStop w:val="720"/>',
      { trackRevisions: false },
    )
    expect(childNamesOf(xml)).toEqual([
      'w:zoom',
      'w:doNotTrackMoves',
      'w:doNotTrackFormatting',
      'w:defaultTabStop',
    ])
  })

  it('adds the flag at its CT_Settings position when tracking was turned on', async () => {
    const xml = await resave(
      '<w:zoom w:percent="100"/><w:revisionView w:markup="1"/><w:documentProtection w:edit="readOnly" w:enforcement="0"/><w:defaultTabStop w:val="720"/>',
      { trackRevisions: true },
    )
    expect(childNamesOf(xml)).toEqual([
      'w:zoom',
      'w:revisionView',
      'w:trackRevisions',
      'w:documentProtection',
      'w:defaultTabStop',
    ])
    expect((xml.match(/<w:trackRevisions/g) ?? []).length).toBe(1)
  })

  it('creates settings.xml for the flag when the package has none', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const xml = await settingsOf(await saveDocx(parsed, ORIGINAL, { trackRevisions: true }))
    expect(xml).toContain('<w:trackRevisions/>')
  })

  it('leaves the flag alone when the save does not carry the option', async () => {
    expect(await resave('<w:trackRevisions/>', {})).toContain('<w:trackRevisions/>')
    expect(await resave('<w:trackRevisions/>', { mirrorMargins: true })).toContain(
      '<w:trackRevisions/>',
    )
    expect(await resave('', { mirrorMargins: true })).not.toContain('trackRevisions')
  })
})
