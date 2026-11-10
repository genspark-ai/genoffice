import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

/** Rename the main part to word/trial.xml the way LO's tdf104713 sample does. */
async function withRenamedMainPart(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const move = async (from: string, to: string) => {
    const data = await zip.file(from)!.async('uint8array')
    zip.remove(from)
    zip.file(to, data)
  }
  await move('word/document.xml', 'word/trial.xml')
  await move('word/_rels/document.xml.rels', 'word/_rels/trial.xml.rels')
  for (const meta of ['_rels/.rels', '[Content_Types].xml']) {
    const xml = await zip.file(meta)!.async('string')
    zip.file(meta, xml.replaceAll('word/document.xml', 'word/trial.xml'))
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

describe('nonstandard main document part', () => {
  it('parses when the officeDocument relationship targets word/trial.xml', async () => {
    const bytes = await withRenamedMainPart(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>trial part text</w:t></w:r></w:p>' }),
    )
    const doc = await parseDocx(bytes)
    const allText = JSON.stringify(doc.blocks)
    expect(allText).toContain('trial part text')
  })

  it('save writes the edited XML into the renamed part', async () => {
    const bytes = await withRenamedMainPart(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>before edit</w:t></w:r></w:p>' }),
    )
    const doc = await parseDocx(bytes)
    const finalBlocks: SaveBlock[] = [
      { kind: 'generated', block: { type: 'paragraph', runs: [{ text: 'after edit' }] } },
    ]
    const saved = await saveDocx(doc, finalBlocks)
    const zip = await JSZip.loadAsync(saved)
    expect(zip.file('word/document.xml')).toBeNull()
    const xml = await zip.file('word/trial.xml')!.async('string')
    expect(xml).toContain('after edit')
    expect(xml).not.toContain('before edit')
  })

  it('rejects an ODT masquerading as .docx with a clear error', async () => {
    const zip = new JSZip()
    zip.file('mimetype', 'application/vnd.oasis.opendocument.text')
    zip.file('content.xml', `${XML_DECL}<office:document-content/>`)
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    await expect(parseDocx(bytes)).rejects.toThrow(/OpenDocument/)
  })
})
