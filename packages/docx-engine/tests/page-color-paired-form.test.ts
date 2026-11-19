import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { parseDocx, readPageColor, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const BODY = '<w:p><w:r><w:t>body</w:t></w:r></w:p>'
const VML_NAMESPACES =
  'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'

/** The paired spelling Word writes for a gradient page background: w:color is only a fallback. */
const PAIRED_BACKGROUND =
  '<w:background w:color="FFFFFF">' +
  '<v:background id="_x0000_s1025" o:bwmode="white" o:targetscreensize="1024,768">' +
  '<v:fill color2="fill darken(118)" rotate="t" angle="180" focus="100%" type="gradient"/>' +
  '</v:background>' +
  '</w:background>'

async function docXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/document.xml')!.async('string')
}

/** buildDocx only appends children to w:body, so inject a w:background as w:document's first child. */
async function buildWithBackground(backgroundXml: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(
    await buildDocx({ bodyXml: BODY, docRootExtraAttrs: VML_NAMESPACES }),
  )
  const xml = await zip.file('word/document.xml')!.async('string')
  zip.file(
    'word/document.xml',
    xml.replace(/(<w:document[^>]*>)/, (_match, open: string) => open + backgroundXml),
  )
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

async function saveWithPageColor(
  backgroundXml: string,
  pageColor: string | null,
): Promise<Uint8Array> {
  const parsed = await parseDocx(await buildWithBackground(backgroundXml))
  const blocks: SaveBlock[] = [{ kind: 'original', docxIndex: 0 }]
  return saveDocx(parsed, blocks, { pageColor })
}

const count = (xml: string, tag: string): number =>
  (xml.match(new RegExp(`<${tag}`, 'g')) ?? []).length

describe('page color set on a document that already has a w:background', () => {
  // CT_Document admits a single optional w:background. Producers write it either
  // self-closing or, when it carries a VML fill, as an element pair. Removing
  // only the self-closing form left the paired element in place, so saving
  // emitted two w:background children and Word reported the file as corrupt.
  it('replaces a paired w:background so exactly one remains', async () => {
    const xml = await docXmlOf(await saveWithPageColor(PAIRED_BACKGROUND, 'FFF2CC'))

    expect(count(xml, 'w:background')).toBe(1)
    expect(xml).not.toContain('</w:background>')
    expect(xml).toMatch(/<w:document[^>]*><w:background w:color="FFF2CC"\/>/)
  })

  it('removes a paired w:background when the page color is cleared', async () => {
    const xml = await docXmlOf(await saveWithPageColor(PAIRED_BACKGROUND, null))

    expect(count(xml, 'w:background')).toBe(0)
    expect(count(xml, 'v:background')).toBe(0)
    expect(xml).not.toContain('</w:background>')
  })

  it('still handles the self-closing spelling unchanged', async () => {
    const set = await docXmlOf(
      await saveWithPageColor('<w:background w:color="FFFFFF"/>', 'FFF2CC'),
    )
    expect(count(set, 'w:background')).toBe(1)
    expect(set).toMatch(/<w:document[^>]*><w:background w:color="FFF2CC"\/>/)

    const cleared = await docXmlOf(
      await saveWithPageColor('<w:background w:color="FFFFFF"/>', null),
    )
    expect(count(cleared, 'w:background')).toBe(0)
  })

  // Round-trip: the saved bytes must reparse, report the color through the
  // public reader, and stay valid for a second save (idempotent).
  it('round-trips the page color through parse and save', async () => {
    const saved = await saveWithPageColor(PAIRED_BACKGROUND, 'FFF2CC')
    const reparsed = await parseDocx(saved)

    expect(readPageColor(reparsed)).toBe('FFF2CC')

    const blocks: SaveBlock[] = [{ kind: 'original', docxIndex: 0 }]
    const resaved = await saveDocx(reparsed, blocks, { pageColor: 'FFF2CC' })
    const xml = await docXmlOf(resaved)

    expect(count(xml, 'w:background')).toBe(1)
    expect(readPageColor(await parseDocx(resaved))).toBe('FFF2CC')
  })
})
