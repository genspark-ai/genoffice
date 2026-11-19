import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const BODY = '<w:p><w:r><w:t>Document with no styles part</w:t></w:r></w:p>'

describe('package without styles.xml', () => {
  it('takes Word built-in Normal defaults: Aptos 12pt, after 8pt, line 1.15', async () => {
    const zip = await JSZip.loadAsync(await buildDocx({ bodyXml: BODY }))
    zip.remove('word/styles.xml')
    const doc = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    expect(doc.docDefaults).toMatchObject({
      asciiFont: 'Aptos',
      sizeHalfPoints: 24,
      spaceAfterTwips: 160,
      lineRawTwips: 276,
      lineRule: 'auto',
    })
  })

  it('an unparseable styles part degrades to the same defaults', async () => {
    const zip = await JSZip.loadAsync(await buildDocx({ bodyXml: BODY }))
    zip.file('word/styles.xml', '<w:styles')
    const doc = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    expect(doc.docDefaults).toMatchObject({ asciiFont: 'Aptos', sizeHalfPoints: 24 })
  })

  it('a present styles part keeps its own docDefaults', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    expect(doc.docDefaults?.asciiFont).not.toBe('Aptos')
  })
})
