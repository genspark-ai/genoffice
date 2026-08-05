import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/** overwrite the declared uncompressed size of every central-directory record */
function patchCentralSizes(bytes: Uint8Array, size: number): Uint8Array {
  const out = bytes.slice()
  const dv = new DataView(out.buffer, out.byteOffset, out.byteLength)
  for (let i = 0; i + 28 <= out.length; i++) {
    if (out[i] === 0x50 && out[i + 1] === 0x4b && out[i + 2] === 0x01 && out[i + 3] === 0x02) {
      dv.setUint32(i + 24, size, true)
    }
  }
  return out
}

const PLAIN_PARA = '<w:p><w:r><w:t>hello</w:t></w:r></w:p>'

describe('zip bomb protection', () => {
  it('rejects a part whose declared uncompressed size exceeds the per-part limit', async () => {
    const bytes = await buildDocx({ bodyXml: PLAIN_PARA })
    const bomb = patchCentralSizes(bytes, 600 * 1024 * 1024)
    await expect(parseDocx(bomb)).rejects.toThrow(/docx rejected.*uncompressed/)
  })

  it('rejects an archive whose total declared uncompressed size exceeds the limit', async () => {
    const bytes = await buildDocx({ bodyXml: PLAIN_PARA })
    const zip = await JSZip.loadAsync(bytes)
    expect(Object.values(zip.files).filter((f) => !f.dir).length).toBeGreaterThanOrEqual(4)
    const bomb = patchCentralSizes(bytes, 450 * 1024 * 1024)
    await expect(parseDocx(bomb)).rejects.toThrow(/total uncompressed/)
  })

  it('rejects an archive with too many parts', async () => {
    const zip = await JSZip.loadAsync(await buildDocx({ bodyXml: PLAIN_PARA }))
    for (let i = 0; i < 10010; i++) zip.file(`junk/${i}.bin`, 'x')
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    await expect(parseDocx(bytes)).rejects.toThrow(/parts exceeds/)
  })

  it('still opens a normal document', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PLAIN_PARA }))
    expect(doc.blocks[0].runs?.[0]?.text).toBe('hello')
  })
})

describe('malformed XML degrades instead of failing the document', () => {
  const DEEP = 3000

  it('turns an unparseable paragraph into a byte-preserving passthrough block', async () => {
    const nested =
      '<w:p>' +
      '<w:smartTag>'.repeat(DEEP) +
      '<w:r><w:t>deep</w:t></w:r>' +
      '</w:smartTag>'.repeat(DEEP) +
      '</w:p>'
    const source = await buildDocx({ bodyXml: nested + PLAIN_PARA })
    const doc = await parseDocx(source)
    expect(doc.blocks[0].type).toBe('passthrough')
    expect(doc.blocks[0].previewText).toContain('deep')
    expect(doc.blocks[1].runs?.[0]?.text).toBe('hello')
    const saved = await saveDocx(doc, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'original', docxIndex: 1 },
    ])
    expect(saved).toEqual(source)
  })

  it('drops the textbox display model when its XML is unparseable, keeping the block', async () => {
    const drawing =
      '<w:p><w:r><w:drawing><wp:anchor><w:txbxContent>' +
      '<a:g>'.repeat(DEEP) +
      '</a:g>'.repeat(DEEP) +
      '<w:p><w:r><w:t>boxed</w:t></w:r></w:p>' +
      '</w:txbxContent></wp:anchor></w:drawing></w:r></w:p>'
    const source = await buildDocx({ bodyXml: drawing })
    const doc = await parseDocx(source)
    expect(doc.blocks[0].type).toBe('passthrough')
    expect(doc.blocks[0].textboxes).toBeUndefined()
    const saved = await saveDocx(doc, [{ kind: 'original', docxIndex: 0 }])
    expect(saved).toEqual(source)
  })

  it('degrades an unparseable styles.xml to an empty style table', async () => {
    const zip = await JSZip.loadAsync(await buildDocx({ bodyXml: PLAIN_PARA }))
    const deepStyles =
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:x>'.repeat(DEEP) +
      '</w:x>'.repeat(DEEP) +
      '</w:styles>'
    zip.file('word/styles.xml', deepStyles)
    const doc = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    expect(doc.styles.size).toBe(0)
    expect(doc.blocks[0].runs?.[0]?.text).toBe('hello')
  })
})
