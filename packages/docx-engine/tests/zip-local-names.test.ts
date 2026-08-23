import { crc32 } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const decoyDocumentXml = (text: string) =>
  XML_DECL +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  `<w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`

function findEocd(bytes: Uint8Array): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  for (let i = bytes.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === 0x06054b50) return i
  }
  throw new Error('no EOCD')
}

/**
 * Give `entry`'s central-directory record a crc-valid Info-ZIP Unicode Path
 * (0x7075) extra field claiming the name `claimed`, and clear its UTF-8 flag
 * so the field takes effect — the layout of the POI unicode-path sample.
 */
function injectUnicodePath(bytes: Uint8Array, entry: string, claimed: string): Uint8Array {
  const eocd = findEocd(bytes)
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const count = view.getUint16(eocd + 10, true)
  const decoder = new TextDecoder()
  const encoder = new TextEncoder()
  let p = view.getUint32(eocd + 16, true)
  for (let i = 0; i < count; i++) {
    expect(view.getUint32(p, true)).toBe(0x02014b50)
    const nameLen = view.getUint16(p + 28, true)
    const extraLen = view.getUint16(p + 30, true)
    const commentLen = view.getUint16(p + 32, true)
    const recordEnd = p + 46 + nameLen + extraLen + commentLen
    if (decoder.decode(bytes.subarray(p + 46, p + 46 + nameLen)) === entry) {
      const claimedBytes = encoder.encode(claimed)
      const field = new Uint8Array(4 + 5 + claimedBytes.length)
      const fv = new DataView(field.buffer)
      fv.setUint16(0, 0x7075, true)
      fv.setUint16(2, 5 + claimedBytes.length, true)
      field[4] = 1
      fv.setUint32(5, crc32(encoder.encode(entry)), true)
      field.set(claimedBytes, 9)

      const out = new Uint8Array(bytes.length + field.length)
      out.set(bytes.subarray(0, recordEnd))
      out.set(field, recordEnd)
      out.set(bytes.subarray(recordEnd), recordEnd + field.length)
      const ov = new DataView(out.buffer)
      ov.setUint16(p + 8, view.getUint16(p + 8, true) & ~0x800, true) // clear UTF-8 flag
      ov.setUint16(p + 30, extraLen + field.length, true)
      const newEocd = eocd + field.length
      ov.setUint32(newEocd + 12, view.getUint32(eocd + 12, true) + field.length, true) // cd size
      return out
    }
    p = recordEnd
  }
  throw new Error(`entry not found: ${entry}`)
}

describe('zip unicode-path extra field shadowing (POI unicode-path corpus)', () => {
  it('resolves word/document.xml by header name, like Word, not by a 0x7075 rename', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>RIGHT</w:t></w:r></w:p>',
      extraParts: [
        {
          path: 'decoy-content.xml',
          xml: decoyDocumentXml('WRONG'),
          contentType: 'application/xml',
        },
      ],
    })
    // both entries claim each other's name, like the POI sample: an
    // upath-honoring reader swaps them and reads the decoy as the document
    const hostile = injectUnicodePath(
      injectUnicodePath(bytes, 'decoy-content.xml', 'word/document.xml'),
      'word/document.xml',
      'decoy-content.xml',
    )

    const doc = await parseDocx(hostile)
    expect(doc.blocks[0].runs).toEqual([{ text: 'RIGHT' }])
  })

  it('leaves plain zips untouched', async () => {
    const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>plain</w:t></w:r></w:p>' })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].runs).toEqual([{ text: 'plain' }])
  })
})
