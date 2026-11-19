import { describe, expect, it } from 'vitest'
import { pastedBase64Image } from '../src/renderer/ai/base64-paste'

const PNG_SIG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]
const JPG_SIG = [0xff, 0xd8, 0xff, 0xe0]

function b64(bytes: number[], pad = 0): string {
  return Buffer.from([...bytes, ...new Array<number>(pad).fill(0)]).toString('base64')
}

describe('pastedBase64Image', () => {
  it('detects a pasted data: URL and types it from the signature', () => {
    const r = pastedBase64Image(`data:image/png;base64,${b64(PNG_SIG, 64)}`)
    expect(r?.ext).toBe('png')
    expect([...r!.bytes.slice(0, 4)]).toEqual(PNG_SIG.slice(0, 4))
  })

  it('maps a declared jpeg to jpg and survives whitespace line wraps', () => {
    const wrapped = b64(JPG_SIG, 64).replace(/(.{40})/g, '$1\n')
    const r = pastedBase64Image(`  data:image/jpeg;base64,${wrapped}  `)
    expect(r?.ext).toBe('jpg')
  })

  it('detects a bare base64 dump only when long and signature-typed', () => {
    const big = b64(PNG_SIG, 6000)
    expect(pastedBase64Image(big)?.ext).toBe('png')
    // short blobs and non-image bytes stay ordinary text
    expect(pastedBase64Image(b64(PNG_SIG, 32))).toBeNull()
    expect(pastedBase64Image(b64([1, 2, 3, 4], 6000))).toBeNull()
  })

  it('leaves ordinary prose and invalid base64 alone', () => {
    expect(pastedBase64Image('please add my logo at the top of the page')).toBeNull()
    expect(pastedBase64Image(`data:image/png;base64,%%%not-base64%%%`)).toBeNull()
    expect(pastedBase64Image(`Here it is: data:image/png;base64,${b64(PNG_SIG, 64)}`)).toBeNull()
  })
})
