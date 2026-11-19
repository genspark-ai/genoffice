import { describe, expect, it } from 'vitest'
import {
  LEGACY_DOC_LIMITS,
  LegacyDocError,
  openCompoundFile,
  openLegacyDoc,
  readLegacyDocText,
} from '../src/index'
import { fixtureBytes } from './helpers/fixtures'
import { buildCompound, buildSyntheticDoc, encodeClx } from './helpers/synthetic-doc'

/** run a reader call and return its LegacyDocError code (fails on success or another error type) */
function errorCode(run: () => unknown): string {
  try {
    run()
  } catch (error) {
    expect(error).toBeInstanceOf(LegacyDocError)
    return (error as LegacyDocError).code
  }
  throw new Error('expected the reader to throw a LegacyDocError')
}

const ONE_PIECE = [{ text: 'body\r' }]

describe('legacy .doc container errors', () => {
  it('rejects an empty file as not_cfb', () => {
    expect(errorCode(() => readLegacyDocText(new Uint8Array(0)))).toBe('not_cfb')
  })

  it('rejects a plain-text file as not_cfb', () => {
    const bytes = new TextEncoder().encode('This is not an OLE2 compound document at all.')
    expect(errorCode(() => readLegacyDocText(bytes))).toBe('not_cfb')
  })

  it('rejects a renamed zip (.docx) as not_cfb', () => {
    const bytes = Uint8Array.from([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0, 0, 0])
    expect(errorCode(() => readLegacyDocText(bytes))).toBe('not_cfb')
  })

  it('reports a truncated OLE2 header as truncated', () => {
    const bytes = fixtureBytes('text-basic').slice(0, 200)
    expect(errorCode(() => readLegacyDocText(bytes))).toBe('truncated')
  })

  it('reports a truncated CFB body as a LegacyDocError', () => {
    const bytes = fixtureBytes('text-basic').slice(0, 4000)
    expect(() => readLegacyDocText(bytes)).toThrow(LegacyDocError)
  })

  it('rejects a container without a WordDocument stream', () => {
    const bytes = buildCompound([['1Table', Uint8Array.from([0x02, 4, 0, 0, 0, 0, 0, 0, 0])]])
    expect(errorCode(() => readLegacyDocText(bytes))).toBe('missing_stream')
  })

  it('rejects a header whose table stream is missing', () => {
    const bytes = buildSyntheticDoc({ pieces: ONE_PIECE, omitTableStream: true })
    expect(errorCode(() => readLegacyDocText(bytes))).toBe('missing_stream')
  })

  it('caps the input size', () => {
    const bytes = fixtureBytes('text-basic')
    expect(
      errorCode(() =>
        openCompoundFile(bytes, {
          maxBytes: 1024,
          maxStreamBytes: LEGACY_DOC_LIMITS.maxStreamBytes,
        }),
      ),
    ).toBe('limits_exceeded')
  })

  it('caps single stream sizes', () => {
    const bytes = fixtureBytes('text-basic')
    expect(
      errorCode(() =>
        openCompoundFile(bytes, { maxBytes: LEGACY_DOC_LIMITS.maxBytes, maxStreamBytes: 32 }),
      ),
    ).toBe('limits_exceeded')
  })
})

describe('legacy .doc FIB errors', () => {
  it('rejects a bad FibBase.wIdent', () => {
    expect(errorCode(() => readLegacyDocText(buildSyntheticDoc({ wIdent: 0x1234 })))).toBe(
      'bad_fib',
    )
  })

  it('rejects a Word 6/95 document (nFib 0x0065) as unsupported_version', () => {
    expect(errorCode(() => readLegacyDocText(buildSyntheticDoc({ nFib: 0x0065 })))).toBe(
      'unsupported_version',
    )
  })

  it('rejects a newer-than-Word-2003 format as unsupported_version', () => {
    expect(errorCode(() => readLegacyDocText(buildSyntheticDoc({ nFib: 0x0113 })))).toBe(
      'unsupported_version',
    )
  })

  it('rejects an encrypted document', () => {
    expect(
      errorCode(() => readLegacyDocText(buildSyntheticDoc({ encrypted: true, pieces: ONE_PIECE }))),
    ).toBe('encrypted')
  })

  it('rejects a header without fcClx/lcbClx', () => {
    expect(errorCode(() => readLegacyDocText(buildSyntheticDoc({ lcbClx: 0 })))).toBe('bad_fib')
  })
})

describe('legacy .doc CLX / piece-table errors', () => {
  it('rejects an unknown CLX record type', () => {
    const bytes = buildSyntheticDoc({ clx: Uint8Array.from([0x03, 0, 0, 0, 0]) })
    expect(errorCode(() => readLegacyDocText(bytes))).toBe('bad_clx')
  })

  it('rejects a Pcdt payload that is not 12n+4 bytes', () => {
    const clx = new Uint8Array(5 + 13)
    clx[0] = 0x02
    new DataView(clx.buffer).setUint32(1, 13, true)
    const bytes = buildSyntheticDoc({ clx })
    expect(errorCode(() => readLegacyDocText(bytes))).toBe('bad_clx')
  })

  it('rejects non-monotonic piece boundaries', () => {
    const clx = encodeClx([0, 10, 5], [{ fc: 2048 }, { fc: 2048 }])
    const bytes = buildSyntheticDoc({ clx, ccpText: 5 })
    expect(errorCode(() => readLegacyDocText(bytes))).toBe('bad_piece_table')
  })

  it('rejects a piece whose text range lies outside the WordDocument stream', () => {
    const clx = encodeClx([0, 5], [{ fc: 1 << 20 }])
    const bytes = buildSyntheticDoc({ clx, ccpText: 5 })
    expect(errorCode(() => readLegacyDocText(bytes))).toBe('bad_piece_table')
  })

  it('rejects a CLX range that lies outside the table stream', () => {
    const clx = encodeClx([0], [])
    const bytes = buildSyntheticDoc({ clx, lcbClx: 999 })
    expect(errorCode(() => readLegacyDocText(bytes))).toBe('bad_clx')
  })

  it('rejects a ccpText the piece table does not cover', () => {
    const bytes = buildSyntheticDoc({ pieces: [{ text: 'abc' }], ccpText: 10 })
    expect(errorCode(() => readLegacyDocText(bytes))).toBe('bad_piece_table')
  })

  it('rejects a CLX with no Pcdt record (Prc only)', () => {
    const bytes = buildSyntheticDoc({ clx: Uint8Array.from([0x01, 2, 0, 0xaa, 0xbb]) })
    expect(errorCode(() => readLegacyDocText(bytes))).toBe('bad_clx')
  })

  it('exposes the error codes and an actionable message', () => {
    try {
      openLegacyDoc(buildSyntheticDoc({ encrypted: true, pieces: ONE_PIECE }))
    } catch (error) {
      expect(error).toBeInstanceOf(LegacyDocError)
      const failure = error as LegacyDocError
      expect(failure.code).toBe('encrypted')
      expect(failure.message).toContain('password')
      expect(failure.message).toContain('Word or LibreOffice')
    }
  })
})
