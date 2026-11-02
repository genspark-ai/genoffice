import { describe, expect, it } from 'vitest'
import { LegacyDocError } from '../src/index'
import {
  SPRM_LONG_PARAGRAPH,
  SPRM_LONG_TABLE,
  iterateSprms,
  readSprmI16,
  readSprmU16,
  readSprmU32,
  readSprmU8,
  sprmOperandSize,
} from '../src/sprm'
import { parsePlcBte } from '../src/plcf'
import { readChpxFkp, readPapxFkp } from '../src/fkp'
import { cpOfFc, fcOfCp } from '../src/formatting'
import type { PieceTable } from '../src/clx'
import { encodeChpxFkp, encodePapxFkp, encodePlcBte } from './helpers/synthetic-doc'

function bytes(...values: number[]): Uint8Array {
  return Uint8Array.from(values)
}

function sprm(opcode: number, operand: number[]): Uint8Array {
  return Uint8Array.from([opcode & 0xff, (opcode >> 8) & 0xff, ...operand])
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

function errorCode(run: () => unknown): string {
  try {
    run()
  } catch (error) {
    expect(error).toBeInstanceOf(LegacyDocError)
    return (error as LegacyDocError).code
  }
  throw new Error('expected a LegacyDocError')
}

describe('sprm operand sizing (MS-DOC 2.6.1)', () => {
  it('decodes every spra class', () => {
    const grpprl = concat(
      sprm(0x0001, [0xaa]), // spra 0: 1 byte
      sprm(0x2001, [0xaa]), // spra 1: 1 byte
      sprm(0x4001, [0xaa, 0xbb]), // spra 2: 2 bytes
      sprm(0x6001, [1, 2, 3, 4]), // spra 3: 4 bytes
      sprm(0x8001, [0xaa, 0xbb]), // spra 4: 2 bytes
      sprm(0xa001, [0xaa, 0xbb]), // spra 5: 2 bytes
      sprm(0xc001, [3, 1, 2, 3]), // spra 6: length byte 3 + 3 bytes
      sprm(0xe001, [1, 2, 3]), // spra 7: 3 bytes
    )
    expect(iterateSprms(grpprl).map((entry) => entry.operand.length)).toEqual([
      1, 1, 2, 4, 2, 2, 3, 3,
    ])
    expect(sprmOperandSize(grpprl, 0)).toBe(1)
  })

  it('uses a u16 length for sprmTDefTable / sprmPChgTabs (declared = payload + 1)', () => {
    for (const opcode of [SPRM_LONG_TABLE, SPRM_LONG_PARAGRAPH]) {
      // Word writes 3 for a 2-byte operand: the u16 length counts one of its own bytes
      const grpprl = concat(sprm(opcode, [3, 0, 0xaa, 0xbb]), sprm(0x0835, [1]))
      const sprms = iterateSprms(grpprl)
      expect(sprms).toHaveLength(2)
      expect(sprms[0]!.opcode).toBe(opcode)
      expect([...sprms[0]!.operand]).toEqual([0xaa, 0xbb])
      expect(sprms[1]!.opcode).toBe(0x0835)
    }
  })

  it('allows an empty u16-length operand and rejects a declared zero length', () => {
    const empty = iterateSprms(concat(sprm(SPRM_LONG_TABLE, [1, 0]), sprm(0x0835, [1])))
    expect(empty.map((entry) => entry.opcode)).toEqual([SPRM_LONG_TABLE, 0x0835])
    expect(empty[0]!.operand).toHaveLength(0)
    expect(errorCode(() => iterateSprms(sprm(SPRM_LONG_TABLE, [0, 0])))).toBe('bad_sprm')
  })

  it('excludes the length byte from a variable-length operand', () => {
    const sprms = iterateSprms(concat(sprm(0xc001, [2, 0xaa, 0xbb]), sprm(0x0835, [1])))
    expect(sprms).toHaveLength(2)
    expect([...sprms[0]!.operand]).toEqual([0xaa, 0xbb])
    expect(sprms[1]!.opcode).toBe(0x0835)
  })

  it('accepts a zero-length variable operand and keeps walking', () => {
    const sprms = iterateSprms(concat(sprm(0xc001, [0]), sprm(0x0835, [1])))
    expect(sprms.map((entry) => entry.opcode)).toEqual([0xc001, 0x0835])
    expect(sprms[0]!.operand).toHaveLength(0)
  })

  it('rejects a truncated opcode', () => {
    expect(errorCode(() => iterateSprms(bytes(0x35)))).toBe('bad_sprm')
  })

  it('rejects a truncated operand', () => {
    expect(errorCode(() => iterateSprms(sprm(0x4a43, [0x1c])))).toBe('bad_sprm')
  })

  it('rejects a variable-length sprm with no length byte', () => {
    expect(errorCode(() => iterateSprms(bytes(0xc1, 0xc0)))).toBe('bad_sprm')
  })

  it('reads u8/u16/i16/u32 operands with bounds checks', () => {
    const operand = bytes(0x34, 0x12, 0xff, 0xfe)
    expect(readSprmU8(operand)).toBe(0x34)
    expect(readSprmU16(operand)).toBe(0x1234)
    expect(readSprmI16(bytes(0xfe, 0xff))).toBe(-2)
    expect(readSprmU32(bytes(0x78, 0x56, 0x34, 0x12))).toBe(0x12345678)
    expect(readSprmU8(bytes())).toBeUndefined()
    expect(readSprmU16(bytes(0x34))).toBeUndefined()
    expect(readSprmU32(bytes(1, 2, 3))).toBeUndefined()
  })
})

describe('PLCF BTE tables', () => {
  it('parses FCs and page numbers, masking the page field', () => {
    const fcs = [2048, 4096, 8192]
    const pages = [5 | 0x400000, 7]
    const table = encodePlcBte(fcs, pages)
    const plc = parsePlcBte(table, 0, table.length)
    expect(plc.fcs).toEqual(fcs)
    expect(plc.pages).toEqual([5, 7])
  })

  it('accepts a zero-length table', () => {
    expect(parsePlcBte(new Uint8Array(0), 0, 0)).toEqual({ fcs: [], pages: [] })
  })

  it('rejects a length that is not 8n+4', () => {
    expect(errorCode(() => parsePlcBte(new Uint8Array(32), 0, 5))).toBe('bad_fkp')
  })

  it('rejects a range outside the table stream', () => {
    expect(errorCode(() => parsePlcBte(new Uint8Array(8), 4, 12))).toBe('bad_fkp')
  })
})

describe('CHPX FKP pages', () => {
  const grpprl = sprm(0x0835, [1])

  function pageAt(wordDocument: Uint8Array, page: Uint8Array): void {
    wordDocument.set(page, 512)
  }

  it('reads runs, FC boundaries and grpprls', () => {
    const page = encodeChpxFkp([
      { fcStart: 2048, fcEnd: 2052, grpprl },
      { fcStart: 2052, fcEnd: 2056 },
    ])
    const word = new Uint8Array(1024)
    pageAt(word, page)
    const runs = readChpxFkp(word, 1)
    expect(runs).toHaveLength(2)
    expect(runs[0]).toMatchObject({ fcStart: 2048, fcEnd: 2052 })
    expect([...runs[0]!.grpprl]).toEqual([...grpprl])
    expect([...runs[1]!.grpprl]).toEqual([])
  })

  it('reads an empty page (crun 0)', () => {
    const word = new Uint8Array(1024)
    pageAt(word, encodeChpxFkp([]))
    expect(readChpxFkp(word, 1)).toEqual([])
  })

  it('rejects an FKP page outside the WordDocument stream', () => {
    expect(errorCode(() => readChpxFkp(new Uint8Array(512), 1))).toBe('bad_fkp')
  })

  it('rejects an impossible run count', () => {
    const word = new Uint8Array(512)
    word[511] = 200
    expect(errorCode(() => readChpxFkp(word, 0))).toBe('bad_fkp')
  })

  it('rejects a CHPX offset outside the run data area', () => {
    const word = new Uint8Array(512)
    word[511] = 1
    // one run: rgfc at 0..8, rgb at 8; data starts at 9 (rounded up to 10)
    word[8] = 1 // offset 2 -> at the rgfc area
    expect(errorCode(() => readChpxFkp(word, 0))).toBe('bad_fkp')
  })

  it('rejects a CHPX record that overruns the page', () => {
    const word = new Uint8Array(512)
    word[511] = 1
    word[8] = 255 // offset 510
    word[510] = 5 // cb 5 needs bytes up to 516
    expect(errorCode(() => readChpxFkp(word, 0))).toBe('bad_fkp')
  })
})

describe('PAPX FKP pages', () => {
  const longSprms = sprm(0xa413, [0xf0, 0x00]) // 4 bytes: istd + 4 = 6 -> long form
  const shortSprms = sprm(0x2406, [1]) // 3 bytes: istd + 3 = 5 -> short form

  it('reads paragraph start FCs, istd and sprms in both length forms', () => {
    const page = encodePapxFkp([
      { fcStart: 2048, fcEnd: 2056, istd: 2, grpprl: longSprms },
      { fcStart: 2056, fcEnd: 2064, istd: 1, grpprl: shortSprms },
      { fcStart: 2064, fcEnd: 2072, defaults: true },
    ])
    const word = new Uint8Array(1024)
    word.set(page, 512)
    const entries = readPapxFkp(word, 1)
    expect(entries.map((entry) => entry.fcStart)).toEqual([2048, 2056, 2064])
    expect(entries[0]).toMatchObject({ istd: 2, defaults: false })
    expect([...entries[0]!.grpprl]).toEqual([...longSprms])
    expect(entries[1]).toMatchObject({ istd: 1, defaults: false })
    expect([...entries[1]!.grpprl]).toEqual([...shortSprms])
    expect(entries[2]).toMatchObject({ istd: 0, defaults: true })
    expect([...entries[2]!.grpprl]).toEqual([])

    // the first entry is stored in the long form (cb 0 + word count), the second short
    const rgbxStart = 4 * 4
    const longOffset = page[rgbxStart]! * 2
    const shortOffset = page[rgbxStart + 13]! * 2
    expect(page[longOffset]).toBe(0)
    expect(page[longOffset + 1]).toBe(3)
    expect(page[shortOffset]).toBe(3)
  })

  it('rejects a run count no 512-byte page can hold', () => {
    const word = new Uint8Array(512)
    word[511] = 100
    expect(errorCode(() => readPapxFkp(word, 0))).toBe('bad_fkp')
  })

  it('rejects a bOffset pointing outside the run data area', () => {
    const word = new Uint8Array(512)
    word[511] = 1
    word[8] = 255 // BxPap.bOffset -> byte 510
    word[510] = 0 // long form: second length byte would land past the page
    word[511] = 1
    expect(errorCode(() => readPapxFkp(word, 0))).toBe('bad_fkp')
  })

  it('rejects a PAPX shorter than its istd', () => {
    const word = new Uint8Array(512)
    word[511] = 1
    word.set([2048, 0, 0, 0, 2056, 0, 0, 0], 0) // FCs
    word[8] = 11 // bOffset*2 = 22
    word[22] = 1 // short form: grpprl length 1 (needs >= 2 for istd)
    expect(errorCode(() => readPapxFkp(word, 0))).toBe('bad_fkp')
  })
})

describe('FC <-> CP mapping', () => {
  const pieceTable: PieceTable = {
    cps: [0, 3, 5],
    pieces: [
      { cpStart: 0, cpEnd: 3, fc: 100, compressed: true },
      { cpStart: 3, cpEnd: 5, fc: 200, compressed: false },
    ],
  }

  it('maps FCs inside compressed and UTF-16 pieces to CPs', () => {
    expect(cpOfFc(pieceTable, 100)).toEqual({ cp: 0 })
    expect(cpOfFc(pieceTable, 101)).toEqual({ cp: 1 })
    expect(cpOfFc(pieceTable, 200)).toEqual({ cp: 3 })
    expect(cpOfFc(pieceTable, 202)).toEqual({ cp: 4 })
    expect(cpOfFc(pieceTable, 204)).toEqual({ cp: 5 })
  })

  it('flags an unaligned UTF-16 byte offset and an FC outside every piece', () => {
    expect(cpOfFc(pieceTable, 201)).toBe('unaligned')
    expect(cpOfFc(pieceTable, 0)).toBe('outside')
    expect(cpOfFc(pieceTable, 300)).toBe('outside')
  })

  it('maps CPs back to FCs', () => {
    expect(fcOfCp(pieceTable, 0)).toBe(100)
    expect(fcOfCp(pieceTable, 1)).toBe(101)
    expect(fcOfCp(pieceTable, 3)).toBe(200)
    expect(fcOfCp(pieceTable, 4)).toBe(202)
    expect(fcOfCp(pieceTable, 5)).toBe(204)
    expect(fcOfCp(pieceTable, 6)).toBeUndefined()
  })
})
