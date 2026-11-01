import * as CFB from 'cfb'
import { describe, expect, it } from 'vitest'
import {
  LegacyDocError,
  openLegacyDoc,
  parsePlcBte,
  readFormatting,
  readLegacyDocModel,
  readLegacyDocText,
} from '../src/index'
import { FIXTURE_NAMES, fixtureBytes, type FixtureName } from './helpers/fixtures'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodeChpxFkp,
  encodePapxFkp,
} from './helpers/synthetic-doc'

/**
 * Mutation validation: a truncated or randomly corrupted .doc must either be
 * read or fail with LegacyDocError. A raw RangeError/TypeError (or a hang)
 * escaping the reader is a bug, because the import UI can only act on the
 * documented codes.
 */

/** deterministic LCG so a failure is reproducible from the seed alone */
function seededRandom(seed: number): () => number {
  let state = seed >>> 0
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    return state / 0x100000000
  }
}

/** read one mutated file; the only allowed outcomes are a result or a LegacyDocError */
function attempt(bytes: Uint8Array): 'read' | 'rejected' {
  try {
    const result = readLegacyDocText(bytes)
    expect(typeof result.text).toBe('string')
    expect(Array.isArray(result.paragraphs)).toBe(true)
    return 'read'
  } catch (error) {
    expect(error).toBeInstanceOf(LegacyDocError)
    return 'rejected'
  }
}

/** the data streams of a fixture, without cfb's producer marker */
function streamsOf(bytes: Uint8Array): Array<[string, Uint8Array]> {
  const container = CFB.read(bytes, { type: 'buffer' })
  const streams: Array<[string, Uint8Array]> = []
  container.FullPaths.forEach((path, index) => {
    const entry = container.FileIndex[index]
    if (!entry || entry.type !== 2) return
    const name = path.startsWith('Root Entry/')
      ? path.slice('Root Entry/'.length)
      : path.replace(/^\/+/, '')
    if (!name || name.startsWith('\u0001')) return
    streams.push([name, new Uint8Array(entry.content as ArrayLike<number>)])
  })
  return streams
}

/** rebuild the container with one FIB byte flipped (positions 0-63 cover FibBase + FibRgW97) */
function flipFibByte(bytes: Uint8Array, position: number, xor: number): Uint8Array {
  const streams = streamsOf(bytes)
  const word = streams.find(([name]) => name === 'WordDocument')
  if (!word) throw new Error('fixture has no WordDocument stream')
  word[1][position] = (word[1][position] ?? 0) ^ xor
  const container = CFB.utils.cfb_new()
  for (const [name, content] of streams) CFB.utils.cfb_add(container, name, Array.from(content))
  return new Uint8Array(CFB.write(container, { type: 'buffer', fileType: 'cfb' }) as Uint8Array)
}

function mutationRun(name: FixtureName): { read: number; rejected: number } {
  const bytes = fixtureBytes(name)
  const random = seededRandom(0xd0c2_0000 ^ name.length)
  const outcomes = { read: 0, rejected: 0 }
  const count = (outcome: 'read' | 'rejected') => {
    outcomes[outcome]++
  }
  // 20 truncations at seeded offsets
  for (let index = 0; index < 20; index++) {
    count(attempt(bytes.slice(0, Math.floor(random() * bytes.length))))
  }
  // 20 single-byte flips in the CFB header
  for (let index = 0; index < 20; index++) {
    const mutated = bytes.slice()
    const position = Math.floor(random() * Math.min(512, mutated.length))
    mutated[position] = (mutated[position] ?? 0) ^ (1 + Math.floor(random() * 255))
    count(attempt(mutated))
  }
  // 20 single-byte flips in the FIB region of the WordDocument stream
  for (let index = 0; index < 20; index++) {
    const position = Math.floor(random() * 64)
    const xor = 1 + Math.floor(random() * 255)
    count(attempt(flipFibByte(bytes, position, xor)))
  }
  return outcomes
}

describe('mutation validation: every corrupted input ends as data or LegacyDocError', () => {
  for (const name of FIXTURE_NAMES) {
    it(`${name}: truncations and header/FIB byte flips never leak a raw error`, () => {
      const outcomes = mutationRun(name)
      expect(outcomes.read + outcomes.rejected).toBe(60)
      // the fixtures must still be recognizable for at least the FIB flips that
      // land in unused header bytes; both outcomes are otherwise fine
      expect(outcomes.rejected).toBeGreaterThan(0)
    })
  }

  it('keeps the real fixtures readable after the loop', () => {
    for (const name of FIXTURE_NAMES) {
      expect(readLegacyDocText(fixtureBytes(name)).paragraphs.length).toBeGreaterThan(1)
    }
  })
})

/** read one mutated file through the stage-2 model; a model or a LegacyDocError are the only outcomes */
function attemptModel(bytes: Uint8Array): 'read' | 'rejected' {
  try {
    const model = readLegacyDocModel(bytes)
    expect(Array.isArray(model.paragraphs)).toBe(true)
    for (const paragraph of model.paragraphs) expect(typeof paragraph.text).toBe('string')
    return 'read'
  } catch (error) {
    expect(error).toBeInstanceOf(LegacyDocError)
    return 'rejected'
  }
}

/** read one mutated FKP/PLCF state through the formatting layer */
function attemptFormatting(source: Parameters<typeof readFormatting>[0]): 'read' | 'rejected' {
  try {
    const paragraphs = readFormatting(source)
    expect(Array.isArray(paragraphs)).toBe(true)
    return 'read'
  } catch (error) {
    expect(error).toBeInstanceOf(LegacyDocError)
    return 'rejected'
  }
}

describe('mutation validation: stage-2 formatting model', () => {
  it('formatting.doc: truncations, CFB header flips and FIB flips never leak a raw error', () => {
    const bytes = fixtureBytes('formatting')
    const random = seededRandom(0x5f0c_0001)
    const outcomes = { read: 0, rejected: 0 }
    const count = (outcome: 'read' | 'rejected') => {
      outcomes[outcome]++
    }
    for (let index = 0; index < 15; index++)
      count(attemptModel(bytes.slice(0, Math.floor(random() * bytes.length))))
    for (let index = 0; index < 15; index++) {
      const mutated = bytes.slice()
      const position = Math.floor(random() * Math.min(512, mutated.length))
      mutated[position] = (mutated[position] ?? 0) ^ (1 + Math.floor(random() * 255))
      count(attemptModel(mutated))
    }
    for (let index = 0; index < 15; index++)
      count(
        attemptModel(flipFibByte(bytes, Math.floor(random() * 64), 1 + Math.floor(random() * 255))),
      )
    expect(outcomes.read + outcomes.rejected).toBe(45)
    expect(outcomes.rejected).toBeGreaterThan(0)
  })

  it('formatting.doc: FKP and PLCF byte flips end as a formatting model or a LegacyDocError', () => {
    const doc = openLegacyDoc(fixtureBytes('formatting'))
    const chpxPages = parsePlcBte(doc.table, doc.fib.fcPlcfbteChpx, doc.fib.lcbPlcfbteChpx).pages
    const papxPages = parsePlcBte(doc.table, doc.fib.fcPlcfbtePapx, doc.fib.lcbPlcfbtePapx).pages
    const random = seededRandom(0x5f0c_0002)
    const outcomes = { read: 0, rejected: 0 }
    for (let index = 0; index < 30; index++) {
      const wordDocument = doc.wordDocument.slice()
      const table = doc.table.slice()
      const target = index % 3
      if (target === 0 || target === 1) {
        const pages = target === 0 ? chpxPages : papxPages
        const page = pages[Math.floor(random() * pages.length)] ?? 5
        const position = page * 512 + Math.floor(random() * 512)
        wordDocument[position] = (wordDocument[position] ?? 0) ^ (1 + Math.floor(random() * 255))
      } else {
        const usingChpx = random() < 0.5
        const fc = usingChpx ? doc.fib.fcPlcfbteChpx : doc.fib.fcPlcfbtePapx
        const lcb = usingChpx ? doc.fib.lcbPlcfbteChpx : doc.fib.lcbPlcfbtePapx
        const position = fc + Math.floor(random() * lcb)
        table[position] = (table[position] ?? 0) ^ (1 + Math.floor(random() * 255))
      }
      outcomes[
        attemptFormatting({
          fib: doc.fib,
          pieces: doc.pieces,
          wordDocument,
          table,
          rawText: doc.rawText,
          stream: doc.stream,
        })
      ]++
    }
    expect(outcomes.read + outcomes.rejected).toBe(30)
    expect(outcomes.rejected).toBeGreaterThan(0)
  })

  it('formatting.doc: truncating the WordDocument stream below an FKP page is a LegacyDocError', () => {
    const doc = openLegacyDoc(fixtureBytes('formatting'))
    const chpxPages = parsePlcBte(doc.table, doc.fib.fcPlcfbteChpx, doc.fib.lcbPlcfbteChpx).pages
    const page = chpxPages[0]!
    const outcome = attemptFormatting({
      fib: doc.fib,
      pieces: doc.pieces,
      wordDocument: doc.wordDocument.slice(0, page * 512),
      table: doc.table,
      rawText: doc.rawText,
      stream: doc.stream,
    })
    expect(outcome).toBe('rejected')
  })

  it('formatting.doc: truncating the table stream before the PLCFs is a LegacyDocError', () => {
    const doc = openLegacyDoc(fixtureBytes('formatting'))
    const outcome = attemptFormatting({
      fib: doc.fib,
      pieces: doc.pieces,
      wordDocument: doc.wordDocument,
      table: doc.table.slice(0, doc.fib.fcPlcfbteChpx),
      rawText: doc.rawText,
      stream: doc.stream,
    })
    expect(outcome).toBe('rejected')
  })

  it('synthetic FKP documents: FKP and PLCF byte flips end as a model or a LegacyDocError', () => {
    const text = 'bold and plain\rsecond paragraph\r'
    const first = text.indexOf('\r')
    const second = text.indexOf('\r', first + 1)
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      chpxFkpPages: [
        encodeChpxFkp([
          {
            fcStart: SYNTHETIC_TEXT_OFFSET,
            fcEnd: SYNTHETIC_TEXT_OFFSET + (first + 1) * 2,
            grpprl: Uint8Array.from([0x35, 0x08, 1]),
          },
          {
            fcStart: SYNTHETIC_TEXT_OFFSET + (first + 1) * 2,
            fcEnd: SYNTHETIC_TEXT_OFFSET + (second + 1) * 2,
          },
        ]),
      ],
      papxFkpPages: [
        encodePapxFkp([
          {
            fcStart: SYNTHETIC_TEXT_OFFSET,
            fcEnd: SYNTHETIC_TEXT_OFFSET + (first + 1) * 2,
            grpprl: Uint8Array.from([0x06, 0x24, 1]),
          },
        ]),
      ],
    })
    const doc = openLegacyDoc(bytes)
    const chpxPages = parsePlcBte(doc.table, doc.fib.fcPlcfbteChpx, doc.fib.lcbPlcfbteChpx).pages
    const papxPages = parsePlcBte(doc.table, doc.fib.fcPlcfbtePapx, doc.fib.lcbPlcfbtePapx).pages
    const random = seededRandom(0x5f0c_0003)
    const outcomes = { read: 0, rejected: 0 }
    for (let index = 0; index < 24; index++) {
      const wordDocument = doc.wordDocument.slice()
      const table = doc.table.slice()
      switch (index % 4) {
        case 0: {
          const page = chpxPages[0]!
          const position = page * 512 + Math.floor(random() * 512)
          wordDocument[position] = (wordDocument[position] ?? 0) ^ (1 + Math.floor(random() * 255))
          break
        }
        case 1: {
          const page = papxPages[0]!
          const position = page * 512 + Math.floor(random() * 512)
          wordDocument[position] = (wordDocument[position] ?? 0) ^ (1 + Math.floor(random() * 255))
          break
        }
        case 2: {
          const position = doc.fib.fcPlcfbteChpx + Math.floor(random() * doc.fib.lcbPlcfbteChpx)
          table[position] = (table[position] ?? 0) ^ (1 + Math.floor(random() * 255))
          break
        }
        default: {
          const position = doc.fib.fcPlcfbtePapx + Math.floor(random() * doc.fib.lcbPlcfbtePapx)
          table[position] = (table[position] ?? 0) ^ (1 + Math.floor(random() * 255))
        }
      }
      outcomes[
        attemptFormatting({
          fib: doc.fib,
          pieces: doc.pieces,
          wordDocument,
          table,
          rawText: doc.rawText,
          stream: doc.stream,
        })
      ]++
    }
    expect(outcomes.read + outcomes.rejected).toBe(24)
    expect(outcomes.rejected).toBeGreaterThan(0)
  })

  it('keeps every fixture readable through the formatting model after the loops', () => {
    for (const name of FIXTURE_NAMES) {
      const model = readLegacyDocModel(fixtureBytes(name))
      expect(model.paragraphs.length).toBeGreaterThan(0)
      expect(model.warnings).toEqual([])
    }
  })
})
