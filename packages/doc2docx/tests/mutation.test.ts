import * as CFB from 'cfb'
import { describe, expect, it } from 'vitest'
import { LegacyDocError, readLegacyDocText } from '../src/index'
import { FIXTURE_NAMES, fixtureBytes, type FixtureName } from './helpers/fixtures'

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
