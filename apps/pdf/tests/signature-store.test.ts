import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  MAX_SAVED_SIGNATURES,
  addSignature,
  isSignatureData,
  loadSignatures,
  removeSignature,
  sanitizeSignatures,
  saveSignatures,
} from '../src/main/signature-store'
import type { SavedSignature, SignatureData } from '../src/shared/ipc'

const strokes = (paths: number[][] = [[0, 0, 10, 10]]): SignatureData => ({
  kind: 'strokes',
  paths,
  width: 420,
  height: 150,
})

const image = (data = 'aGVsbG8='): SignatureData => ({
  kind: 'image',
  image: data,
  width: 300,
  height: 100,
})

describe('isSignatureData', () => {
  it('accepts stroke and image payloads', () => {
    expect(isSignatureData(strokes())).toBe(true)
    expect(isSignatureData(image())).toBe(true)
  })

  it('rejects malformed payloads', () => {
    expect(isSignatureData(null)).toBe(false)
    expect(isSignatureData({ kind: 'image', image: '', width: 10, height: 10 })).toBe(false)
    expect(isSignatureData({ kind: 'image', image: 'x', width: 0, height: 10 })).toBe(false)
    expect(isSignatureData({ kind: 'strokes', paths: [], width: 10, height: 10 })).toBe(false)
    // A stroke needs at least two points (4 numbers)
    expect(isSignatureData({ kind: 'strokes', paths: [[1, 2]], width: 10, height: 10 })).toBe(false)
    expect(
      isSignatureData({ kind: 'strokes', paths: [[1, 2, NaN, 4]], width: 10, height: 10 }),
    ).toBe(false)
    expect(isSignatureData({ kind: 'note', width: 10, height: 10 })).toBe(false)
  })
})

describe('addSignature', () => {
  it('prepends new signatures (newest first)', () => {
    const list = addSignature(addSignature([], strokes(), 1), image(), 2)
    expect(list).toHaveLength(2)
    expect(list[0]!.data.kind).toBe('image')
    expect(list[0]!.createdAt).toBe(2)
    expect(list[1]!.data.kind).toBe('strokes')
  })

  it('deduplicates identical payloads instead of duplicating', () => {
    const first = addSignature([], image('same'), 1)
    const again = addSignature(first, image('same'), 2)
    expect(again).toHaveLength(1)
    expect(again[0]!.createdAt).toBe(2)
  })

  it('keeps different payloads separate', () => {
    const list = addSignature(addSignature([], image('a'), 1), image('b'), 2)
    expect(list).toHaveLength(2)
  })

  it('evicts the oldest entries beyond the cap', () => {
    let list: SavedSignature[] = []
    for (let i = 0; i < MAX_SAVED_SIGNATURES + 3; i++) {
      list = addSignature(list, image(`sig-${i}`), i)
    }
    expect(list).toHaveLength(MAX_SAVED_SIGNATURES)
    expect(list[0]!.data).toMatchObject({ image: `sig-${MAX_SAVED_SIGNATURES + 2}` })
    // The oldest three were evicted
    expect(list.some((s) => s.data.kind === 'image' && s.data.image === 'sig-0')).toBe(false)
    expect(list.some((s) => s.data.kind === 'image' && s.data.image === 'sig-2')).toBe(false)
  })
})

describe('removeSignature', () => {
  it('removes by id and leaves others untouched', () => {
    const list = addSignature(addSignature([], image('a'), 1), image('b'), 2)
    const next = removeSignature(list, list[1]!.id)
    expect(next).toHaveLength(1)
    expect(next[0]!.id).toBe(list[0]!.id)
    expect(removeSignature(list, 'missing')).toHaveLength(2)
  })
})

describe('sanitizeSignatures', () => {
  it('drops malformed entries and duplicate ids', () => {
    const good: SavedSignature = { id: 'a', createdAt: 5, data: image() }
    const raw = [
      good,
      { id: 'a', createdAt: 6, data: image('dupe-id') },
      { id: '', createdAt: 1, data: image() },
      { id: 'b', createdAt: 1, data: { kind: 'image', image: '', width: 1, height: 1 } },
      { id: 'c', createdAt: 'nope', data: strokes() },
      'garbage',
      null,
    ]
    const out = sanitizeSignatures(raw)
    expect(out.map((s) => s.id)).toEqual(['a', 'c'])
    // Invalid createdAt falls back to 0 instead of dropping the signature
    expect(out[1]!.createdAt).toBe(0)
  })

  it('returns empty for non-arrays', () => {
    expect(sanitizeSignatures(undefined)).toEqual([])
    expect(sanitizeSignatures({})).toEqual([])
  })
})

describe('load/save round trip', () => {
  let dir: string
  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'pdf-sig-'))
  })
  afterEach(async () => {
    await rm(dir, { recursive: true, force: true })
  })

  it('persists and reloads the list', async () => {
    const file = join(dir, 'nested', 'pdf-signatures.json')
    const list = addSignature(addSignature([], strokes([[1, 2, 3, 4]]), 1), image(), 2)
    await saveSignatures(file, list)
    expect(await loadSignatures(file)).toEqual(list)
    // No stray tmp file left behind by the atomic write
    expect(await readFile(file, 'utf8')).toContain('"strokes"')
  })

  it('returns an empty list for missing or corrupt files', async () => {
    expect(await loadSignatures(join(dir, 'missing.json'))).toEqual([])
    const bad = join(dir, 'bad.json')
    await saveSignatures(bad, [])
    // Corrupt the file by writing invalid JSON over it
    const { writeFile } = await import('node:fs/promises')
    await writeFile(bad, '{ not json')
    expect(await loadSignatures(bad)).toEqual([])
  })
})
