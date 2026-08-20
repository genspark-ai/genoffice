import { describe, expect, it } from 'vitest'
import {
  chunkText,
  cosineSimilarity,
  ollamaApiRoot,
  pickEmbeddingModel,
  searchChunks,
} from '../src/embed'

describe('ollamaApiRoot', () => {
  it('strips /v1 and trailing slashes', () => {
    expect(ollamaApiRoot('http://localhost:11434/v1')).toBe('http://localhost:11434')
    expect(ollamaApiRoot('http://localhost:11434/v1/')).toBe('http://localhost:11434')
    expect(ollamaApiRoot('http://localhost:11434//')).toBe('http://localhost:11434')
    expect(ollamaApiRoot()).toBe('http://localhost:11434')
  })
})

describe('chunkText', () => {
  it('returns nothing for empty input', () => {
    expect(chunkText('')).toEqual([])
    expect(chunkText('   ')).toEqual([])
  })

  it('keeps short text as one chunk', () => {
    expect(chunkText('hello world')).toEqual(['hello world'])
  })

  it('splits long text with overlap', () => {
    const text = 'a '.repeat(500)
    const chunks = chunkText(text, 200, 50)
    expect(chunks.length).toBeGreaterThan(1)
    expect(chunks.every((c) => c.length <= 200)).toBe(true)
    // every chunk (after the first) re-includes a tail of the previous one
    for (let i = 1; i < chunks.length; i++) {
      const tail = chunks[i - 1].slice(-50).trim()
      expect(chunks[i].includes(tail)).toBe(true)
    }
  })

  it('prefers paragraph boundaries', () => {
    const text = `${'x'.repeat(100)}\n\n${'y'.repeat(100)}`
    const chunks = chunkText(text, 120, 10)
    expect(chunks.length).toBe(2)
    expect(chunks[0].includes('x')).toBe(true)
    expect(chunks[1].includes('y')).toBe(true)
  })
})

describe('cosineSimilarity', () => {
  it('is 1 for identical vectors and 0 for orthogonal ones', () => {
    expect(cosineSimilarity([1, 2, 3], [1, 2, 3])).toBeCloseTo(1)
    expect(cosineSimilarity([1, 0], [0, 1])).toBeCloseTo(0)
  })

  it('handles empty/mismatched vectors', () => {
    expect(cosineSimilarity([], [])).toBe(0)
    expect(cosineSimilarity([1], [1, 2])).toBe(0)
  })
})

describe('pickEmbeddingModel', () => {
  it('prefers embed-capable names', () => {
    expect(pickEmbeddingModel(['qwen2.5:7b', 'nomic-embed-text:latest'])).toBe('nomic-embed-text:latest')
    expect(pickEmbeddingModel(['llama3.2:3b', 'bge-m3:latest'])).toBe('bge-m3:latest')
  })

  it('returns null when nothing looks embeddable', () => {
    expect(pickEmbeddingModel(['qwen2.5:7b', 'llama3.2:3b'])).toBeNull()
    expect(pickEmbeddingModel([])).toBeNull()
  })
})

describe('searchChunks', () => {
  const chunks = [
    { file: 'a.md', text: 'alpha', vector: [1, 0, 0] },
    { file: 'b.md', text: 'beta', vector: [0, 1, 0] },
    { file: 'c.md', text: 'gamma', vector: [0.8, 0.2, 0] },
  ]

  it('ranks by cosine and returns only positive matches', () => {
    const hits = searchChunks(chunks, [1, 0, 0], 2)
    expect(hits.map((h) => h.file)).toEqual(['a.md', 'c.md'])
  })

  it('respects k and drops unrelated chunks', () => {
    const hits = searchChunks(chunks, [0, 0, 1], 5)
    expect(hits).toEqual([])
  })
})
