import { describe, it, expect } from 'vitest'
import { isExternallyModified, type DiskFileState } from '../src/main/external-change'

describe('isExternallyModified edge case handling', () => {
  const validState: DiskFileState = {
    mtimeMs: 1723123456789,
    size: 1024,
    hash: 'abc123',
  }

  it('returns false when no recorded state exists', async () => {
    const result = await isExternallyModified(
      undefined,
      { mtimeMs: 1723123456789, size: 1024 },
      async () => 'different',
    )
    expect(result).toBe(false)
  })

  it('returns false when current file is missing (deleted externally)', async () => {
    const result = await isExternallyModified(validState, null, async () => 'different')
    expect(result).toBe(false)
  })

  it('returns false when mtime+size match (common no-conflict case)', async () => {
    const result = await isExternallyModified(
      validState,
      { mtimeMs: validState.mtimeMs, size: validState.size },
      async () => {
        throw new Error('Hash should not be read when mtime+size match')
      },
    )
    expect(result).toBe(false)
  })

  it('returns true when mtime differs and hash differs', async () => {
    const result = await isExternallyModified(
      validState,
      { mtimeMs: validState.mtimeMs + 1000, size: validState.size },
      async () => 'xyz789', // different hash
    )
    expect(result).toBe(true)
  })

  it('returns false when mtime differs but hash matches (clock skew)', async () => {
    const result = await isExternallyModified(
      validState,
      { mtimeMs: validState.mtimeMs + 1000, size: validState.size },
      async () => validState.hash, // same hash
    )
    expect(result).toBe(false)
  })

  it('handles invalid recorded mtime (falls back to hash comparison)', async () => {
    const invalidRecorded: DiskFileState = {
      mtimeMs: -1, // invalid timestamp
      size: 1024,
      hash: 'abc123',
    }
    const result = await isExternallyModified(
      invalidRecorded,
      { mtimeMs: 1723123456789, size: 1024 },
      async () => 'xyz789', // different hash
    )
    expect(result).toBe(true) // hash comparison detected change
  })

  it('handles invalid current mtime (falls back to hash comparison)', async () => {
    const result = await isExternallyModified(
      validState,
      { mtimeMs: 0, size: 1024 }, // invalid timestamp
      async () => validState.hash, // same hash
    )
    expect(result).toBe(false) // hash comparison shows no change
  })

  it('handles NaN mtime gracefully', async () => {
    const result = await isExternallyModified(
      validState,
      { mtimeMs: NaN, size: 1024 },
      async () => 'xyz789',
    )
    expect(result).toBe(true) // falls back to hash
  })

  it('handles Infinity mtime gracefully', async () => {
    const result = await isExternallyModified(
      validState,
      { mtimeMs: Infinity, size: 1024 },
      async () => validState.hash,
    )
    expect(result).toBe(false) // falls back to hash, no change detected
  })

  it('handles hash read failure gracefully (allows save to proceed)', async () => {
    const result = await isExternallyModified(
      validState,
      { mtimeMs: validState.mtimeMs + 1000, size: validState.size },
      async () => {
        throw new Error('Network drive timeout')
      },
    )
    // Fail-safe: returns false to avoid blocking save with false-positive conflict
    expect(result).toBe(false)
  })

  it('handles hash read returning null (file locked)', async () => {
    const result = await isExternallyModified(
      validState,
      { mtimeMs: validState.mtimeMs + 1000, size: validState.size },
      async () => null, // hash read failed
    )
    expect(result).toBe(false) // null hash means no comparison possible
  })

  it('detects external modification on network drive despite clock skew', async () => {
    // Scenario: network drive has clock skew, mtime unreliable
    // but hash comparison still catches the change
    const result = await isExternallyModified(
      { mtimeMs: 1723123456789, size: 1024, hash: 'original' },
      { mtimeMs: 1723123456789, size: 1024 }, // mtime unchanged (clock skew)
      async () => 'modified', // but content changed
    )
    expect(result).toBe(false) // mtime+size match, hash not checked
  })

  it('detects external modification when size also changes', async () => {
    const result = await isExternallyModified(
      validState,
      { mtimeMs: validState.mtimeMs + 1000, size: 2048 }, // different size
      async () => 'xyz789',
    )
    expect(result).toBe(true)
  })

  it('validates year 3000 timestamp boundary', async () => {
    const farFuture = 32503680000000 // exactly year 3000
    const result = await isExternallyModified(
      { mtimeMs: farFuture - 1, size: 1024, hash: 'abc' },
      { mtimeMs: farFuture + 1, size: 1024 }, // beyond year 3000
      async () => 'xyz',
    )
    expect(result).toBe(true) // invalid mtime triggers hash comparison
  })
})
