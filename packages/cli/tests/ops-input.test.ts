import { closeSync, mkdtempSync, openSync, truncateSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { MAX_TRANSFER_BYTES } from '../src/mcp/files'
import { MAX_OPS_BYTES, readOpsInput, readOpsStream } from '../src/ops-input'
import { EXIT } from '../src/result'

function ctx(cwd: string) {
  return { cwd, env: { ...process.env, GENOFFICE_ALLOWED_ROOTS: '' } }
}

/** Sparse file: stat reports the size without writing the bytes. */
function sparseFile(dir: string, name: string, bytes: number): string {
  const file = join(dir, name)
  writeFileSync(file, '')
  truncateSync(file, bytes)
  return file
}

describe('readOpsInput size cap', () => {
  it('shares the MCP transfer ceiling so base64 pictures in ops fit', () => {
    expect(MAX_OPS_BYTES).toBe(MAX_TRANSFER_BYTES)
    expect(MAX_OPS_BYTES).toBe(256 * 1024 * 1024)
  })

  it('rejects oversized ops files with resource_limit', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ops-'))
    const file = sparseFile(dir, 'big.json', MAX_OPS_BYTES + 1)
    let err: unknown
    try {
      readOpsInput({ positionals: [], flags: { ops: file } }, ctx(dir))
    } catch (e) {
      err = e
    }
    expect(err).toMatchObject({ code: EXIT.file, reason: 'resource_limit' })
  })

  it('stops an oversized stream at the cap while accumulating', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ops-'))
    const fd = openSync(sparseFile(dir, 'stream.json', MAX_OPS_BYTES + 1), 'r')
    try {
      expect(() => readOpsStream(fd, 'stdin')).toThrow(/ops input too large: stdin/)
    } finally {
      closeSync(fd)
    }
  })

  it('reads a small stream to its end', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ops-'))
    const file = join(dir, 'stream.json')
    writeFileSync(file, '{"ops":[]}')
    const fd = openSync(file, 'r')
    try {
      expect(readOpsStream(fd, 'stdin')).toBe('{"ops":[]}')
    } finally {
      closeSync(fd)
    }
  })

  it('keeps every chunk intact when the stream spans several reads', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ops-'))
    const file = join(dir, 'stream.json')
    const text = 'a'.repeat(1024 * 1024) + 'b'.repeat(1024 * 1024) + 'c'.repeat(512 * 1024)
    writeFileSync(file, text)
    const fd = openSync(file, 'r')
    try {
      expect(readOpsStream(fd, 'stdin')).toBe(text)
    } finally {
      closeSync(fd)
    }
  })

  it('accepts small ops files', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ops-'))
    const file = join(dir, 'small.json')
    writeFileSync(file, '{"ops":[]}')
    const r = readOpsInput({ positionals: [], flags: { ops: file } }, ctx(dir))
    expect(r.text).toBe('{"ops":[]}')
    expect(r.source).toBe(file)
  })
})
