import { existsSync, mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  checkAttachmentSize,
  contentAddress,
  INLINE_PREVIEW_CAP_BYTES,
  OUTGOING_WARN_BYTES,
  safeFilename,
  stageAttachment,
} from '../src/attachments'

describe('safeFilename', () => {
  it('strips directory traversal and separators', () => {
    expect(safeFilename('../../evil.exe')).toBe('evil.exe')
    expect(safeFilename('a/b\\c.txt')).toBe('c.txt')
    expect(safeFilename('...')).toBe('unnamed')
    expect(safeFilename('')).toBe('unnamed')
  })
})

describe('contentAddress', () => {
  it('is stable and content-bound', () => {
    const a = contentAddress(new TextEncoder().encode('hello'))
    const b = contentAddress(new TextEncoder().encode('hello'))
    const c = contentAddress(new TextEncoder().encode('other'))
    expect(a).toBe(b)
    expect(a).not.toBe(c)
    expect(a).toMatch(/^[0-9a-f]{64}$/)
  })
})

describe('checkAttachmentSize', () => {
  it('warns above 25 MB and rejects above 50 MB', () => {
    expect(checkAttachmentSize(1024)).toBe('ok')
    expect(checkAttachmentSize(OUTGOING_WARN_BYTES + 1)).toBe('warn')
    expect(checkAttachmentSize(INLINE_PREVIEW_CAP_BYTES + 1)).toBe('reject')
  })
})

describe('stageAttachment', () => {
  it('writes a content-addressed file and returns the staged handle', () => {
    const stageDir = mkdtempSync(join(tmpdir(), 'mail-stage-'))
    const bytes = new TextEncoder().encode('staged bytes')
    const staged = stageAttachment({
      filename: '../../notes.txt',
      mime: 'text/plain',
      bytes,
      stageDir,
    })
    expect(staged.filename).toBe('notes.txt')
    expect(staged.size).toBe(bytes.byteLength)
    expect(existsSync(staged.path)).toBe(true)
    expect(readFileSync(staged.path)).toEqual(Buffer.from(bytes))
    expect(staged.path.startsWith(stageDir)).toBe(true)
  })

  it('throws on oversize input without writing', () => {
    const stageDir = mkdtempSync(join(tmpdir(), 'mail-stage-'))
    const bytes = new Uint8Array(INLINE_PREVIEW_CAP_BYTES + 1)
    expect(() =>
      stageAttachment({ filename: 'big.bin', mime: 'application/octet-stream', bytes, stageDir }),
    ).toThrow()
  })
})
