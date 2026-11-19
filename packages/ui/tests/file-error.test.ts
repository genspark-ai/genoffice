import { describe, expect, it } from 'vitest'
import { shortFileError } from '../src/file-error'

describe('shortFileError', () => {
  it('returns the code of a Node error', () => {
    const err = Object.assign(new Error('ENOSPC: no space left on device, write'), {
      code: 'ENOSPC',
    })
    expect(shortFileError(err)).toBe('ENOSPC')
  })

  it('finds the code in an error that crossed IPC as a string', () => {
    expect(shortFileError("Error: EBUSY: resource busy or locked, open '/a/b.docx'")).toBe('EBUSY')
    expect(shortFileError("EACCES: permission denied, open '/a/b.docx'")).toBe('EACCES')
    expect(shortFileError(new Error("EPERM: operation not permitted, rename 'a' -> 'b'"))).toBe(
      'EPERM',
    )
  })

  it('keeps a short message and cuts a long one', () => {
    expect(shortFileError('Canceled')).toBe('Canceled')
    expect(shortFileError('Export: the file is not a document')).toBe(
      'Export: the file is not a document',
    )
    const long = 'x'.repeat(80)
    expect(shortFileError(long)).toBe(`${'x'.repeat(60)}\u2026`)
    expect(shortFileError(null)).toBe('')
  })
})
