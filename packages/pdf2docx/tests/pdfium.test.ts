import { describe, expect, it } from 'vitest'
import type { PdfiumModule } from '../src/extract'
import { PdfLoadError, withPdfDocument } from '../src/extract'

/** Minimal module stub: only the surface withPdfDocument touches. */
function stubModule(opts: {
  loadResult?: number
  lastError?: number
  onLoad?: (ptr: number, size: number, password: number) => void
}): PdfiumModule & { freed: number[]; heap: Uint8Array } {
  const heap = new Uint8Array(1 << 16)
  const freed: number[] = []
  let next = 8
  return {
    HEAPU8: heap,
    heap,
    freed,
    _malloc: (size: number) => {
      const ptr = next
      next += size
      return ptr
    },
    _free: (ptr: number) => {
      freed.push(ptr)
    },
    _FPDF_LoadMemDocument: (ptr: number, size: number, password: number) => {
      opts.onLoad?.(ptr, size, password)
      return opts.loadResult ?? 0
    },
    _FPDF_GetLastError: () => opts.lastError ?? 0,
    _FPDF_CloseDocument: () => {},
  } as unknown as PdfiumModule & { freed: number[]; heap: Uint8Array }
}

describe('withPdfDocument load-failure classification (P22)', () => {
  it('maps FPDF_ERR_PASSWORD to a password-required PdfLoadError', () => {
    const m = stubModule({ loadResult: 0, lastError: 4 })
    let caught: unknown
    try {
      withPdfDocument(m, new Uint8Array([1, 2, 3]), () => 'unreachable')
    } catch (err) {
      caught = err
    }
    expect(caught).toBeInstanceOf(PdfLoadError)
    // backward compat: pre-P22 callers catch plain Error and print .message
    expect(caught).toBeInstanceOf(Error)
    const e = caught as PdfLoadError
    expect(e.code).toBe('password-required')
    expect(e.pdfiumError).toBe(4)
    expect(e.message).toContain('could not load')
  })

  it('maps FPDF_ERR_FORMAT / FPDF_ERR_FILE / FPDF_ERR_UNKNOWN to corrupt', () => {
    for (const lastError of [3, 2, 1]) {
      const m = stubModule({ loadResult: 0, lastError })
      expect(() => withPdfDocument(m, new Uint8Array(4), () => 0)).toThrowError(
        expect.objectContaining({ code: 'corrupt', pdfiumError: lastError }),
      )
    }
  })

  it('maps FPDF_ERR_SECURITY to unsupported', () => {
    const m = stubModule({ loadResult: 0, lastError: 5 })
    expect(() => withPdfDocument(m, new Uint8Array(4), () => 0)).toThrowError(
      expect.objectContaining({ code: 'unsupported' }),
    )
  })

  it('frees the pdf and password buffers on load failure', () => {
    const m = stubModule({ loadResult: 0, lastError: 4 })
    expect(() => withPdfDocument(m, new Uint8Array(4), () => 0, 'secret')).toThrow(PdfLoadError)
    expect(m.freed.length).toBe(2)
  })

  it('passes the password through as a NUL-terminated UTF-8 string', () => {
    let seen: { password: number } | null = null
    const m = stubModule({
      loadResult: 77,
      onLoad: (_ptr, _size, password) => {
        seen = { password }
      },
    })
    const heapAtLoad = m.heap
    const out = withPdfDocument(m, new Uint8Array([9]), (doc) => doc * 2, 'pässword')
    expect(out).toBe(154)
    expect(seen).not.toBeNull()
    const pwPtr = seen!.password
    expect(pwPtr).toBeGreaterThan(0)
    const expected = new TextEncoder().encode('pässword\0')
    expect([...heapAtLoad.slice(pwPtr, pwPtr + expected.length)]).toEqual([...expected])
  })

  it('omits the password pointer (0) when no password is given', () => {
    let pw = -1
    const m = stubModule({
      loadResult: 5,
      onLoad: (_p, _s, password) => {
        pw = password
      },
    })
    withPdfDocument(m, new Uint8Array(4), () => 0)
    expect(pw).toBe(0)
  })
})
