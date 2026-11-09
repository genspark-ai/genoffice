import { describe, expect, it } from 'vitest'

import { csvGridShape, type CsvImportProgress } from '../src/gateway/csv-import'
import { csvToXlsxBufferForOpen } from '../src/gateway/csv-import-xlsx'

/**
 * A delimited import reports how far it is and can be stopped: a 250 MB file
 * takes a while, and the user should see the bar move and be able to walk
 * away.
 */

const CSV = 'name,qty\nwidget,3\ngadget,5\n'

describe('the delimited import progress', () => {
  it('reports the read and the write phase', async () => {
    const seen: CsvImportProgress[] = []
    await csvToXlsxBufferForOpen(CSV, 'Sheet1', ',', {
      onProgress: (progress) => seen.push(progress),
    })
    expect(seen.length).toBeGreaterThan(0)
    expect(seen.some((progress) => progress.phase === 'read')).toBe(true)
    expect(seen.some((progress) => progress.phase === 'write')).toBe(true)
    // both phases end at the end of their work
    const last = seen[seen.length - 1]!
    expect(last.phase).toBe('write')
    expect(last.fraction).toBeGreaterThan(0)
  })

  it('reports a fraction that only grows', async () => {
    const seen: CsvImportProgress[] = []
    await csvToXlsxBufferForOpen(CSV, 'Sheet1', ',', {
      onProgress: (progress) => seen.push(progress),
    })
    const fractions = seen.map((progress) =>
      progress.phase === 'read' ? progress.fraction : 1 + progress.fraction,
    )
    for (let index = 1; index < fractions.length; index += 1) {
      expect(fractions[index]!).toBeGreaterThanOrEqual(fractions[index - 1]!)
    }
  })

  it('stops when the import is aborted', async () => {
    const controller = new AbortController()
    const promise = csvToXlsxBufferForOpen(CSV, 'Sheet1', ',', {
      onProgress: () => controller.abort(),
      signal: controller.signal,
    })
    await expect(promise).rejects.toThrow(/abort/i)
  })

  it('stops during the write phase, where the time goes', async () => {
    // The read is a pass over the text; the write is the archive, and the
    // archive is the seconds (0.7 s against 12 s on a 39 MB file). An abort
    // that only works during the read leaves the user waiting out the
    // compression, which is the part they wanted to skip.
    const controller = new AbortController()
    const promise = csvToXlsxBufferForOpen(CSV, 'Sheet1', ',', {
      onProgress: (progress) => {
        if (progress.phase === 'write') controller.abort()
      },
      signal: controller.signal,
    })
    await expect(promise).rejects.toThrow(/abort/i)
  })

  it('aborts the read phase too', () => {
    const controller = new AbortController()
    controller.abort()
    expect(() => csvGridShape(CSV, ',', { signal: controller.signal })).toThrow(/abort/i)
  })
})
