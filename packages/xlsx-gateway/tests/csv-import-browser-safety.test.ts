import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

/**
 * The CSV parser is shared with the sheets renderer (Data › From Text/CSV),
 * and the renderer is bundled for the browser: a node API in that module
 * breaks the browser build ("Readable" is not exported by
 * "__vite-browser-external"), which no unit test sees. The archive writer
 * therefore lives in its own module, and this pins the split.
 */

const here = dirname(fileURLToPath(import.meta.url))
const read = (rel: string): string => readFileSync(join(here, '..', rel), 'utf8')

describe('the CSV import modules', () => {
  it('keeps the parser the renderer imports free of node APIs', () => {
    const parser = read('src/gateway/csv-import.ts')
    expect(parser).not.toMatch(/from 'node:/)
  })

  it('holds the archive writer, and the node stream, in the desktop module', () => {
    const writer = read('src/gateway/csv-import-xlsx.ts')
    expect(writer).toMatch(/from 'node:stream'/)
    expect(writer).toMatch(/export async function csvToXlsxBufferForOpen/)
  })
})
