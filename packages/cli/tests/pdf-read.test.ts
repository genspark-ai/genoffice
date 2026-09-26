import { writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { run, tempDir, writeMinimalPdf } from './helpers'

const REPO = resolve(__dirname, '../../..')
const ENCRYPTED = join(REPO, 'apps/shell/tests/fixtures/testPassword4Spaces.pdf')

describe('genoffice pdf read', () => {
  it('returns page count, metadata and per-page text with page sizes', async () => {
    const pdf = writeMinimalPdf(
      join(tempDir(), 'three.pdf'),
      ['Alpha page', 'Beta page', 'Gamma page'],
      { title: 'Quarterly notes', author: 'GenOffice' },
    )
    const r = await run(['pdf', 'read', pdf, '--json'])
    expect(r.code).toBe(0)
    const { detail, summary } = r.json()
    expect(summary).toContain('3 pages')
    expect(detail).toMatchObject({
      pages: 3,
      encrypted: false,
      title: 'Quarterly notes',
      author: 'GenOffice',
      range: '1-3',
    })
    expect(detail.truncated).toBeUndefined()
    expect(detail.pages_read).toHaveLength(3)
    expect(detail.pages_read[1]).toMatchObject({
      page: 2,
      width_pt: 612,
      height_pt: 792,
      text: 'Beta page',
      truncated: false,
    })
  })

  it('narrows to one page or a range and refuses both together', async () => {
    const pdf = writeMinimalPdf(join(tempDir(), 'three.pdf'), ['One', 'Two', 'Three'])
    const one = await run(['pdf', 'read', pdf, '--page', '2', '--json'])
    expect(one.json().detail.pages_read.map((p: { text: string }) => p.text)).toEqual(['Two'])
    const span = await run(['pdf', 'read', pdf, '--range', '2-3', '--json'])
    expect(span.json().detail.range).toBe('2-3')
    expect(span.json().detail.pages_read.map((p: { page: number }) => p.page)).toEqual([2, 3])
    const both = await run(['pdf', 'read', pdf, '--page', '1', '--range', '1-2', '--json'])
    expect(both.code).toBe(1)
    expect(both.json().error).toBe('invalid_argument')
  })

  it('reports out-of-range and malformed spans as usage errors with the valid range', async () => {
    const pdf = writeMinimalPdf(join(tempDir(), 'two.pdf'), ['One', 'Two'])
    const beyond = await run(['pdf', 'read', pdf, '--page', '9', '--json'])
    expect(beyond.code).toBe(1)
    expect(beyond.json()).toMatchObject({
      error: 'out_of_range',
      detail: { valid_range: [1, 2] },
    })
    const zero = await run(['pdf', 'read', pdf, '--range', '0-1', '--json'])
    expect(zero.json().error).toBe('out_of_range')
    const garbage = await run(['pdf', 'read', pdf, '--range', 'a-b', '--json'])
    expect(garbage.json().error).toBe('invalid_argument')
  })

  it('reads the first 20 pages by default and says how to get the rest', async () => {
    const texts = Array.from({ length: 25 }, (_, i) => `Page ${i + 1}`)
    const pdf = writeMinimalPdf(join(tempDir(), 'long.pdf'), texts)
    const r = await run(['pdf', 'read', pdf, '--json'])
    const detail = r.json().detail
    expect(detail.pages).toBe(25)
    expect(detail.pages_read).toHaveLength(20)
    expect(detail.range).toBe('1-20')
    expect(detail.truncated).toContain('--range 21-25')
    expect(detail.truncated).toContain('--full')

    const rest = await run(['pdf', 'read', pdf, '--range', '21-25', '--json'])
    expect(rest.json().detail.pages_read.map((p: { text: string }) => p.text)).toEqual(
      texts.slice(20),
    )
    expect(rest.json().detail.truncated).toBeUndefined()

    const full = await run(['pdf', 'read', pdf, '--full', '--json'])
    expect(full.json().detail.pages_read).toHaveLength(25)
    expect(full.json().detail.truncated).toBeUndefined()
  })

  it('clips page text to --max-chars and marks the page truncated', async () => {
    const pdf = writeMinimalPdf(join(tempDir(), 'one.pdf'), 'Hello genoffice readers')
    const r = await run(['pdf', 'read', pdf, '--max-chars', '5', '--json'])
    expect(r.json().detail.pages_read[0]).toMatchObject({
      text: 'Hello…(+18 chars)',
      truncated: true,
    })
    const full = await run(['pdf', 'read', pdf, '--full', '--json'])
    expect(full.json().detail.pages_read[0].text).toBe('Hello genoffice readers')
  })

  it('needs the right password for an encrypted PDF and reports it as encrypted', async () => {
    const locked = await run(['pdf', 'read', ENCRYPTED, '--json'])
    expect(locked.code).toBe(1)
    expect(locked.json()).toMatchObject({ error: 'missing_argument' })
    expect(locked.json().suggestion).toContain('--password')

    const wrong = await run(['pdf', 'read', ENCRYPTED, '--password', 'nope', '--json'])
    expect(wrong.code).toBe(1)
    expect(wrong.json().error).toBe('invalid_argument')

    const ok = await run(['pdf', 'read', ENCRYPTED, '--password', '    ', '--json'])
    expect(ok.code).toBe(0)
    expect(ok.json().detail).toMatchObject({ encrypted: true })
    expect(ok.json().detail.pages).toBeGreaterThan(0)
  })

  it('rejects a missing file, a non-PDF and an unknown verb with the documented reasons', async () => {
    const dir = tempDir()
    const missing = await run(['pdf', 'read', join(dir, 'none.pdf'), '--json'])
    expect(missing.code).toBe(2)
    expect(missing.json().error).toBe('file_not_found')

    const txt = join(dir, 'notes.txt')
    writeFileSync(txt, 'plain text')
    const notPdf = await run(['pdf', 'read', txt, '--json'])
    expect(notPdf.code).toBe(1)
    expect(notPdf.json().error).toBe('unsupported')

    const fake = join(dir, 'fake.pdf')
    writeFileSync(fake, 'not really a pdf')
    const corrupt = await run(['pdf', 'read', fake, '--json'])
    expect(corrupt.code).toBe(2)
    expect(corrupt.json().error).toBe('unsupported')

    const verb = await run(['pdf', 'count', fake, '--json'])
    expect(verb.json().error).toBe('invalid_argument')
    const none = await run(['pdf', '--json'])
    expect(none.json().error).toBe('missing_argument')
  })

  it('info shows the same metadata for a PDF', async () => {
    const pdf = writeMinimalPdf(join(tempDir(), 'meta.pdf'), 'Hi', { title: 'Titled' })
    const r = await run(['info', pdf, '--json'])
    expect(r.json().detail).toMatchObject({ pages: 1, encrypted: false, title: 'Titled' })
  })
})
