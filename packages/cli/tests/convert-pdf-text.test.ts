import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { run, tempDir, writeMinimalPdf } from './helpers'

describe('genoffice convert pdf → md', () => {
  it('reconstructs the text through the docx route', async () => {
    const dir = tempDir()
    const pdf = writeMinimalPdf(join(dir, 'notes.pdf'), ['Alpha line', 'Beta line'])
    const r = await run(['convert', pdf, '--to', 'md', '--json'])
    expect(r.code).toBe(0)
    const md = readFileSync(join(dir, 'notes.md'), 'utf8')
    expect(md).toContain('Alpha line')
    expect(md).toContain('Beta line')
    const detail = r.json().detail
    expect(detail.intermediate).toBe('docx')
    expect(detail.pages).toBe(2)
  }, 60_000)

  it('lists html and md among the pdf routes', async () => {
    const r = await run([
      'convert',
      writeMinimalPdf(join(tempDir(), 'x.pdf')),
      '--to',
      'rtf',
      '--json',
    ])
    expect(r.code).not.toBe(0)
    expect(r.json().detail.supported).toMatch(/pdf→docx\/pptx\/xlsx\/md\/html/)
  })
})
