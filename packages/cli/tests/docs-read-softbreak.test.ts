import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { buildBlankDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'

/** A docx whose body is one paragraph with a soft break (Shift+Enter, <w:br/>). */
async function docxWithSoftBreak(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await buildBlankDocx())
  const xml = await zip.file('word/document.xml')!.async('string')
  const paragraph =
    '<w:p><w:r><w:t>Date: 14 November 2026</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>Time: 3 pm</w:t></w:r></w:p>'
  zip.file('word/document.xml', xml.replace('<w:p/>', paragraph))
  return zip.generateAsync({ type: 'uint8array' })
}

describe('docs read soft line breaks', () => {
  it('keeps a <w:br/> as a newline in the block text', async () => {
    const dir = tempDir()
    const file = join(dir, 'invite.docx')
    writeFileSync(file, await docxWithSoftBreak())
    const read = await run(['docs', 'read', file, '--full', '--json'])
    expect(read.code).toBe(0)
    const item = read.json().detail.items[0] as { type: string; text: string }
    expect(item.type).toBe('paragraph')
    expect(item.text).toBe('Date: 14 November 2026\nTime: 3 pm')
  })

  it('does not turn a page break into a newline', async () => {
    const dir = tempDir()
    const file = join(dir, 'page.docx')
    const zip = await JSZip.loadAsync(await buildBlankDocx())
    const xml = await zip.file('word/document.xml')!.async('string')
    const paragraph =
      '<w:p><w:r><w:t>before</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>after</w:t></w:r></w:p>'
    zip.file('word/document.xml', xml.replace('<w:p/>', paragraph))
    writeFileSync(file, await zip.generateAsync({ type: 'uint8array' }))
    const read = await run(['docs', 'read', file, '--full', '--json'])
    expect(read.code).toBe(0)
    expect((read.json().detail.items[0] as { text: string }).text).toBe('beforeafter')
  })
})
