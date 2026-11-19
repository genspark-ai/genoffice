import { copyFileSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../../docx-engine/src/index'
import { run, tempDir } from './helpers'

const FIXTURE = fileURLToPath(new URL('./fixtures/legacy.doc', import.meta.url))

async function zipText(path: string, entry: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return (await zip.file(entry)?.async('string')) ?? ''
}

function paragraphText(block: { runs?: Array<{ text: string }> }): string {
  return (block.runs ?? []).map((run) => run.text).join('')
}

describe('genoffice convert legacy .doc import', () => {
  it('converts .doc to .docx headlessly and keeps text, lists and hyperlinks', async () => {
    const dir = tempDir()
    const out = join(dir, 'legacy.docx')
    const r = await run(['convert', FIXTURE, '--to', 'docx', '--out', out, '--json'])
    expect(r.code).toBe(0)
    const bytes = new Uint8Array(readFileSync(out))
    expect(String.fromCharCode(bytes[0]!, bytes[1]!)).toBe('PK')
    const document = await zipText(out, 'word/document.xml')
    expect(document).toContain('Bullet A')
    expect(document).toContain('Bullet B')
    expect(document).toContain('Number 1')
    expect(document).toContain('the site')
    expect(document).toContain('<w:numPr>')
    expect(document).toContain('<w:hyperlink r:id=')
    const numbering = await zipText(out, 'word/numbering.xml')
    expect(numbering).toContain('<w:numFmt w:val="bullet"/>')
    expect(numbering).toContain('<w:numFmt w:val="decimal"/>')
    expect(r.json().detail.warnings).toEqual([])

    const parsed = await parseDocx(bytes)
    const paragraphs = parsed.blocks.filter(
      (block) => block.type === 'paragraph' || block.type === 'listItem',
    )
    expect(paragraphs.map(paragraphText).filter((text) => text !== '')).toEqual([
      'Bullet A',
      'Bullet B',
      'Number 1',
      'Visit the site',
    ])
    const listItems = paragraphs.filter((block) => block.list)
    expect(listItems.map((block) => block.list?.kind)).toEqual(['bullet', 'bullet', 'ordered'])
    expect(listItems.every((block) => block.list?.ilvl === 0)).toBe(true)
    const linkRun = paragraphs.flatMap((block) => block.runs ?? []).find((run) => run.link)
    expect(linkRun?.text).toBe('the site')
    expect(linkRun?.link?.href).toBe('https://example.org/')
  })

  it('accepts .dot templates through the same importer', async () => {
    const dir = tempDir()
    const dot = join(dir, 'legacy.dot')
    copyFileSync(FIXTURE, dot)
    const out = join(dir, 'legacy-dot.docx')
    const r = await run(['convert', dot, '--to', 'docx', '--out', out, '--json'])
    expect(r.code).toBe(0)
    expect(await zipText(out, 'word/document.xml')).toContain('Bullet A')
    expect(r.json().detail.warnings).toEqual([])
  })

  it('keeps the unsupported-route and conversion-error conventions', async () => {
    const bogus = await run(['convert', FIXTURE, '--to', 'pptx', '--json'])
    expect(bogus.code).toBe(1)
    expect(bogus.json().error).toBe('unsupported')
    expect(bogus.json().message).toContain('cannot convert .doc to .pptx')

    const dir = tempDir()
    const broken = join(dir, 'broken.doc')
    writeFileSync(broken, 'not a compound document')
    const r = await run([
      'convert',
      broken,
      '--to',
      'docx',
      '--out',
      join(dir, 'broken.docx'),
      '--json',
    ])
    expect(r.code).toBe(3)
    const payload = r.json()
    expect(payload.status).toBe('error')
    expect(payload.error).toBe('conversion_failed')
    expect(payload.detail.reason).toBe('not_cfb')
  })

  it('leaves the other route tables untouched', async () => {
    const help = await run(['help', 'convert'])
    expect(help.code).toBe(0)
    expect(help.stdout).toContain('doc→docx')
    expect(help.stdout).toContain('dot→docx')
    expect(help.stdout).toContain('docx→md')
    expect(help.stdout).toContain('odt→docx')
  })
})
