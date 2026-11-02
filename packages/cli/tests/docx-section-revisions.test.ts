import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { run, tempDir } from './helpers'

const SECTION_CHANGE =
  '<w:sectPrChange w:id="5" w:author="SectionEditor" w:date="2026-01-01T00:00:00Z">' +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1800" w:right="1800" w:bottom="1800" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr></w:sectPrChange>'

async function fixture(bodyRevision = false) {
  const dir = tempDir()
  const input = join(dir, 'section.docx')
  const zip = await JSZip.loadAsync(
    await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>Original body</w:t></w:r></w:p>' +
        (bodyRevision
          ? '<w:p><w:ins w:id="6" w:author="BodyEditor"><w:r><w:t>Other insertion</w:t></w:r></w:ins></w:p>'
          : ''),
    }),
  )
  const xml = await zip.file('word/document.xml')!.async('string')
  zip.file('word/document.xml', xml.replace('</w:sectPr>', `${SECTION_CHANGE}</w:sectPr>`))
  const bytes = await zip.generateAsync({ type: 'uint8array' })
  writeFileSync(input, bytes)
  return { dir, input, bytes }
}

async function apply(f: Awaited<ReturnType<typeof fixture>>, ops: Record<string, unknown>[]) {
  const opsFile = join(f.dir, 'ops.json')
  const output = join(f.dir, 'result.docx')
  writeFileSync(opsFile, JSON.stringify(ops))
  const result = await run(['docs', 'apply', f.input, '--ops', opsFile, '--out', output, '--json'])
  expect(result.code, result.stderr || result.stdout).toBe(0)
  const zip = await JSZip.loadAsync(readFileSync(output))
  const original = await JSZip.loadAsync(f.bytes)
  for (const part of ['word/styles.xml', 'word/_rels/document.xml.rels']) {
    expect(await zip.file(part)!.async('uint8array')).toEqual(
      await original.file(part)!.async('uint8array'),
    )
  }
  return { result, output, xml: await zip.file('word/document.xml')!.async('string') }
}

describe('CLI trailing section revisions', () => {
  it('accepts the trailing section change and keeps live page properties', async () => {
    const f = await fixture()
    const { result, output, xml } = await apply(f, [{ op: 'accept_changes', all: true }])
    expect(result.json().detail.results[0].output).toContain('Accepted 1 tracked change(s)')
    expect(xml).not.toContain('<w:sectPrChange')
    expect(xml).toContain('<w:pgSz w:w="11906" w:h="16838"/>')
    expect(xml).toContain('w:top="1440"')
    expect(xml).toContain('Original body')
    const read = await run(['docs', 'read', output, '--revisions', '--json'])
    expect(read.code).toBe(0)
    expect(read.json().detail.revisions).toEqual([])
  })

  it('rejects the trailing section change using its recorded page properties', async () => {
    const { xml } = await apply(await fixture(), [{ op: 'reject_changes', all: true }])
    expect(xml).not.toContain('<w:sectPrChange')
    expect(xml).toContain('<w:pgSz w:w="12240" w:h="15840"/>')
    expect(xml).toContain('w:top="1800"')
    expect(xml).not.toContain('w:top="1440"')
    expect(xml).toContain('Original body')
  })

  it('resolves only the selected section author and preserves another author’s insertion', async () => {
    const { xml } = await apply(await fixture(true), [
      { op: 'accept_changes', author: 'SectionEditor' },
    ])
    expect(xml).not.toContain('<w:sectPrChange')
    expect(xml).toContain('<w:ins w:id="6" w:author="BodyEditor">')
    expect(xml).toContain('Other insertion')
  })

  it('leaves the unselected section record byte-for-byte when accepting body changes', async () => {
    const { xml } = await apply(await fixture(true), [
      { op: 'accept_changes', author: 'BodyEditor' },
    ])
    expect(xml).toContain(SECTION_CHANGE)
    expect(xml).not.toContain('<w:ins ')
    expect(xml).toContain('Other insertion')
  })

  it('keeps earlier body edits when accepting the section in the same batch', async () => {
    const { xml } = await apply(await fixture(), [
      { op: 'insert_content', html: '<p>Before acceptance</p>' },
      { op: 'accept_changes', all: true },
    ])
    expect(xml).not.toContain('<w:sectPrChange')
    expect(xml).toContain('Original body')
    expect(xml).toContain('Before acceptance')
    expect(xml).toContain('w:top="1440"')
  })

  it('preserves an unresolved section record when only the body is edited', async () => {
    const { xml } = await apply(await fixture(), [
      { op: 'insert_content', html: '<p>Plain body edit</p>' },
    ])
    expect(xml).toContain(SECTION_CHANGE)
    expect(xml).toContain('Plain body edit')
  })
})
