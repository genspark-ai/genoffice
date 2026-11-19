import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { run, tempDir } from './helpers'

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const DATE = '2026-01-01T00:00:00Z'
const revision = (kind: string, content: string, author = 'A') =>
  `<w:${kind} w:id="1" w:author="${author}" w:date="${DATE}">${content}</w:${kind}>`
const ref = '<w:r><w:footnoteReference w:id="1"/></w:r>'
const br = '<w:r><w:br/></w:r>'
const cases = [
  { name: 'footnote insertion alone', kind: 'ins', runs: ref, atom: 'footnoteReference' },
  { name: 'footnote deletion alone', kind: 'del', runs: ref, atom: 'footnoteReference' },
  { name: 'break insertion alone', kind: 'ins', runs: br, atom: 'br' },
  { name: 'break deletion alone', kind: 'del', runs: br, atom: 'br' },
  {
    name: 'endnote insertion alone',
    kind: 'ins',
    runs: '<w:r><w:endnoteReference w:id="1"/></w:r>',
    atom: 'endnoteReference',
  },
  {
    name: 'text and footnote in one insertion',
    kind: 'ins',
    runs: '<w:r><w:t>new </w:t></w:r>' + ref,
    atom: 'footnoteReference',
  },
  {
    name: 'footnote followed by a text revision',
    kind: 'ins',
    runs: ref,
    atom: 'footnoteReference',
    other: true,
  },
  {
    name: 'already supported tab insertion',
    kind: 'ins',
    runs: '<w:r><w:tab/></w:r>',
    atom: 'tab',
  },
]

async function fixture(dir: string, body: string) {
  const source = join(dir, 'source.docx')
  const notes = ['footnote', 'endnote'].map((kind) => ({
    path: `word/${kind}s.xml`,
    contentType: `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml`,
    xml: `<w:${kind}s xmlns:w="${W}"><w:${kind} w:id="1"><w:p><w:r><w:${kind}Ref/></w:r><w:r><w:t>Public synthetic note.</w:t></w:r></w:p></w:${kind}></w:${kind}s>`,
  }))
  writeFileSync(
    source,
    await buildDocx({
      bodyXml: body,
      extraRels: ['footnote', 'endnote']
        .map(
          (kind) =>
            `<Relationship Id="${kind}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s" Target="${kind}s.xml"/>`,
        )
        .join(''),
      extraParts: [
        ...notes,
        {
          path: 'customXml/sentinel.xml',
          contentType: 'application/xml',
          xml: '<untouched>Exact sentinel bytes</untouched>',
        },
      ],
    }),
  )
  return source
}

async function revisions(path: string) {
  const result = await run(['docs', 'read', path, '--revisions', '--json'])
  expect(result.code, result.stderr).toBe(0)
  return result.json().detail.revisions as Array<{ id: string; type: string; author: string }>
}

async function apply(source: string, out: string, op: Record<string, unknown>) {
  const ops = join(tempDir(), 'ops.json')
  writeFileSync(ops, JSON.stringify([op]))
  return run(['docs', 'apply', source, '--ops', ops, '--out', out, '--json'])
}

describe('tracked inline atoms through the real Docs CLI', () => {
  for (const mode of ['accept', 'reject'] as const) {
    it.each(cases)(`${mode}: $name is listed and resolved through save/reopen`, async (test) => {
      const dir = tempDir(),
        out = join(dir, 'result.docx')
      const other = test.other ? revision('ins', '<w:r><w:t>Other change</w:t></w:r>', 'B') : ''
      const source = await fixture(
        dir,
        `<w:p><w:r><w:t>Before.</w:t></w:r>${revision(test.kind, test.runs)}${other}<w:r><w:t>After.</w:t></w:r></w:p>`,
      )
      const listed = await revisions(source)
      expect(listed).toHaveLength(test.other ? 2 : 1)
      expect(listed[0]).toMatchObject({ author: 'A' })
      const result = await apply(source, out, { op: `${mode}_changes`, all: true })
      expect(result.code, result.stderr || result.stdout).toBe(0)
      expect(await revisions(out)).toEqual([])
      const before = await JSZip.loadAsync(readFileSync(source)),
        after = await JSZip.loadAsync(readFileSync(out))
      const xml = await after.file('word/document.xml')!.async('string')
      expect(xml).not.toMatch(/<w:(?:ins|del)\b/)
      expect(xml).toContain('Before.')
      expect(xml).toContain('After.')
      const keep = (mode === 'accept') === (test.kind === 'ins')
      expect(new RegExp(`<w:${test.atom}\\b`).test(xml)).toBe(keep)
      if (test.runs.includes('new ')) expect(xml.includes('new')).toBe(keep)
      if (test.other) expect(xml.includes('Other change')).toBe(mode === 'accept')
      for (const part of ['word/footnotes.xml', 'word/endnotes.xml', 'customXml/sentinel.xml']) {
        expect(await after.file(part)!.async('string')).toBe(
          await before.file(part)!.async('string'),
        )
      }
    })
  }

  it('selects an atom by revision id and leaves another author pending', async () => {
    const dir = tempDir(),
      out = join(dir, 'selected.docx')
    const source = await fixture(
      dir,
      `<w:p><w:r><w:t>Before.</w:t></w:r>${revision('ins', ref)}<w:r><w:t> between </w:t></w:r>${revision('ins', br, 'B')}<w:r><w:t>After.</w:t></w:r></w:p>`,
    )
    expect(await revisions(source)).toHaveLength(2)
    const result = await apply(source, out, { op: 'accept_changes', ids: ['r1'] })
    expect(result.code, result.stderr || result.stdout).toBe(0)
    expect(await revisions(out)).toMatchObject([{ author: 'B' }])
    const zip = await JSZip.loadAsync(readFileSync(out)),
      xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('footnoteReference')
    expect(xml.match(/<w:ins\b/g)).toHaveLength(1)
    expect(xml).toMatch(/<w:ins\b[^>]*w:author="B"/)
  })
})
