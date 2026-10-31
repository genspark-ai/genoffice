import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

// A block-level content control: a caption and a signature block that must
// not be reflowed, found downstream by their tag.
const LOCKED_SDT =
  '<w:sdt>' +
  '<w:sdtPr>' +
  '<w:alias w:val="Caption"/>' +
  '<w:tag w:val="example:caption"/>' +
  '<w:id w:val="4242"/>' +
  '<w:lock w:val="sdtLocked"/>' +
  '<w:richText/>' +
  '</w:sdtPr>' +
  '<w:sdtContent>' +
  '<w:p><w:r><w:t>Caption text</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>Signature block</w:t></w:r></w:p>' +
  '</w:sdtContent>' +
  '</w:sdt>'

const PLAIN_P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

const PLAIN_BODY =
  PLAIN_P('Before') + PLAIN_P('Caption text') + PLAIN_P('Signature block') + PLAIN_P('After')

const originalBlocks = [0, 1, 2, 3].map((docxIndex) => ({ kind: 'original' as const, docxIndex }))

/** the text of every visible block (the trailing sectPr is a hidden block) */
const textsOf = (doc: {
  blocks: Array<{ hidden?: boolean; runs?: Array<{ text: string }> }>
}): string[] =>
  doc.blocks.filter((b) => !b.hidden).map((b) => (b.runs ?? []).map((r) => r.text).join(''))

describe('content control metadata', () => {
  it('reads the lock and the id of a content control', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: PLAIN_P('Before') + LOCKED_SDT + PLAIN_P('After') }),
    )
    const shell = doc.blocks[1]?.sdtShell
    expect(shell).toBeDefined()
    expect(shell!.tag).toBe('example:caption')
    expect(shell!.alias).toBe('Caption')
    expect(shell!.lock).toBe('sdtLocked')
    expect(shell!.id).toBe('4242')
  })

  it('reports an empty lock and id when the control carries none', async () => {
    const sdt =
      '<w:sdt><w:sdtPr><w:tag w:val="plain"/><w:text/></w:sdtPr>' +
      '<w:sdtContent><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    const doc = await parseDocx(await buildDocx({ bodyXml: sdt }))
    expect(doc.blocks[0]?.sdtShell?.lock).toBe('')
    expect(doc.blocks[0]?.sdtShell?.id).toBe('')
  })
})

describe('content control wrapping', () => {
  it('wraps a block range in one content control', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PLAIN_BODY }))
    const saved = await saveDocx(doc, originalBlocks, {
      contentControls: {
        wrap: [
          {
            docxIndexes: [1, 2],
            spec: { tag: 'example:caption', title: 'Caption', lock: 'sdtLocked' },
          },
        ],
      },
    })
    const reparsed = await parseDocx(saved)
    const xml = reparsed.internal.documentXml
    expect(xml).toContain('<w:sdt>')
    expect(xml).toContain('<w:tag w:val="example:caption"/>')
    expect(xml).toContain('<w:alias w:val="Caption"/>')
    expect(xml).toContain('<w:lock w:val="sdtLocked"/>')
    // the wrapped blocks sit inside the control, the others stay outside it
    const sdt = /<w:sdt>[\s\S]*?<\/w:sdt>/.exec(xml)?.[0] ?? ''
    expect(sdt).toContain('Caption text')
    expect(sdt).toContain('Signature block')
    expect(sdt).not.toContain('Before')
    expect(sdt).not.toContain('After')
  })

  it('keeps the wrapped blocks editable and the text unchanged', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PLAIN_BODY }))
    const saved = await saveDocx(doc, originalBlocks, {
      contentControls: {
        wrap: [{ docxIndexes: [1, 2], spec: { tag: 'example:caption', lock: 'sdtLocked' } }],
      },
    })
    const reparsed = await parseDocx(saved)
    expect(textsOf(reparsed)).toEqual(['Before', 'Caption text', 'Signature block', 'After'])
    expect(reparsed.blocks[1]?.sdtShell?.tag).toBe('example:caption')
    expect(reparsed.blocks[2]?.sdtShell?.tag).toBe('example:caption')
    expect(reparsed.blocks[1]?.sdtShell?.lock).toBe('sdtLocked')
  })

  it('gives the new control a free id', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PLAIN_BODY }))
    const saved = await saveDocx(doc, originalBlocks, {
      contentControls: { wrap: [{ docxIndexes: [1], spec: { tag: 'one' } }] },
    })
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[1]?.sdtShell?.id).toBe('1')
  })
})

describe('content control unwrapping', () => {
  it('removes the wrapper and leaves the blocks as plain body children', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: PLAIN_P('Before') + LOCKED_SDT + PLAIN_P('After') }),
    )
    const saved = await saveDocx(
      doc,
      [0, 1, 2, 3].map((docxIndex) => ({ kind: 'original' as const, docxIndex })),
      {
        contentControls: { unwrap: [{ tag: 'example:caption' }] },
      },
    )
    const reparsed = await parseDocx(saved)
    const xml = reparsed.internal.documentXml
    expect(xml).not.toContain('<w:sdt>')
    expect(xml).not.toContain('example:caption')
    expect(textsOf(reparsed)).toEqual(['Before', 'Caption text', 'Signature block', 'After'])
    expect(reparsed.blocks[1]?.sdtShell).toBeUndefined()
  })

  it('strips the shell from a block the editor saved as its original xml', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: PLAIN_P('Before') + LOCKED_SDT + PLAIN_P('After') }),
    )
    const saved = await saveDocx(
      doc,
      [
        { kind: 'original' as const, docxIndex: 0 },
        { kind: 'xml' as const, xml: doc.blocks[1]!.originalXml!, docxIndex: 1 },
        { kind: 'xml' as const, xml: doc.blocks[2]!.originalXml! },
        { kind: 'original' as const, docxIndex: 3 },
      ],
      { contentControls: { unwrap: [{ tag: 'example:caption' }] } },
    )
    const reparsed = await parseDocx(saved)
    expect(reparsed.internal.documentXml).not.toContain('<w:sdt>')
    expect(textsOf(reparsed)).toEqual(['Before', 'Caption text', 'Signature block', 'After'])
  })
})

// ---- Nested block-level controls and inline (run-level) controls ----

// outer wraps two paragraphs around an inner control, and a run-level control
// sits inside the third paragraph (matches the genoffice#2062 nested fixture)
const NESTED_SDT =
  '<w:sdt><w:sdtPr><w:tag w:val="outer"/><w:alias w:val="Outer"/><w:id w:val="10"/>' +
  '<w:lock w:val="sdtLocked"/><w:richText/></w:sdtPr><w:sdtContent>' +
  PLAIN_P('outer para') +
  '<w:sdt><w:sdtPr><w:tag w:val="inner"/><w:alias w:val="Inner"/><w:id w:val="11"/>' +
  '<w:lock w:val="contentLocked"/><w:richText/></w:sdtPr><w:sdtContent>' +
  PLAIN_P('inner para') +
  '</w:sdtContent></w:sdt>' +
  '<w:p><w:r><w:t>inline </w:t></w:r>' +
  '<w:sdt><w:sdtPr><w:tag w:val="run"/><w:id w:val="12"/><w:text/></w:sdtPr>' +
  '<w:sdtContent><w:r><w:t>INLINE</w:t></w:r></w:sdtContent></w:sdt>' +
  '<w:r><w:t> tail</w:t></w:r></w:p>' +
  '</w:sdtContent></w:sdt>'

const NESTED_BODY = PLAIN_P('Before') + NESTED_SDT + PLAIN_P('After')

const tagsOf = (doc: Awaited<ReturnType<typeof parseDocx>>): string[][] =>
  doc.blocks
    .filter((b) => !b.hidden)
    .map((b) => (b.sdtShells ?? (b.sdtShell ? [b.sdtShell] : [])).map((shell) => shell.tag))

const saveAllOriginal = (doc: Awaited<ReturnType<typeof parseDocx>>) =>
  doc.blocks
    .filter((b) => !b.hidden)
    .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))

describe('nested content controls', () => {
  it('captures the control chain of every block (outer -> inner)', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: NESTED_BODY }))
    expect(textsOf(doc)).toEqual([
      'Before',
      'outer para',
      'inner para',
      'inlineINLINEtail',
      'After',
    ])
    expect(tagsOf(doc)).toEqual([[], ['outer'], ['outer', 'inner'], ['outer'], []])
    const inner = doc.blocks[2]!.sdtShells!
    expect(inner[1]!.level).toBe(1)
    expect(inner[1]!.alias).toBe('Inner')
    expect(inner[1]!.lock).toBe('contentLocked')
    expect(inner[1]!.id).toBe('11')
    // sdtShell stays the outermost control for existing consumers
    expect(doc.blocks[1]!.sdtShell!.tag).toBe('outer')
    expect(doc.blocks[2]!.sdtShell!.tag).toBe('outer')
    expect(doc.blocks[1]!.sdtShell!.group).toBe(doc.blocks[2]!.sdtShell!.group)
  })

  it('reports a run-level control without changing the run text', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: NESTED_BODY }))
    expect(doc.blocks[3]!.inlineContentControls).toEqual([
      { tag: 'run', title: '', lock: '', id: '12', text: 'INLINE' },
    ])
    expect(doc.blocks[3]!.runs?.map((r) => r.text).join('')).toBe('inlineINLINEtail')
    expect(doc.blocks[1]!.inlineContentControls).toBeUndefined()
  })

  it('captures an inner control wrapped around the single child paragraph', async () => {
    const single =
      '<w:sdt><w:sdtPr><w:tag w:val="outer"/></w:sdtPr><w:sdtContent>' +
      '<w:sdt><w:sdtPr><w:tag w:val="inner"/></w:sdtPr><w:sdtContent>' +
      PLAIN_P('only para') +
      '</w:sdtContent></w:sdt>' +
      '</w:sdtContent></w:sdt>'
    const bytes = await buildDocx({ bodyXml: single })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0]!.sdtShells!.map((shell) => shell.tag)).toEqual(['outer', 'inner'])
    const saved = await saveDocx(doc, saveAllOriginal(doc))
    expect(Buffer.from(saved).equals(Buffer.from(bytes))).toBe(true)
  })

  it('an untouched nested document saves byte-identical', async () => {
    const bytes = await buildDocx({ bodyXml: NESTED_BODY })
    const doc = await parseDocx(bytes)
    const saved = await saveDocx(doc, saveAllOriginal(doc))
    expect(Buffer.from(saved).equals(Buffer.from(bytes))).toBe(true)
  })

  it('unwraps an inner tag and keeps the outer shell around the content', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: NESTED_BODY }))
    const saved = await saveDocx(doc, saveAllOriginal(doc), {
      contentControls: { unwrap: [{ tag: 'inner' }] },
    })
    const reparsed = await parseDocx(saved)
    const xml = reparsed.internal.documentXml
    expect(xml).toContain('w:val="outer"')
    expect(xml).not.toContain('w:val="inner"')
    expect(textsOf(reparsed)).toEqual([
      'Before',
      'outer para',
      'inner para',
      'inlineINLINEtail',
      'After',
    ])
    expect(tagsOf(reparsed)).toEqual([[], ['outer'], ['outer'], ['outer'], []])
  })

  it('unwraps the outermost tag and leaves the inner control as its own shell', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: NESTED_BODY }))
    const saved = await saveDocx(doc, saveAllOriginal(doc), {
      contentControls: { unwrap: [{ tag: 'outer' }] },
    })
    const reparsed = await parseDocx(saved)
    const xml = reparsed.internal.documentXml
    expect(xml).not.toContain('w:val="outer"')
    expect(xml).toContain('w:val="inner"')
    expect(textsOf(reparsed)).toEqual([
      'Before',
      'outer para',
      'inner para',
      'inlineINLINEtail',
      'After',
    ])
    expect(tagsOf(reparsed)).toEqual([[], [], ['inner'], [], []])
  })

  it('keeps the nested controls byte-identical when a block outside them is edited', async () => {
    const bytes = await buildDocx({ bodyXml: NESTED_BODY })
    const doc = await parseDocx(bytes)
    const saved = await saveDocx(doc, [
      { kind: 'original' as const, docxIndex: 0 },
      { kind: 'original' as const, docxIndex: 1 },
      { kind: 'original' as const, docxIndex: 2 },
      { kind: 'original' as const, docxIndex: 3 },
      {
        kind: 'generated' as const,
        block: { type: 'paragraph', runs: [{ text: 'After edited' }] },
      },
    ])
    const original = Buffer.from(bytes).toString('latin1')
    const nested = original.slice(original.indexOf('<w:sdt>'), original.indexOf('</w:sdt>') + 7)
    const savedXml = (await parseDocx(saved)).internal.documentXml
    expect(savedXml).toContain('After edited')
    expect(savedXml).toContain(nested)
  })
})
