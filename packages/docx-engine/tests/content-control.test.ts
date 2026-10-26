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
