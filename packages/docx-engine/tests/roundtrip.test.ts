import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildKitchenSinkDocx } from './helpers/build-docx'

async function entryBytes(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(bytes)
  const out = new Map<string, Uint8Array>()
  for (const [name, entry] of Object.entries(zip.files)) {
    if (!entry.dir) out.set(name, await entry.async('uint8array'))
  }
  return out
}

function asOriginal(indexes: number[]): SaveBlock[] {
  return indexes.map((docxIndex) => ({ kind: 'original', docxIndex }))
}

describe('paragraph-patch save fidelity', () => {
  it('no edits -> output is the exact original file (byte identical)', async () => {
    const bytes = await buildKitchenSinkDocx()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden).map((b) => b.docxIndex!)
    const saved = await saveDocx(doc, asOriginal(visible))
    expect(saved).toBe(bytes)
  })

  it('editing one paragraph leaves every other entry and every other w:p byte-identical', async () => {
    const bytes = await buildKitchenSinkDocx()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)

    // replace paragraph at docxIndex 1 (the rich text paragraph)
    const finalBlocks: SaveBlock[] = visible.map((b) =>
      b.docxIndex === 1
        ? {
            kind: 'generated',
            block: {
              type: 'paragraph',
              runs: [{ text: 'AI 重写后的段落。', bold: false }],
            },
          }
        : { kind: 'original', docxIndex: b.docxIndex! },
    )
    const saved = await saveDocx(doc, finalBlocks)

    const before = await entryBytes(bytes)
    const after = await entryBytes(saved)
    expect([...after.keys()].sort()).toEqual([...before.keys()].sort())
    for (const [name, content] of before) {
      if (name === 'word/document.xml') continue
      expect(after.get(name), name).toEqual(content)
    }

    // untouched top-level elements keep their exact original XML slices
    const newDocXml = new TextDecoder().decode(after.get('word/document.xml')!)
    for (const block of doc.blocks) {
      if (block.docxIndex === 1) continue
      expect(newDocXml).toContain(block.originalXml!)
    }
    expect(newDocXml).toContain('AI 重写后的段落。')
    expect(newDocXml).not.toContain('红色14号')

    // and the patched file still parses
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks.filter((b) => !b.hidden)).toHaveLength(visible.length)
  })

  it('supports insert / delete / heading generation with existing styles', async () => {
    const bytes = await buildKitchenSinkDocx()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)

    const finalBlocks: SaveBlock[] = []
    for (const b of visible) {
      if (b.docxIndex === 10) continue // delete the trailing paragraph
      finalBlocks.push({ kind: 'original', docxIndex: b.docxIndex! })
      if (b.docxIndex === 0) {
        finalBlocks.push({
          kind: 'generated',
          block: { type: 'heading', level: 2, runs: [{ text: '新增小节', bold: true }] },
        })
        finalBlocks.push({
          kind: 'generated',
          block: {
            type: 'listItem',
            list: { kind: 'bullet', numId: '1', ilvl: 0 },
            runs: [{ text: '新增列表项' }],
          },
        })
      }
    }
    const saved = await saveDocx(doc, finalBlocks)
    const newDocXml = new TextDecoder().decode(
      (await entryBytes(saved)).get('word/document.xml')!,
    )
    expect(newDocXml).toContain('<w:pStyle w:val="Heading2"/>')
    expect(newDocXml).toContain('<w:pStyle w:val="ListParagraph"/>')
    expect(newDocXml).toContain('<w:numId w:val="1"/>')
    expect(newDocXml).not.toContain('尾段。')
    // sectPr survives exactly once, at the end of the body
    expect(newDocXml.match(/<w:sectPr>/g)).toHaveLength(1)
    expect(newDocXml).toMatch(/<w:sectPr>[\s\S]*<\/w:sectPr><\/w:body>/)

    const reparsed = await parseDocx(saved)
    const types = reparsed.blocks.filter((b) => !b.hidden).map((b) => b.type)
    expect(types[1]).toBe('heading')
    expect(types[2]).toBe('listItem')
  })

  it('new hyperlinks get relationships appended to document.xml.rels', async () => {
    const bytes = await buildKitchenSinkDocx()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)

    const finalBlocks: SaveBlock[] = visible.map((b) =>
      b.docxIndex === 10
        ? {
            kind: 'generated',
            block: {
              type: 'paragraph',
              runs: [
                { text: '访问 ' },
                { text: '新链接', link: { href: 'https://new.example.org/' } },
              ],
            },
          }
        : { kind: 'original', docxIndex: b.docxIndex! },
    )
    const saved = await saveDocx(doc, finalBlocks)
    const entries = await entryBytes(saved)
    const rels = new TextDecoder().decode(entries.get('word/_rels/document.xml.rels')!)
    expect(rels).toContain('Target="https://new.example.org/"')
    expect(rels).toContain('TargetMode="External"')

    const relId = /<Relationship Id="(rId\d+)"[^>]*new\.example\.org/.exec(rels)![1]
    const docXml = new TextDecoder().decode(entries.get('word/document.xml')!)
    expect(docXml).toContain(`<w:hyperlink r:id="${relId}">`)

    const reparsed = await parseDocx(saved)
    const lastPara = reparsed.blocks.filter((b) => !b.hidden).at(-1)!
    expect(lastPara.runs?.find((r) => r.link)?.link?.href).toBe('https://new.example.org/')
  })

  it('passthrough blocks (table / image / math) can be moved as whole units', async () => {
    const bytes = await buildKitchenSinkDocx()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)
    const tableIdx = visible.find((b) => b.type === 'table')!.docxIndex!
    const order = visible.map((b) => b.docxIndex!).filter((i) => i !== tableIdx)
    order.unshift(tableIdx) // move table to the top

    const saved = await saveDocx(doc, asOriginal(order))
    const reparsed = await parseDocx(saved)
    const newVisible = reparsed.blocks.filter((b) => !b.hidden)
    expect(newVisible[0].type).toBe('table')
    expect(newVisible[0].originalXml).toBe(doc.blocks[tableIdx].originalXml)
  })
})
