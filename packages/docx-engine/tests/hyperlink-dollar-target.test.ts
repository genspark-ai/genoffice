import JSZip from 'jszip'
import { SaxesParser } from 'saxes'
import { expect, it } from 'vitest'
import { buildBlankDocx, parseDocx, saveDocx, type SaveBlock } from '../src/index'

const relsPath = 'word/_rels/document.xml.rels'
function relationships(xml: string) {
  const entries: Record<string, string>[] = []
  const parser = new SaxesParser({ xmlns: true })
  parser.on('opentag', (node) => {
    if (node.local === 'Relationship') {
      entries.push(
        Object.fromEntries(Object.entries(node.attributes).map(([key, attr]) => [key, attr.value])),
      )
    }
  })
  parser.write(xml).close()
  return entries
}

it.each(['$&', '$$', '$`', "$'", '$1'])(
  'saves a hyperlink containing literal %s without changing existing relationships',
  async (literal) => {
    const existingHref = 'https://example.com/existing'
    const blank = await parseDocx(await buildBlankDocx())
    const seeded = await saveDocx(blank, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: 'Existing', link: { href: existingHref } }],
        },
      },
    ])
    const parsed = await parseDocx(seeded)
    const before = await JSZip.loadAsync(seeded)
    const oldRelationships = relationships(await before.file(relsPath)!.async('string'))
    const href =
      literal === '$&'
        ? 'https://example.com/?total=100$&currency=USD'
        : `https://example.com/?total=100${literal}&currency=USD`
    const saved = await saveDocx(parsed, [
      ...parsed.blocks.map((block): SaveBlock => ({
        kind: 'original',
        docxIndex: block.docxIndex!,
      })),
      {
        kind: 'generated',
        block: { type: 'paragraph', runs: [{ text: 'Invoice', link: { href } }] },
      },
    ])
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file(relsPath)!.async('string')
    const entries = relationships(xml)
    expect(entries).toHaveLength(oldRelationships.length + 1)
    for (const original of oldRelationships) expect(entries).toContainEqual(original)
    expect(entries.find((entry) => entry.Target === href)).toMatchObject({
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
      TargetMode: 'External',
    })
    for (const part of Object.keys(before.files)) {
      if (
        before.files[part].dir ||
        part === relsPath ||
        part === 'word/document.xml' ||
        part === 'docProps/core.xml'
      )
        continue
      expect(await zip.file(part)!.async('uint8array')).toEqual(
        await before.file(part)!.async('uint8array'),
      )
    }
    const reopened = await parseDocx(saved)
    expect(
      reopened.blocks.flatMap((block) => block.runs ?? []).map((run) => run.link?.href),
    ).toEqual([existingHref, href])
    const unchanged = await saveDocx(
      reopened,
      reopened.blocks.map((block): SaveBlock => ({
        kind: 'original',
        docxIndex: block.docxIndex!,
      })),
    )
    const again = await JSZip.loadAsync(unchanged)
    expect(await again.file(relsPath)!.async('string')).toBe(xml)
    expect(
      (await parseDocx(unchanged)).blocks
        .flatMap((block) => block.runs ?? [])
        .map((run) => run.link?.href),
    ).toEqual([existingHref, href])
  },
)
