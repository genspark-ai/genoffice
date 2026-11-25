/**
 * The range a pasted link points at: a bookmark that covers text across
 * paragraphs, written by its two ends and read back as the same range, so
 * editing one of its paragraphs leaves the range whole.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, pairRangeBookmarks, saveDocx } from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
const BODY = ['Intro stays', 'First linked', 'Second linked', 'Outro stays'].map(p).join('')
const MARK = { type: 'docRangeBookmark', attrs: { name: 'LinkSource1' } }

async function saved(doc: PmNode, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(doc, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  return { bytes, xml }
}

const marked = (doc: PmNode): string[] =>
  (doc.content ?? []).flatMap((block) =>
    (block.content ?? [])
      .filter((node) => node.marks?.some((m) => m.type === MARK.type))
      .map((node) => node.text ?? ''),
  )

describe('a bookmark over a range of text', () => {
  it('pairs the two ends under one id and drops an end that has no other', () => {
    const xml =
      '<w:bookmarkStart w:id="4" w:name="Old"/><w:bookmarkEnd w:id="4"/>' +
      '<w:bookmarkStart w:id="" w:name="A"/>x<w:bookmarkStart w:id="" w:name="B"/>y' +
      '<w:bookmarkEnd w:id="" w:name="A"/><w:bookmarkEnd w:id="" w:name="C"/>'
    expect(pairRangeBookmarks(xml)).toBe(
      '<w:bookmarkStart w:id="4" w:name="Old"/><w:bookmarkEnd w:id="4"/>' +
        '<w:bookmarkStart w:id="5" w:name="A"/>xy<w:bookmarkEnd w:id="5"/>',
    )
  })

  it('is written around exactly the marked text and read back as the same range', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const doc = blocksToPmDoc(parsed.blocks)
    for (const block of doc.content!.slice(1, 3))
      for (const node of block.content!) node.marks = [...(node.marks ?? []), MARK]
    const first = await saved(doc, parsed)
    const start = /<w:bookmarkStart w:id="(\d+)" w:name="LinkSource1"\/>/.exec(first.xml)!
    const order = [
      'Intro stays',
      start[0],
      'First linked',
      'Second linked',
      `<w:bookmarkEnd w:id="${start[1]}"/>`,
      'Outro stays',
    ].map((s) => first.xml.indexOf(s))
    expect(order.every((at) => at > 0)).toBe(true)
    expect([...order].sort((a, b) => a - b)).toEqual(order)

    // read again: the same two paragraphs carry the mark
    const reparsed = await parseDocx(first.bytes)
    const again = blocksToPmDoc(reparsed.blocks)
    expect(marked(again)).toEqual(['First linked', 'Second linked'])

    // saved again untouched, the range is where it was
    const body = (xml: string) =>
      /<w:body>[\s\S]*<\/w:body>/.exec(xml)?.[0].replace(/w:id="\d+"/g, 'w:id=""')
    expect(body((await saved(again, reparsed)).xml)).toBe(body(first.xml))

    // the first paragraph of the range is edited: the range still ends after the second
    const text = again.content![1]!.content![0]!
    text.text = 'First linked, now longer'
    const edited = await saved(again, reparsed)
    expect(edited.xml.match(/w:name="LinkSource1"/g)).toHaveLength(1)
    const id = /<w:bookmarkStart w:id="(\d+)" w:name="LinkSource1"\/>/.exec(edited.xml)![1]
    expect(edited.xml.match(new RegExp(`<w:bookmarkEnd w:id="${id}"/>`, 'g'))).toHaveLength(1)
    const after = [
      `w:name="LinkSource1"`,
      'First linked, now longer',
      'Second linked',
      `<w:bookmarkEnd w:id="${id}"/>`,
      'Outro stays',
    ].map((s) => edited.xml.indexOf(s))
    expect([...after].sort((a, b) => a - b)).toEqual(after)
    expect(marked(blocksToPmDoc((await parseDocx(edited.bytes)).blocks))).toEqual([
      'First linked, now longer',
      'Second linked',
    ])
  })

  it('is read whatever the order of the attributes and the form of the end tag', async () => {
    const body =
      '<w:p><w:bookmarkStart w:name="LinkSource2" w:colFirst="0" w:id="7"></w:bookmarkStart><w:r><w:t>One</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Two</w:t></w:r><w:bookmarkEnd w:id="7"></w:bookmarkEnd></w:p>' +
      '<w:p><w:r><w:t>Three</w:t></w:r></w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    expect(marked(blocksToPmDoc(parsed.blocks))).toEqual(['One', 'Two'])
  })

  it('is not a range without its end, and many of them do not slow the read', async () => {
    const open = Array.from(
      { length: 20000 },
      (_, i) => `<w:bookmarkStart w:id="${i + 10}" w:name="LinkSource${i + 10}"/>`,
    ).join('')
    const body =
      `<w:p>${open}<w:bookmarkStart w:id="7" w:name="LinkSource7"/><w:r><w:t>One</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>` +
      '<w:p><w:r><w:t>Two</w:t></w:r></w:p>'
    const bytes = await buildDocx({ bodyXml: body })
    const started = performance.now()
    const parsed = await parseDocx(bytes)
    expect(performance.now() - started).toBeLessThan(5000)
    expect(marked(blocksToPmDoc(parsed.blocks))).toEqual(['One'])
  })

  it('a bookmark of any other name is read as it always was', async () => {
    const body =
      '<w:p><w:bookmarkStart w:id="1" w:name="OLE_LINK1"/><w:r><w:t>One</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Two</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    expect(
      parsed.blocks.flatMap((b) => ('runs' in b ? (b.runs ?? []) : [])).map((r) => r.text),
    ).toEqual(['One', 'Two'])
    expect(marked(blocksToPmDoc(parsed.blocks))).toEqual([])
  })
})
