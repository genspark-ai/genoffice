/**
 * OOXML schema gate (tools/ooxml-validate, xmllint + ISO/IEC 29500-4 schemas):
 *  - the gate itself catches the shapes of damage PowerPoint repairs (double fill, sz floor, a:ln order)
 *  - decks the engine builds from scratch validate clean
 *  - editing every element of a foreign deck adds no violation over the original
 */
import { describe, it, expect } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import JSZip from 'jszip'
import {
  validatePptx,
  newProblems,
  xmllintAvailable,
} from '../../../tools/ooxml-validate/validate-pptx.mjs'
import {
  openPptx,
  savePptx,
  addElement,
  addTable,
  createBlankPptx,
  setElementFont,
  setElementFill,
  setElementParagraphFormat,
  type TextElement,
} from '../src/index'

const available = xmllintAvailable()
if (!available && !process.env.CI) console.warn('xmllint not on PATH: ooxml-schema tests skipped')

describe.skipIf(!available && !process.env.CI)('OOXML schema gate', () => {
  const damage: Array<[string, (xml: string) => string, RegExp]> = [
    [
      'sz below 100',
      (x) => x.replace(/<a:rPr\b([^>]*?)\/>/, '<a:rPr$1 sz="50"/>'),
      /'sz'.*'50' is less than the minimum/,
    ],
    [
      'two fills in a:rPr',
      (x) =>
        x.replace(
          /<a:rPr\b([^>]*?)\/>/,
          '<a:rPr$1><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:noFill/></a:rPr>',
        ),
      /noFill.*not expected/,
    ],
    [
      'a:ln after a:effectLst',
      (x) => x.replace('</p:spPr>', '<a:effectLst/><a:ln><a:noFill/></a:ln></p:spPr>'),
      /\bln': This element is not expected/,
    ],
  ]
  it.each(damage)('flags %s', async (_label, damageXml, expected) => {
    const opened = await openPptx(await createBlankPptx())
    addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 1000, cy: 1000 },
      paragraphs: [{ runs: [{ text: 'x' }] }],
    })
    const zip = await JSZip.loadAsync(await savePptx(opened))
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    const bad = damageXml(xml)
    expect(bad).not.toBe(xml)
    zip.file('ppt/slides/slide1.xml', bad)
    const problems = await validatePptx(await zip.generateAsync({ type: 'uint8array' }))
    expect(problems.map((p) => p.message).join('\n')).toMatch(expected)
  })

  it('a deck built from scratch validates clean', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    addElement(slide, {
      kind: 'rect',
      offset: { x: 100, y: 100, cx: 2000000, cy: 1000000 },
      fillColor: '#1F77B480',
      stroke: { color: '#000000', widthEmu: 12700 },
      paragraphs: [{ runs: [{ text: 'hello', bold: true, fontSize: 24, color: '#FF0000' }] }],
    })
    addElement(slide, {
      kind: 'textbox',
      offset: { x: 100, y: 2000000, cx: 3000000, cy: 500000 },
      paragraphs: [
        { runs: [{ text: 'a' }, { text: 'b', italic: true }] },
        { runs: [{ text: '' }] },
      ],
    })
    addTable(opened, 0, {
      rows: 2,
      cols: 3,
      offset: { x: 100, y: 3000000, cx: 4000000, cy: 1000000 },
    })
    for (const el of slide.elements) {
      if (el.type === 'text' || el.type === 'shape') {
        setElementFont(slide, el.id, { color: '#00FF00', fontSizePt: 14 })
        setElementParagraphFormat(slide, el.id, { align: 'center' })
      }
      if (el.type === 'shape') setElementFill(opened, slide, el.id, '#123456')
    }
    expect(await validatePptx(await savePptx(opened))).toEqual([])
  })

  const fixtures = fs
    .readdirSync(path.join(__dirname, 'fixtures'))
    .filter((f) => f.endsWith('.pptx'))
  it.each(fixtures)('editing every element of %s adds no violation', async (name) => {
    const bytes = fs.readFileSync(path.join(__dirname, 'fixtures', name))
    const before = await validatePptx(bytes)
    const opened = await openPptx(bytes)
    let edits = 0
    for (const slide of opened.deck.slides) {
      for (const el of slide.elements) {
        if (el.type !== 'text' && el.type !== 'shape') continue
        if (!(el as TextElement).text?.paragraphs.length) continue
        if (setElementFont(slide, el.id, { color: '#FFFFFF', fontSizePt: 0.5 })) edits++
        setElementParagraphFormat(slide, el.id, { align: 'right' })
        if (el.type === 'shape') setElementFill(opened, slide, el.id, '#FF8800')
        el.transform.offset.x += 1
        el.dirtyTransform = true
      }
    }
    expect(edits).toBeGreaterThan(0)
    const after = await validatePptx(await savePptx(opened))
    expect(newProblems(before, after)).toEqual([])
  })
})
