import { describe, expect, it } from 'vitest'
import {
  addElement,
  createBlankPptx,
  openPptx,
  patchSlideXml,
  savePptx,
  type OpenedPptx,
  type TextElement,
} from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

let opened: OpenedPptx
let textId: string

async function openWith(paragraphs: Array<Array<{ text: string; bold?: boolean }>>): Promise<void> {
  opened = await openPptx(await createBlankPptx())
  const el = addElement(opened.deck.slides[0]!, {
    kind: 'textbox',
    offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
    paragraphs: paragraphs.map((runs) => ({ runs: runs.map((run) => ({ ...run })) })),
  })
  textId = el.id
}

function textElement(): TextElement {
  const el = opened.deck.slides[0]!.elements.find((item) => item.id === textId)
  if (!el || (el.type !== 'text' && el.type !== 'shape')) throw new Error('missing text element')
  return el
}

function paragraphSlices(xml: string): string[] {
  return xml.match(/<a:p>[\s\S]*?<\/a:p>/g) ?? []
}

function runSlices(paragraphXml: string): string[] {
  return paragraphXml.match(/<a:r>[\s\S]*?<\/a:r>/g) ?? []
}

const setText = (paragraphs: unknown[]) =>
  runTxn(opened, {
    ops: [{ op: 'setText', target: { slide: 0, el: textId }, paragraphs }],
  })

describe('setText paragraph/run reshaping', () => {
  it('writes 2 two-run paragraphs into a 4-one-run-paragraph box as 2 paragraphs', async () => {
    await openWith([
      [{ text: 'alpha' }],
      [{ text: 'beta' }],
      [{ text: 'gamma' }],
      [{ text: 'delta' }],
    ])
    const result = setText([
      { runs: [{ text: 'Label: ', bold: true }, { text: 'one' }] },
      { runs: [{ text: 'Note: ', bold: true }, { text: 'two' }] },
    ])
    expect(result.applied).toBe(true)
    const xml = patchSlideXml(opened.deck.slides[0]!)
    const paragraphs = paragraphSlices(xml)
    expect(paragraphs).toHaveLength(2)
    for (const paragraph of paragraphs) expect(runSlices(paragraph)).toHaveLength(2)
    expect(paragraphs[0]).toContain('<a:t>Label: </a:t>')
    expect(paragraphs[0]).toContain('<a:t>one</a:t>')
    expect(paragraphs[1]).toContain('<a:t>Note: </a:t>')
    expect(paragraphs[1]).toContain('<a:t>two</a:t>')
    const firstRuns = runSlices(paragraphs[0]!)
    expect(firstRuns[0]).toContain('b="1"')
    expect(firstRuns[1]).not.toContain('b="1"')
    const reopened = await openPptx(await savePptx(opened))
    const saved = reopened.deck.slides[0]!.elements.find(
      (item) => item.type === 'text' || item.type === 'shape',
    ) as TextElement
    expect(saved.text!.paragraphs).toHaveLength(2)
    expect(saved.text!.paragraphs.map((paragraph) => paragraph.runs.length)).toEqual([2, 2])
    expect(saved.text!.paragraphs[0]!.runs[0]!.bold).toBe(true)
    expect(saved.text!.paragraphs[0]!.runs[1]!.bold ?? false).toBe(false)
  })

  it('keeps the moved run formatting when paragraphs are reordered', async () => {
    await openWith([[{ text: 'X:', bold: true }], [{ text: 'y' }], [{ text: 'z' }]])
    const result = setText([
      { srcPara: 2, runs: [{ text: 'z', srcRun: 0 }] },
      { srcPara: 1, runs: [{ text: 'y', srcRun: 0 }] },
      { srcPara: 0, runs: [{ text: 'X:', srcRun: 0, bold: true }] },
    ])
    expect(result.applied).toBe(true)
    expect(textElement().textMoved).toBe(true)
    const xml = patchSlideXml(opened.deck.slides[0]!)
    const paragraphs = paragraphSlices(xml)
    expect(paragraphs).toHaveLength(3)
    expect(paragraphs[0]).toContain('<a:t>z</a:t>')
    expect(paragraphs[0]).not.toContain('b="1"')
    expect(paragraphs[1]).toContain('<a:t>y</a:t>')
    expect(paragraphs[1]).not.toContain('b="1"')
    expect(paragraphs[2]).toContain('<a:t>X:</a:t>')
    expect(paragraphs[2]).toContain('b="1"')
    const paragraphsModel = textElement().text!.paragraphs
    expect(paragraphsModel.map((paragraph) => paragraph.runs.map((run) => run.text))).toEqual([
      ['z'],
      ['y'],
      ['X:'],
    ])
    expect(paragraphsModel.map((paragraph) => paragraph.runs[0]!.bold ?? false)).toEqual([
      false,
      false,
      true,
    ])
    const reopened = await openPptx(await savePptx(opened))
    const saved = reopened.deck.slides[0]!.elements.find(
      (item) => item.type === 'text' || item.type === 'shape',
    ) as TextElement
    expect(saved.text!.paragraphs.map((paragraph) => paragraph.runs[0]!.text)).toEqual([
      'z',
      'y',
      'X:',
    ])
    expect(saved.text!.paragraphs.map((paragraph) => paragraph.runs[0]!.bold ?? false)).toEqual([
      false,
      false,
      true,
    ])
  })
})
