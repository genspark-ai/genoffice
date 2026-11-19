import { describe, expect, it } from 'vitest'
import { patchTextElementXml } from '../src/index'
import type { TextElement } from '../src/types'

interface RunModel {
  text: string
  bold?: boolean
}

function model(paragraphs: RunModel[][]): TextElement {
  return {
    type: 'text',
    text: { paragraphs: paragraphs.map((runs) => ({ runs })) },
  } as unknown as TextElement
}

function original(runsPerParagraph: number[], bodyPr = '<a:bodyPr wrap="none"/>'): string {
  const paragraphs = runsPerParagraph
    .map((runCount, pi) => {
      const runs = Array.from(
        { length: runCount },
        (_, ri) => '<a:r><a:rPr lang="en-US"/><a:t>old ' + pi + '.' + ri + '</a:t></a:r>',
      ).join('')
      return '<a:p>' + runs + '</a:p>'
    })
    .join('')
  return '<p:sp><p:txBody>' + bodyPr + paragraphs + '</p:txBody></p:sp>'
}

function paragraphSlices(xml: string): string[] {
  return xml.match(/<a:p>[\s\S]*?<\/a:p>/g) ?? []
}

function runSlices(paragraphXml: string): string[] {
  return paragraphXml.match(/<a:r>[\s\S]*?<\/a:r>/g) ?? []
}

function runCounts(xml: string): number[] {
  return paragraphSlices(xml).map((paragraph) => runSlices(paragraph).length)
}

describe('patchTextElementXml paragraph structure', () => {
  it('writes a 4-paragraph box reshaped to 2 paragraphs x 2 runs as 2 paragraphs', () => {
    const el = model([
      [{ text: 'Label: ', bold: true }, { text: 'one' }],
      [{ text: 'Note: ', bold: true }, { text: 'two' }],
    ])
    const out = patchTextElementXml(el, original([1, 1, 1, 1]))
    const paragraphs = paragraphSlices(out)
    expect(paragraphs).toHaveLength(2)
    expect(paragraphs.map((paragraph) => runSlices(paragraph).length)).toEqual([2, 2])
    expect(out).toContain('<a:bodyPr wrap="none"/>')
    expect(paragraphs[0]).toContain('<a:t>Label: </a:t>')
    expect(paragraphs[0]).toContain('<a:t>one</a:t>')
    expect(paragraphs[1]).toContain('<a:t>Note: </a:t>')
    expect(paragraphs[1]).toContain('<a:t>two</a:t>')
    const firstRuns = runSlices(paragraphs[0]!)
    expect(firstRuns[0]).toContain('b="1"')
    expect(firstRuns[1]).not.toContain('b="1"')
    expect(out).not.toContain('old 0.0')
  })

  it('rebuilds when the paragraph count differs at the same flattened run count', () => {
    const el = model([[{ text: 'a' }], [{ text: 'b' }], [{ text: 'c' }, { text: 'd' }]])
    const out = patchTextElementXml(el, original([1, 1, 1, 1]))
    expect(runCounts(out)).toEqual([1, 1, 2])
    expect(out).toContain('<a:t>c</a:t>')
    expect(out).toContain('<a:t>d</a:t>')
  })

  it('rebuilds when a paragraph keeps the total run count but changes its own', () => {
    const el = model([[{ text: 'a' }, { text: 'b' }], [{ text: 'c' }]])
    const out = patchTextElementXml(el, original([1, 2]))
    expect(runCounts(out)).toEqual([2, 1])
  })

  it('a textMoved flag rebuilds from the traced runs instead of the old position formatting', () => {
    const originalXml =
      '<p:sp><p:txBody><a:p><a:r><a:rPr b="1"/><a:t>old</a:t></a:r></a:p></p:txBody></p:sp>'
    const plain = model([[{ text: 'moved' }]])
    // without the flag the in-place patch keeps the old run's <a:rPr b="1"/>
    expect(patchTextElementXml(plain, originalXml)).toContain('b="1"')
    const moved = model([[{ text: 'moved' }]])
    moved.textMoved = true
    const out = patchTextElementXml(moved, originalXml)
    expect(out).toContain('<a:t>moved</a:t>')
    expect(out).not.toContain('b="1"')
    expect(paragraphSlices(out)).toHaveLength(1)
  })

  it('a same-shape text edit still takes the byte-preserving in-place path', () => {
    const originalXml =
      '<p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>Hi</a:t></a:r></a:p></p:txBody></p:sp>'
    const out = patchTextElementXml(model([[{ text: 'Bye' }]]), originalXml)
    expect(out).toContain('<a:rPr lang="en-US" sz="1800"/>')
    expect(out).toContain('<a:rPr lang="en-US" sz="1800"/><a:t>Bye</a:t></a:r>')
    expect(paragraphSlices(out)).toHaveLength(1)
  })
})
