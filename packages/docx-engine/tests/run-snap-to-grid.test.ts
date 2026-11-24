/**
 * Run-level w:snapToGrid (Format > Font > Advanced > Snap to grid): the run's
 * own value saves, the value resolved through its styles tells the page
 * whether the run is on the character grid.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { mergeRPrModel } from '../src/generate'
import { buildDocx } from './helpers/build-docx'

const STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:styleId="OffPara"><w:name w:val="OffPara"/>' +
  '<w:rPr><w:snapToGrid w:val="0"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="OffChar"><w:name w:val="OffChar"/>' +
  '<w:rPr><w:snapToGrid w:val="0"/></w:rPr></w:style>' +
  '</w:styles>'

const OFF = '<w:rPr><w:snapToGrid w:val="0"/></w:rPr>'
const run = (rPr: string, text: string) => `<w:r>${rPr}<w:t>${text}</w:t></w:r>`

const runsOf = async (bodyXml: string) =>
  (await parseDocx(await buildDocx({ bodyXml, stylesXml: STYLES }))).blocks[0].runs!

const state = (runs: Awaited<ReturnType<typeof runsOf>>) =>
  runs.map((r) => [r.text, r.offGrid ?? false, r.snapToGridOwn ?? null])

describe('run-level snap to grid', () => {
  it('reads the run value, on and off', async () => {
    const runs = await runsOf(
      `<w:p>${run('', 'on')}${run(OFF, 'off')}` +
        `${run('<w:rPr><w:snapToGrid/></w:rPr>', 'stated')}</w:p>`,
    )
    expect(state(runs)).toEqual([
      ['on', false, null],
      ['off', true, false],
      ['stated', false, true],
    ])
  })

  it('a character style and a paragraph style take their runs off; the run value wins', async () => {
    const byChar = await runsOf(
      `<w:p>${run('<w:rPr><w:rStyle w:val="OffChar"/></w:rPr>', 'styled')}` +
        `${run('<w:rPr><w:rStyle w:val="OffChar"/><w:snapToGrid/></w:rPr>', 'back')}</w:p>`,
    )
    expect(state(byChar)).toEqual([
      ['styled', true, null],
      ['back', false, true],
    ])
    const byPara = await runsOf(
      `<w:p><w:pPr><w:pStyle w:val="OffPara"/></w:pPr>${run('', 'inherited')}</w:p>`,
    )
    expect(state(byPara)).toEqual([['inherited', true, null]])
  })

  it('the paragraph property of the same name is not the run property', async () => {
    const runs = await runsOf(
      `<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr>${run('', 'line grid only')}</w:p>`,
    )
    expect(state(runs)).toEqual([['line grid only', false, null]])
  })

  it('keeps an untouched document byte-identical', async () => {
    const bytes = await buildDocx({ bodyXml: `<w:p>${run(OFF, 'off')}${run('', 'on')}</w:p>` })
    const doc = await parseDocx(bytes)
    const blocks = doc.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    expect(await saveDocx(doc, blocks)).toEqual(bytes)
  })

  it('writes the value in schema order and rebuilds it only when it changed', async () => {
    const RAW = '<w:rPr><w:b/><w:snapToGrid w:val="0"/><w:color w:val="FF0000"/></w:rPr>'
    const [off] = await runsOf(`<w:p>${run(RAW, 'off')}</w:p>`)
    expect(mergeRPrModel(off.rawRPr!, off, false)).toBe(RAW)
    // a run rebuilt without the field keeps what the file had
    expect(mergeRPrModel(off.rawRPr!, { ...off, snapToGridOwn: undefined }, false)).toBe(RAW)
    expect(mergeRPrModel(off.rawRPr!, { ...off, snapToGridOwn: true }, false)).toBe(
      '<w:rPr><w:b/><w:snapToGrid/><w:color w:val="FF0000"/></w:rPr>',
    )
    const [plain] = await runsOf(
      `<w:p>${run('<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>', 'plain')}</w:p>`,
    )
    expect(mergeRPrModel(plain.rawRPr!, { ...plain, snapToGridOwn: false }, false)).toBe(RAW)
  })
})
