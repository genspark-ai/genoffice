/** fit text (w:fitText) and East Asian layout (w:eastAsianLayout) of a run: read, kept and written */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import type { Run } from '../src/types'
import { buildDocx } from './helpers/build-docx'

const run = (rPr: string, text: string) =>
  `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`

/** re-saves the first paragraph from its runs, as the editor does after an edit there */
async function resave(body: string, edit: (runs: Run[]) => Run[]) {
  const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
  const runs = edit((parsed.blocks[0] as { runs: Run[] }).runs)
  const saved = await saveDocx(parsed, [
    { kind: 'generated', block: { type: 'paragraph', runs }, docxIndex: 0 },
  ])
  return (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
}

const FIT = '<w:fitText w:val="1440" w:id="7"/>'
const TWO_LINES = '<w:eastAsianLayout w:id="3" w:combine="1" w:combineBrackets="round"/>'

describe('fit text and East Asian layout of a run', () => {
  it('reads what the run says about itself', async () => {
    const body =
      '<w:p>' +
      run(FIT, 'fit') +
      run(TWO_LINES, 'two') +
      run('<w:eastAsianLayout w:id="4" w:vert="1"/>', '12') +
      run('<w:b/>', 'plain') +
      '</w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const [fit, two, upright, plain] = (parsed.blocks[0] as { runs: Run[] }).runs
    expect(fit.fitTextTwips).toBe(1440)
    expect(two.eaLayout).toEqual({ combine: true, id: 3 })
    expect(upright.eaLayout).toEqual({ vert: true, id: 4 })
    expect(plain.fitTextTwips).toBeUndefined()
    expect(plain.eaLayout).toBeUndefined()
  })

  it('writes them when they are set on new text', async () => {
    const xml = await resave('<w:p><w:r><w:t>ab</w:t></w:r></w:p>', (runs) => [
      { ...runs[0], text: 'a', fitTextTwips: 2000 },
      { ...runs[0], text: 'b', eaLayout: { combine: true, id: 9 } },
    ])
    expect(xml).toContain('<w:fitText w:val="2000"/>')
    expect(xml).toContain('<w:eastAsianLayout w:id="9" w:combine="1"/>')
  })

  it('keeps the bytes of ones that were not touched, switches the model does not carry included', async () => {
    const xml = await resave('<w:p>' + run(FIT, 'fit') + run(TWO_LINES, 'two') + '</w:p>', (runs) =>
      runs.map((r) => ({ ...r, text: `${r.text}!` })),
    )
    expect(xml).toContain(FIT)
    expect(xml).toContain(TWO_LINES)
  })

  it('takes them off when they are cleared', async () => {
    const xml = await resave('<w:p>' + run(FIT, 'fit') + run(TWO_LINES, 'two') + '</w:p>', (runs) =>
      runs.map((r) => ({ ...r, fitTextTwips: 0, eaLayout: {} })),
    )
    expect(xml).not.toContain('<w:fitText')
    expect(xml).not.toContain('<w:eastAsianLayout')
  })

  it('changes the width of fitted text', async () => {
    const xml = await resave('<w:p>' + run(FIT, 'fit') + '</w:p>', (runs) =>
      runs.map((r) => ({ ...r, fitTextTwips: 2880 })),
    )
    expect(xml).toContain('<w:fitText w:val="2880"/>')
    expect(xml).not.toContain('w:val="1440"')
  })
})
