import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { eqLayoutOf, parseDocx, saveDocx, type Run } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const code = (text: string, rPr = '') =>
  `<w:r>${rPr}<w:instrText xml:space="preserve">${text}</w:instrText></w:r>`

/** Enclose Characters as Word writes it with the enclosure enlarged */
const ENCLOSED =
  begin +
  code(' eq \\o\\ac(') +
  code('○', '<w:rPr><w:position w:val="-4"/><w:sz w:val="31"/></w:rPr>') +
  code(',\u5b57)') +
  end

/** Combine Characters as Word writes it */
const COMBINED =
  begin +
  code(' eq \\o(\\s\\up 5(') +
  code('\u5408\u5e76', '<w:rPr><w:sz w:val="10"/></w:rPr>') +
  code('),\\s\\do 2(') +
  code('\u5b57\u7b26', '<w:rPr><w:sz w:val="10"/></w:rPr>') +
  code('))') +
  end

const documentXml = async (bytes: Uint8Array): Promise<string> =>
  (await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) ?? ''

async function savedRuns(runs: Run[]): Promise<{ xml: string; runs: Run[] }> {
  const parsed = await parseDocx(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>plain</w:t></w:r></w:p>' }),
  )
  const bytes = await saveDocx(parsed, [{ kind: 'generated', block: { type: 'paragraph', runs } }])
  return { xml: await documentXml(bytes), runs: (await parseDocx(bytes)).blocks[0].runs ?? [] }
}

describe('enclosed and combined characters (EQ \\o fields)', () => {
  it('tells the two fields Word writes from any other EQ field', () => {
    expect(eqLayoutOf(' eq \\o\\ac(○,\u5b57)')).toEqual({
      kind: 'enclose',
      a: '○',
      b: '\u5b57',
    })
    expect(eqLayoutOf('EQ \\o \\ac (□,12)')).toEqual({ kind: 'enclose', a: '□', b: '12' })
    expect(eqLayoutOf(' eq \\o(\\s\\up 5(ab),\\s\\do 2(cd))')).toEqual({
      kind: 'combine',
      a: 'ab',
      b: 'cd',
    })
    // an overlay of anything else, a fraction, more than two characters inside
    expect(eqLayoutOf(' eq \\o\\ac(x,y)')).toBeNull()
    expect(eqLayoutOf(' eq \\f(1,2)')).toBeNull()
    expect(eqLayoutOf(' eq \\o\\ac(○,abc)')).toBeNull()
  })

  it('reads the fields as one run each, with their parts and their own markup', async () => {
    const p = `<w:p><w:r><w:t>a</w:t></w:r>${ENCLOSED}<w:r><w:t>b</w:t></w:r>${COMBINED}</w:p>`
    const runs = (await parseDocx(await buildDocx({ bodyXml: p }))).blocks[0].runs!
    expect(runs.map((r) => r.text)).toEqual(['a', '\u5b57', 'b', '\u5408\u5e76\u5b57\u7b26'])
    expect(runs[1].eqLayout).toMatchObject({ kind: 'enclose', a: '○', b: '\u5b57' })
    expect(runs[3].eqLayout).toMatchObject({
      kind: 'combine',
      a: '\u5408\u5e76',
      b: '\u5b57\u7b26',
    })
    // the sizes Word chose stay in the kept markup
    expect(runs[1].eqLayout!.xml).toContain('<w:sz w:val="31"/>')
    expect(runs[1].math).toBeUndefined()
  })

  it('writes a field read from the file back as it was when its paragraph is edited', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: `<w:p>${ENCLOSED}</w:p>` }))
    const kept = parsed.blocks[0].runs![0]
    const { xml, runs } = await savedRuns([{ text: 'edited ' }, kept])
    expect(xml).toContain(kept.eqLayout!.xml!)
    expect(runs[1].eqLayout).toMatchObject({ kind: 'enclose', a: '○', b: '\u5b57' })
  })

  it('writes new enclosed characters in the run formatting, sized by the style', async () => {
    const run = (style: 'none' | 'shrink' | 'enlarge'): Run => ({
      text: 'A',
      bold: true,
      sizeHalfPoints: 24,
      eqLayout: { kind: 'enclose', a: '○', b: 'A', style },
    })
    const none = await savedRuns([run('none')])
    expect(none.xml).toContain('<w:instrText xml:space="preserve"> eq \\o\\ac(</w:instrText>')
    expect(none.xml).not.toContain('w:position')
    expect(none.runs[0].eqLayout).toMatchObject({ kind: 'enclose', a: '○', b: 'A' })
    expect(none.runs[0].bold).toBe(true)
    expect(none.runs[0].sizeHalfPoints).toBe(24)

    // the enclosure grows and drops; the text keeps its size
    const enlarge = await savedRuns([run('enlarge')])
    expect(enlarge.xml).toMatch(
      /<w:position w:val="-5"\/><w:sz w:val="36"\/>(?:<w:szCs w:val="36"\/>)?<\/w:rPr><w:instrText xml:space="preserve">○</,
    )
    // the text shrinks and rises; the enclosure keeps its size
    const shrink = await savedRuns([run('shrink')])
    expect(shrink.xml).toMatch(
      /<w:position w:val="2"\/><w:sz w:val="16"\/>(?:<w:szCs w:val="16"\/>)?<\/w:rPr><w:instrText xml:space="preserve">A</,
    )
    expect(shrink.runs[0].eqLayout).toMatchObject({ kind: 'enclose', a: '○', b: 'A' })
  })

  it('reads the fit back from a field it wrote, whatever size the text has', async () => {
    for (const style of ['shrink', 'enlarge'] as const) {
      for (const sizeHalfPoints of [undefined, 24]) {
        const { runs } = await savedRuns([
          {
            text: 'A',
            ...(sizeHalfPoints ? { sizeHalfPoints } : {}),
            eqLayout: { kind: 'enclose', a: '○', b: 'A', style },
          },
        ])
        expect(runs[0].eqLayout?.style, `${style} at ${sizeHalfPoints}`).toBe(style)
      }
    }
    // and from the field Word writes, where only the enclosure carries a size
    const parsed = await parseDocx(await buildDocx({ bodyXml: `<w:p>${ENCLOSED}</w:p>` }))
    expect(parsed.blocks[0].runs![0].eqLayout?.style).toBe('enlarge')
  })

  it('writes new combined characters as two half-size lines', async () => {
    const { xml, runs } = await savedRuns([
      { text: 'x' },
      {
        text: 'abcdef',
        sizeHalfPoints: 24,
        eqLayout: { kind: 'combine', a: 'abc', b: 'def' },
      },
    ])
    expect(xml).toContain(' eq \\o(\\s\\up 6(')
    expect(xml).toContain('),\\s\\do 2(')
    expect(xml.match(/<w:sz w:val="12"\/>/g)).toHaveLength(2)
    expect(runs.map((r) => r.text)).toEqual(['x', 'abcdef'])
    expect(runs[1].eqLayout).toMatchObject({ kind: 'combine', a: 'abc', b: 'def' })
  })
})
