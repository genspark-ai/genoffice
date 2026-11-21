/** character borders (w:bdr): exam answer boxes drawn around runs */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '../src/index'
import type { Run } from '../src/types'
import { buildDocx } from './helpers/build-docx'

const run = (rPr: string, text: string) =>
  `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`

describe('run border parse', () => {
  it('reads style, width, color and text gap of w:bdr', async () => {
    const body =
      '<w:p>' +
      run('<w:bdr w:val="single" w:sz="6" w:space="1" w:color="000000"/>', ' (28) ') +
      run('<w:bdr w:val="dashed" w:sz="12" w:space="0" w:color="auto"/>', 'x') +
      run('<w:bdr w:val="none" w:sz="0" w:space="0" w:color="auto"/>', 'plain') +
      run('<w:b/>', 'bold') +
      '</w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const para = doc.blocks[0]
    expect(para.type).toBe('paragraph')
    const runs = (para as { runs: Array<{ bdr?: unknown }> }).runs
    expect(runs[0].bdr).toEqual({ val: 'single', sz: 6, color: '000000', space: 1 })
    expect(runs[1].bdr).toEqual({ val: 'dashed', sz: 12 })
    expect(runs[2].bdr).toBeUndefined()
    expect(runs[3].bdr).toBeUndefined()
  })

  it('inherits w:bdr from character and paragraph styles unless the run sets its own', async () => {
    const box = '<w:bdr w:val="single" w:sz="8" w:space="2" w:color="FF0000"/>'
    const extraStylesXml =
      '<w:style w:type="character" w:styleId="Boxed"><w:name w:val="Boxed"/><w:rPr>' +
      box +
      '</w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Framed"><w:name w:val="Framed"/><w:basedOn w:val="Normal"/>' +
      '<w:rPr><w:bdr w:val="dotted" w:sz="4" w:space="0" w:color="auto"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Framed2"><w:name w:val="Framed 2"/><w:basedOn w:val="Framed"/>' +
      '<w:rPr><w:b/></w:rPr></w:style>'
    const body =
      '<w:p>' +
      run('<w:rStyle w:val="Boxed"/>', 'styled') +
      run(
        '<w:rStyle w:val="Boxed"/><w:bdr w:val="none" w:sz="0" w:space="0" w:color="auto"/>',
        'off',
      ) +
      run('<w:rStyle w:val="Boxed"/><w:bdr w:val="dashed" w:sz="12"/>', 'own') +
      run('', 'plain') +
      '</w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Framed2"/></w:pPr>' +
      run('<w:i/>', 'para style') +
      '<w:r><w:t>no rPr</w:t></w:r>' +
      run('<w:rStyle w:val="Boxed"/>', 'char wins') +
      '</w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: body, extraStylesXml }))
    const runsOf = (i: number) => (doc.blocks[i] as { runs: Array<{ bdr?: unknown }> }).runs
    const [styled, off, own, plain] = runsOf(0)
    expect(styled.bdr).toEqual({ val: 'single', sz: 8, color: 'FF0000', space: 2 })
    expect(off.bdr).toBeUndefined()
    expect(own.bdr).toEqual({ val: 'dashed', sz: 12 })
    expect(plain.bdr).toBeUndefined()
    const [inherited, bare, charWins] = runsOf(1)
    expect(inherited.bdr).toEqual({ val: 'dotted', sz: 4 })
    expect(bare.bdr).toEqual({ val: 'dotted', sz: 4 })
    expect(charWins.bdr).toEqual({ val: 'single', sz: 8, color: 'FF0000', space: 2 })
  })
})

describe('run border save', () => {
  const OWN = '<w:bdr w:val="single" w:sz="6" w:space="1" w:color="000000"/>'
  const BOXED_STYLE =
    '<w:style w:type="character" w:styleId="Boxed"><w:name w:val="Boxed"/><w:rPr>' +
    '<w:bdr w:val="single" w:sz="8" w:space="2" w:color="FF0000"/></w:rPr></w:style>'

  /** re-saves the first paragraph from its runs, as the editor does after an edit there */
  async function resave(body: string, edit: (runs: Run[]) => Run[], extraStylesXml = '') {
    const parsed = await parseDocx(await buildDocx({ bodyXml: body, extraStylesXml }))
    const runs = edit((parsed.blocks[0] as { runs: Run[] }).runs)
    const saved = await saveDocx(parsed, [
      { kind: 'generated', block: { type: 'paragraph', runs }, docxIndex: 0 },
    ])
    return (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
  }

  it("reads the run's own border apart from one a style gives", async () => {
    const body = '<w:p>' + run(OWN, 'own') + run('<w:rStyle w:val="Boxed"/>', 'styled') + '</w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: body, extraStylesXml: BOXED_STYLE }))
    const [own, styled] = (parsed.blocks[0] as { runs: Run[] }).runs
    expect(own.bdrOwn).toEqual({ val: 'single', sz: 6, color: '000000', space: 1 })
    expect(styled.bdr).toBeDefined()
    expect(styled.bdrOwn).toBeUndefined()
  })

  it('writes a border set on new text', async () => {
    const xml = await resave('<w:p><w:r><w:t>plain</w:t></w:r></w:p>', (runs) => [
      { ...runs[0], bdrOwn: { val: 'single', sz: 4 } },
    ])
    expect(xml).toContain('<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>')
  })

  it('keeps the bytes of a border that was not touched', async () => {
    const xml = await resave('<w:p>' + run(OWN + '<w:b/>', 'boxed') + '</w:p>', (runs) =>
      runs.map((r) => ({ ...r, text: 'boxed and edited' })),
    )
    expect(xml).toContain(OWN)
    expect(xml.match(/<w:bdr/g)).toHaveLength(1)
  })

  it('writes an explicit no-border when the border is taken off', async () => {
    const xml = await resave('<w:p>' + run(OWN, 'boxed') + '</w:p>', (runs) =>
      runs.map((r) => ({ ...r, bdrOwn: null })),
    )
    expect(xml).toContain('<w:bdr w:val="nil"/>')
    expect(xml).not.toContain('w:sz="6"')
  })

  it("does not copy a style's border onto the run", async () => {
    const xml = await resave(
      '<w:p>' + run('<w:rStyle w:val="Boxed"/>', 'styled') + '</w:p>',
      (runs) => runs.map((r) => ({ ...r, text: 'styled and edited' })),
      BOXED_STYLE,
    )
    expect(xml).not.toContain('<w:bdr')
  })
})
