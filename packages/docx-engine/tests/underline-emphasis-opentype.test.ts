import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type Run } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const documentXml = async (bytes: Uint8Array): Promise<string> =>
  (await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) ?? ''

const para = (rPr: string, text = 'text') =>
  `<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r></w:p>`

async function parsedRun(rPr: string): Promise<Run> {
  return (await parseDocx(await buildDocx({ bodyXml: para(rPr) }))).blocks[0].runs![0]
}

async function saved(runs: Run[], bodyXml = para('')) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const bytes = await saveDocx(parsed, [{ kind: 'generated', block: { type: 'paragraph', runs } }])
  return { xml: await documentXml(bytes), runs: (await parseDocx(bytes)).blocks[0].runs ?? [] }
}

const STYLES = [
  'words',
  'double',
  'thick',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashedHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dashDotHeavy',
  'dotDotDash',
  'dashDotDotHeavy',
  'wave',
  'wavyHeavy',
  'wavyDouble',
]

describe('underline style and colour', () => {
  it('reads a plain underline as before: no style, no colour', async () => {
    const run = await parsedRun('<w:u w:val="single"/>')
    expect(run.underline).toBe(true)
    expect(run.underlineStyle).toBeUndefined()
    expect(run.underlineColor).toBeUndefined()
  })

  it.each(STYLES)(
    'reads %s, keeps its bytes untouched, and writes it on a new run',
    async (val) => {
      const rPr = `<w:u w:val="${val}" w:color="FF0000" w:themeColor="accent2"/>`
      const run = await parsedRun(rPr)
      expect(run).toMatchObject({ underline: true, underlineStyle: val, underlineColor: 'FF0000' })

      // the theme link is not in the model: an untouched run keeps the element whole
      const same = await saved([run], para(rPr))
      expect(same.xml).toContain(rPr)

      const fresh = await saved([{ text: 'new', underline: true, underlineStyle: val }])
      expect(fresh.xml).toContain(`<w:u w:val="${val}"/>`)
      expect(fresh.runs[0]).toMatchObject({ underline: true, underlineStyle: val })
      expect(fresh.runs[0].underlineColor).toBeUndefined()
    },
  )

  it('writes a single underline exactly as before', async () => {
    const { xml } = await saved([{ text: 'new', underline: true }])
    expect(xml).toContain('<w:rPr><w:u w:val="single"/></w:rPr>')
    const explicit = await saved([
      { text: 'new', underline: true, underlineStyle: 'single', underlineColor: 'auto' },
    ])
    expect(explicit.xml).toContain('<w:rPr><w:u w:val="single"/></w:rPr>')
  })

  it('keeps a double underline when the model only says underlined', async () => {
    const rPr = '<w:u w:val="double" w:color="00B050"/>'
    const run = await parsedRun(rPr)
    const { xml } = await saved(
      [{ text: 'text', rawRPr: run.rawRPr, underline: true, bold: true }],
      para(rPr),
    )
    expect(xml).toContain(rPr)
    expect(xml).toContain('<w:b/>')
  })

  it('rebuilds the element when the style or the colour changes', async () => {
    const rPr = '<w:b/><w:u w:val="double" w:color="00B050"/><w:em w:val="dot"/>'
    const run = await parsedRun(rPr)
    const restyled = await saved([{ ...run, underlineStyle: 'wave' }], para(rPr))
    expect(restyled.xml).toContain('<w:b/><w:u w:val="wave" w:color="00B050"/><w:em w:val="dot"/>')

    const single = await saved(
      [{ ...run, underlineStyle: 'single', underlineColor: 'auto' }],
      para(rPr),
    )
    expect(single.xml).toContain('<w:b/><w:u w:val="single"/><w:em w:val="dot"/>')
    expect(single.runs[0].underlineStyle).toBeUndefined()

    const recolored = await saved([{ ...run, underlineColor: '0070C0' }], para(rPr))
    expect(recolored.xml).toContain('<w:u w:val="double" w:color="0070C0"/>')
    expect(recolored.runs[0]).toMatchObject({ underlineStyle: 'double', underlineColor: '0070C0' })

    const off = await saved([{ ...run, underline: false }], para(rPr))
    expect(off.xml).not.toContain('<w:u ')
  })

  it('reads an explicit auto colour as none and leaves it in the file', async () => {
    const rPr = '<w:u w:val="dotted" w:color="auto"/>'
    const run = await parsedRun(rPr)
    expect(run.underlineColor).toBeUndefined()
    const same = await saved([{ ...run, underlineColor: 'auto' }], para(rPr))
    expect(same.xml).toContain(rPr)
  })

  it('carries the style and colour of an underlining style', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:pPr><w:pStyle w:val="Wavy"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
        extraStylesXml:
          '<w:style w:type="paragraph" w:styleId="Wavy"><w:name w:val="Wavy"/>' +
          '<w:rPr><w:u w:val="wave" w:color="C00000"/></w:rPr></w:style>',
      }),
    )
    expect(parsed.styles.get('Wavy')?.display).toMatchObject({
      underline: true,
      underlineStyle: 'wave',
      underlineColor: 'C00000',
    })
  })
})

describe('emphasis mark', () => {
  it.each(['dot', 'comma', 'circle', 'underDot'] as const)(
    'writes %s and reads it back',
    async (em) => {
      const { xml, runs } = await saved([{ text: 'new', bold: true, em }])
      expect(xml).toContain(`<w:rPr><w:b/><w:bCs/><w:em w:val="${em}"/></w:rPr>`)
      expect(runs[0].em).toBe(em)
    },
  )

  it('changes and removes the mark of a run from a file, and leaves an untouched one alone', async () => {
    const rPr = '<w:i/><w:em w:val="comma"/><w:lang w:eastAsia="ja-JP"/>'
    const run = await parsedRun(rPr)
    expect(run.em).toBe('comma')

    const same = await saved([run], para(rPr))
    expect(same.xml).toContain(rPr)
    // a rebuilt run that lost the field keeps the file's mark
    const { em: _em, ...without } = run
    expect((await saved([without], para(rPr))).xml).toContain(rPr)

    const changed = await saved([{ ...run, em: 'circle' }], para(rPr))
    expect(changed.xml).toContain('<w:i/><w:em w:val="circle"/><w:lang w:eastAsia="ja-JP"/>')

    const removed = await saved([{ ...run, em: 'none' }], para(rPr))
    expect(removed.xml).not.toContain('<w:em')
    expect(removed.xml).toContain('<w:i/><w:lang w:eastAsia="ja-JP"/>')
    expect(removed.runs[0].em).toBeUndefined()
  })
})

describe('OpenType features (w14)', () => {
  const ALL =
    '<w14:ligatures w14:val="standardContextual"/><w14:numForm w14:val="oldStyle"/>' +
    '<w14:numSpacing w14:val="tabular"/>' +
    '<w14:stylisticSets><w14:styleSet w14:id="3"/></w14:stylisticSets><w14:cntxtAlts/>'

  it('writes the five features in schema order with the namespace, and reads them back', async () => {
    const otf = {
      ligatures: 'standardContextual',
      numForm: 'oldStyle' as const,
      numSpacing: 'tabular' as const,
      stylisticSets: [3],
      cntxtAlts: true,
    }
    const { xml, runs } = await saved([{ text: 'new', bold: true, otf }])
    expect(xml).toMatch(/<w:document[^>]*xmlns:w14=/)
    expect(xml).toContain(`<w:rPr><w:b/><w:bCs/>${ALL}</w:rPr>`)
    expect(runs[0].otf).toEqual(otf)
  })

  it('reads only the elements a run has', async () => {
    expect((await parsedRun('<w:b/>')).otf).toBeUndefined()
    expect((await parsedRun('<w14:numForm w14:val="lining"/>')).otf).toEqual({ numForm: 'lining' })
    expect((await parsedRun('<w14:cntxtAlts w14:val="0"/>')).otf).toEqual({ cntxtAlts: false })
    expect(
      (
        await parsedRun(
          '<w14:stylisticSets><w14:styleSet w14:id="2"/><w14:styleSet w14:id="5" w14:val="0"/>' +
            '</w14:stylisticSets>',
        )
      ).otf,
    ).toEqual({ stylisticSets: [2] })
  })

  it('changes one feature and keeps the bytes of the others; defaults remove the element', async () => {
    const run = await parsedRun(ALL)
    const same = await saved([run], para(ALL))
    expect(same.xml).toContain(`<w:rPr>${ALL}</w:rPr>`)

    const { otf: _otf, ...without } = run
    expect((await saved([{ ...without, bold: true }], para(ALL))).xml).toContain(ALL)

    const changed = await saved(
      [{ ...run, otf: { ...run.otf, ligatures: 'all', stylisticSets: [7] } }],
      para(ALL),
    )
    expect(changed.xml).toContain(
      '<w:rPr><w14:ligatures w14:val="all"/><w14:numForm w14:val="oldStyle"/>' +
        '<w14:numSpacing w14:val="tabular"/>' +
        '<w14:stylisticSets><w14:styleSet w14:id="7"/></w14:stylisticSets><w14:cntxtAlts/></w:rPr>',
    )

    const cleared = await saved(
      [
        {
          ...run,
          otf: {
            ligatures: 'none',
            numForm: 'default',
            numSpacing: 'default',
            stylisticSets: [],
            cntxtAlts: false,
          },
        },
      ],
      para(ALL),
    )
    expect(cleared.xml).not.toContain('w14:ligatures')
    expect(cleared.xml).not.toContain('w14:numForm')
    expect(cleared.xml).not.toContain('w14:numSpacing')
    expect(cleared.xml).not.toContain('w14:stylisticSets')
    expect(cleared.xml).not.toContain('w14:cntxtAlts')
    expect(cleared.runs[0].otf).toBeUndefined()
  })
})
