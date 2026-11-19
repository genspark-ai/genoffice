import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  generateRubyXml,
  parseDocx,
  rubyAlignOf,
  rubyReadingStyleOf,
  saveDocx,
  setRubyAlignXml,
  setRubyBaseXml,
  setRubyReadingStyleXml,
  setRubyReadingXml,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

function rubyXml(base: string, rt: string): string {
  return (
    '<w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="12"/>' +
    '<w:hpsRaise w:val="22"/><w:hpsBaseText w:val="24"/><w:lid w:val="zh-CN"/></w:rubyPr>' +
    `<w:rt><w:r><w:rPr><w:sz w:val="12"/></w:rPr><w:t>${rt}</w:t></w:r></w:rt>` +
    `<w:rubyBase><w:r><w:t>${base}</w:t></w:r></w:rubyBase></w:ruby>`
  )
}

describe('phonetic guide (w:ruby)', () => {
  it('parses ruby runs keeping the base characters and phonetic text', async () => {
    const p =
      `<w:p><w:r>${rubyXml('床', 'chuáng')}</w:r><w:r>${rubyXml('前', 'qián')}</w:r>` +
      '<w:r><w:t>,</w:t></w:r></w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: p }))
    const block = parsed.blocks[0]
    expect(block.type).toBe('paragraph')
    expect(block.runs?.map((r) => r.text)).toEqual(['床', '前', ','])
    expect(block.runs?.[0].ruby).toEqual({ rt: 'chuáng', xml: rubyXml('床', 'chuáng') })
    expect(block.runs?.[1].ruby?.rt).toBe('qián')
    expect(block.runs?.[2].ruby).toBeUndefined()
  })

  it('an edited paragraph re-emits ruby fragments verbatim', async () => {
    const p = `<w:p><w:r>${rubyXml('月', 'yuè')}</w:r><w:r><w:t>。</w:t></w:r></w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: p }))
    const block = parsed.blocks[0]
    const bytes = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: 'edited ' }, ...block.runs!.filter((r) => r.ruby)],
        },
      },
    ])
    const reparsed = await parseDocx(bytes)
    const rb = reparsed.blocks[0]
    expect(rb.runs?.map((r) => r.text)).toEqual(['edited ', '月'])
    expect(rb.runs?.[1].ruby).toEqual({ rt: 'yuè', xml: rubyXml('月', 'yuè') })
  })

  it('an untouched ruby paragraph round-trips byte-preserving', async () => {
    const p = `<w:p><w:r>${rubyXml('霜', 'shuāng')}</w:r></w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: p }))
    const bytes = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }])
    const reparsed = await parseDocx(bytes)
    expect(reparsed.blocks[0].runs?.[0].text).toBe('霜')
    expect(reparsed.blocks[0].runs?.[0].ruby?.xml).toBe(rubyXml('霜', 'shuāng'))
  })

  it('multi-run ruby base and rt concatenate', async () => {
    const ruby =
      '<w:ruby><w:rubyPr/><w:rt><w:r><w:t>gù</w:t></w:r><w:r><w:t>xiāng</w:t></w:r></w:rt>' +
      '<w:rubyBase><w:r><w:t>故</w:t></w:r><w:r><w:t>乡</w:t></w:r></w:rubyBase></w:ruby>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: `<w:p><w:r>${ruby}</w:r></w:p>` }))
    const run = parsed.blocks[0].runs?.[0]
    expect(run?.text).toBe('故乡')
    expect(run?.ruby?.rt).toBe('gùxiāng')
  })

  it('a guide generated for new text saves and reads back as that base and reading', async () => {
    const xml = generateRubyXml('\u6708', 'yu\u00e8 <&>', 24, 'zh-CN')
    expect(xml).toContain('<w:hps w:val="12"/>')
    expect(xml).toContain('<w:hpsBaseText w:val="24"/>')
    expect(xml).toContain('<w:rubyAlign w:val="center"/>')
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>plain</w:t></w:r></w:p>' }),
    )
    const bytes = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: 'a ' }, { text: '\u6708', ruby: { rt: 'yu\u00e8 <&>', xml } }],
        },
      },
    ])
    const runs = (await parseDocx(bytes)).blocks[0].runs
    expect(runs?.map((r) => r.text)).toEqual(['a ', '\u6708'])
    expect(runs?.[1].ruby).toEqual({ rt: 'yu\u00e8 <&>', xml })
  })

  it('the formatting of the base text is saved on the base run inside the guide', async () => {
    const xml = generateRubyXml('\u6708', 'yue', 24)
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>plain</w:t></w:r></w:p>' }),
    )
    const bytes = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: '\u6708', bold: true, color: 'FF0000', ruby: { rt: 'yue', xml } }],
        },
      },
    ])
    const out = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
    expect(out).toMatch(
      /<w:rubyBase><w:r><w:rPr>.*<w:b\/>.*<w:color w:val="FF0000"\/>.*<\/w:rPr><w:t/,
    )
    // a guide whose base run says how it looks is written as it is
    const styled = xml.replace('<w:rubyBase><w:r>', '<w:rubyBase><w:r><w:rPr><w:i/></w:rPr>')
    const kept = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: '\u6708', bold: true, ruby: { rt: 'yue', xml: styled } }],
        },
      },
    ])
    const keptXml = await (await JSZip.loadAsync(kept)).file('word/document.xml')!.async('string')
    expect(keptXml).toContain(styled)
  })

  it('another reading is put into a guide without touching the rest of it', () => {
    const original = rubyXml('\u6708', 'yu\u00e8')
    const changed = setRubyReadingXml(original, 'y <u>')
    expect(changed).toContain('<w:rubyAlign w:val="distributeSpace"/>')
    expect(changed).toContain(
      '<w:rPr><w:sz w:val="12"/></w:rPr><w:t xml:space="preserve">y &lt;u&gt;</w:t>',
    )
    expect(changed).toContain('<w:rubyBase><w:r><w:t>\u6708</w:t></w:r></w:rubyBase>')
    expect(setRubyReadingXml('<w:ruby><w:rubyBase/></w:ruby>', 'x')).toBeNull()
  })

  it('reads and changes the alignment of a guide, adding one when the guide names none', () => {
    const original = rubyXml('\u6708', 'yu\u00e8')
    expect(rubyAlignOf(original)).toBe('distributeSpace')
    const left = setRubyAlignXml(original, 'left')!
    expect(rubyAlignOf(left)).toBe('left')
    expect(left).toBe(original.replace('w:val="distributeSpace"', 'w:val="left"'))
    const bare = '<w:ruby><w:rubyPr><w:hps w:val="12"/></w:rubyPr><w:rt/><w:rubyBase/></w:ruby>'
    expect(rubyAlignOf(bare)).toBeNull()
    expect(setRubyAlignXml(bare, 'right')).toContain(
      '<w:rubyPr><w:rubyAlign w:val="right"/><w:hps w:val="12"/>',
    )
    // an alignment the dialog does not offer reads as none, and a guide without properties is left alone
    expect(rubyAlignOf(original.replace('distributeSpace', 'rightVertical'))).toBeNull()
    expect(setRubyAlignXml('<w:ruby><w:rt/></w:ruby>', 'left')).toBeNull()
  })

  it('other base text is put into a guide without touching the rest of it', () => {
    const original = rubyXml('\u6708', 'yu\u00e8')
    const changed = setRubyBaseXml(original, 'a & b')!
    expect(changed).toContain(
      '<w:rubyBase><w:r><w:t xml:space="preserve">a &amp; b</w:t></w:r></w:rubyBase>',
    )
    expect(changed).toContain('<w:t>yu\u00e8</w:t>')
    expect(setRubyBaseXml('<w:ruby><w:rt/></w:ruby>', 'x')).toBeNull()
  })

  it('a new guide is written with the alignment asked for', () => {
    expect(generateRubyXml('\u6708', 'yue')).toContain('<w:rubyAlign w:val="center"/>')
    expect(generateRubyXml('\u6708', 'yue', 24, 'ja-JP', 'right')).toContain(
      '<w:rubyAlign w:val="right"/>',
    )
  })

  it('reads how the reading of a guide is set', () => {
    expect(rubyReadingStyleOf(rubyXml('a', 'b'))).toEqual({
      font: undefined,
      sizeHalfPoints: 12,
      raiseHalfPoints: 22,
      baseSizeHalfPoints: 24,
    })
    const faced = rubyXml('a', 'b').replace(
      '<w:rPr>',
      '<w:rPr><w:rFonts w:ascii="A &amp; B" w:hAnsi="A &amp; B"/>',
    )
    expect(rubyReadingStyleOf(faced).font).toBe('A & B')
  })

  it('sets the font, size and raise of a reading and leaves the base alone', () => {
    const xml = rubyXml('a', 'b')
    const set = setRubyReadingStyleXml(xml, {
      font: 'Arial',
      sizeHalfPoints: 16,
      raiseHalfPoints: 30,
    })!
    expect(set).toContain('<w:hps w:val="16"/><w:hpsRaise w:val="30"/><w:hpsBaseText w:val="24"/>')
    expect(set).toContain(
      '<w:rt><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>' +
        '<w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t>b</w:t></w:r></w:rt>',
    )
    expect(set).toContain('<w:rubyBase><w:r><w:t>a</w:t></w:r></w:rubyBase>')
    expect(rubyReadingStyleOf(set)).toEqual({
      font: 'Arial',
      sizeHalfPoints: 16,
      raiseHalfPoints: 30,
      baseSizeHalfPoints: 24,
    })
    // one property at a time: the others stay as the guide has them
    expect(setRubyReadingStyleXml(xml, { raiseHalfPoints: 26 })).toBe(
      xml.replace('<w:hpsRaise w:val="22"/>', '<w:hpsRaise w:val="26"/>'),
    )
    // an empty font takes the reading's own face off
    expect(setRubyReadingStyleXml(set, { font: '' })).not.toContain('w:rFonts')
    expect(setRubyReadingStyleXml('<w:ruby><w:rt/></w:ruby>', { sizeHalfPoints: 10 })).toBeNull()
  })

  it('a reading run without properties, and a guide without the sizes, take them', () => {
    const bare =
      '<w:ruby><w:rubyPr><w:lid w:val="ja-JP"/></w:rubyPr>' +
      '<w:rt><w:r><w:t>b</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>a</w:t></w:r></w:rubyBase></w:ruby>'
    expect(setRubyReadingStyleXml(bare, { sizeHalfPoints: 10, raiseHalfPoints: 20 })).toBe(
      '<w:ruby><w:rubyPr><w:hps w:val="10"/><w:hpsRaise w:val="20"/><w:lid w:val="ja-JP"/></w:rubyPr>' +
        '<w:rt><w:r><w:rPr><w:sz w:val="10"/><w:szCs w:val="10"/></w:rPr><w:t>b</w:t></w:r></w:rt>' +
        '<w:rubyBase><w:r><w:t>a</w:t></w:r></w:rubyBase></w:ruby>',
    )
  })
})
