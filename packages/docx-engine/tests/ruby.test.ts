import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
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
})
