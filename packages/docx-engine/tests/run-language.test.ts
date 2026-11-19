/** the proofing language of a run (its own w:lang, by slot): read, kept and written */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import type { Run } from '../src/types'
import { buildDocx } from './helpers/build-docx'

const run = (rPr: string, text: string) =>
  `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`

async function resave(body: string, edit: (runs: Run[]) => Run[]) {
  const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
  const runs = edit((parsed.blocks[0] as { runs: Run[] }).runs)
  const saved = await saveDocx(parsed, [
    { kind: 'generated', block: { type: 'paragraph', runs }, docxIndex: 0 },
  ])
  return (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
}

const MIXED = '<w:lang w:val="en-GB" w:eastAsia="ja-JP"/>'

describe('the proofing language of a run', () => {
  it('reads each slot the run sets, and nothing from a run that sets none', async () => {
    const body =
      '<w:p>' +
      run(MIXED, 'a') +
      run('<w:lang w:bidi="ar-SA"/>', 'b') +
      run('<w:b/>', 'c') +
      '</w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const [mixed, bidi, plain] = (parsed.blocks[0] as { runs: Run[] }).runs
    expect(mixed.lang).toEqual({ val: 'en-GB', eastAsia: 'ja-JP' })
    expect(bidi.lang).toEqual({ bidi: 'ar-SA' })
    expect(plain.lang).toBeUndefined()
  })

  it('keeps the bytes of an untouched one', async () => {
    const xml = await resave('<w:p>' + run(MIXED, 'a') + '</w:p>', (runs) =>
      runs.map((r) => ({ ...r, text: `${r.text}!` })),
    )
    expect(xml).toContain(MIXED)
  })

  it('writes a changed slot and keeps the others', async () => {
    const xml = await resave('<w:p>' + run(MIXED, 'a') + '</w:p>', (runs) =>
      runs.map((r) => ({ ...r, lang: { ...r.lang, val: 'fr-FR' } })),
    )
    expect(xml).toContain('<w:lang w:val="fr-FR" w:eastAsia="ja-JP"/>')
    expect(xml).not.toContain('en-GB')
  })

  it('sets it on text that had none, and takes it off when cleared', async () => {
    const set = await resave('<w:p><w:r><w:t>a</w:t></w:r></w:p>', (runs) =>
      runs.map((r) => ({ ...r, lang: { eastAsia: 'zh-CN' } })),
    )
    expect(set).toContain('<w:lang w:eastAsia="zh-CN"/>')
    const cleared = await resave('<w:p>' + run(MIXED, 'a') + '</w:p>', (runs) =>
      runs.map((r) => ({ ...r, lang: {} })),
    )
    expect(cleared).not.toContain('<w:lang')
  })

  it('reads, keeps, sets and clears the mark that text is not to be proofed', async () => {
    const body = '<w:p>' + run('<w:noProof/>', 'a') + run('<w:b/>', 'b') + '</w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const [off, plain] = (parsed.blocks[0] as { runs: Run[] }).runs
    expect(off.noProof).toBe(true)
    expect(plain.noProof).toBeUndefined()
    const kept = await resave(body, (runs) => runs.map((r) => ({ ...r, text: `${r.text}!` })))
    expect(kept.match(/<w:noProof\/>/g)).toHaveLength(1)
    const set = await resave(body, (runs) => runs.map((r) => ({ ...r, noProof: true })))
    expect(set.match(/<w:noProof\/>/g)).toHaveLength(2)
    const cleared = await resave(body, (runs) => runs.map((r) => ({ ...r, noProof: false })))
    expect(cleared).not.toContain('w:noProof')
  })
})
