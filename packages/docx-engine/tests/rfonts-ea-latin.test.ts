import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { generateParagraphXml, parseDocx, saveDocx, type GenerateContext } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const GEN_CTX: GenerateContext = {
  headingStyleIds: new Map([[1, 'Heading1']]),
  allocateHyperlinkRel: () => 'rId999',
}

const EAST_ASIAN_FONT = 'SimSun'
const LATIN_FONT = 'Times New Roman'
const PAIRED_RFONTS =
  `<w:rFonts w:ascii="${LATIN_FONT}" w:eastAsia="${EAST_ASIAN_FONT}" ` + `w:hAnsi="${LATIN_FONT}"/>`
const PAIRED_PARA = `<w:p><w:r><w:rPr>${PAIRED_RFONTS}</w:rPr><w:t>Hello World</w:t></w:r></w:p>`

async function documentXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  const entry = zip.file('word/document.xml')
  expect(entry).not.toBeNull()
  return entry!.async('text')
}

describe('rFonts east-asian vs latin round-trip', () => {
  it('parses a paired run into distinct eastAsia and latin slots', async () => {
    expect(PAIRED_PARA).toContain(`w:eastAsia="${EAST_ASIAN_FONT}"`)
    expect(PAIRED_PARA).toContain(`w:ascii="${LATIN_FONT}"`)
    expect(PAIRED_PARA).toContain(`w:hAnsi="${LATIN_FONT}"`)

    const doc = await parseDocx(await buildDocx({ bodyXml: PAIRED_PARA }))
    const run = doc.blocks[0].runs![0]
    expect(run.font).toBe(EAST_ASIAN_FONT)
    expect(run.fontAscii).toBe(LATIN_FONT)
    expect(run.rawRPr).toContain(`w:eastAsia="${EAST_ASIAN_FONT}"`)
    expect(run.rawRPr).toContain(`w:ascii="${LATIN_FONT}"`)
    expect(run.rawRPr).toContain(`w:hAnsi="${LATIN_FONT}"`)
  })

  it('regenerating an untouched paired run keeps both attributes', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PAIRED_PARA }))
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml({ type: 'paragraph', runs: [{ ...run }] }, GEN_CTX)
    expect(xml).toContain(`w:eastAsia="${EAST_ASIAN_FONT}"`)
    expect(xml).toContain(`w:ascii="${LATIN_FONT}"`)
    expect(xml).toContain(`w:hAnsi="${LATIN_FONT}"`)
  })

  it('a saved file keeps both attributes after a full parse to generate round-trip', async () => {
    const bytes = await buildDocx({ bodyXml: PAIRED_PARA })
    const doc = await parseDocx(bytes)
    const run = doc.blocks[0].runs![0]

    const saved = await saveDocx(doc, [
      {
        kind: 'generated',
        block: { type: 'paragraph', runs: [{ ...run }] },
      },
    ])
    const documentXml = await documentXmlOf(saved)
    expect(documentXml).toContain(`w:eastAsia="${EAST_ASIAN_FONT}"`)
    expect(documentXml).toContain(`w:ascii="${LATIN_FONT}"`)
    expect(documentXml).toContain(`w:hAnsi="${LATIN_FONT}"`)

    const reparsed = await parseDocx(saved)
    const rerun = reparsed.blocks[0].runs![0]
    expect(rerun.font).toBe(EAST_ASIAN_FONT)
    expect(rerun.fontAscii).toBe(LATIN_FONT)
  })
})
