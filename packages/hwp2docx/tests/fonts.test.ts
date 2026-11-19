import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../../docx-engine/src/index'
import { classifyFont, convertHwpxToDocx } from '../src/index'

describe('classifyFont', () => {
  it.each([
    ['휴먼명조', 'serif', true],
    ['한양신명조', 'serif', true],
    ['함초롬바탕', 'serif', true],
    ['HY신명조', 'serif', true],
    ['신명 견명조', 'serif', true],
    ['한양중고딕', 'sans', true],
    ['HY헤드라인M', 'sans', true],
    ['맑은 고딕', 'sans', true],
    ['함초롬돋움', 'sans', true],
    ['한컴 윤고딕 740', 'sans', true],
    ['나눔스퀘어 ExtraBold', 'sans', true],
    ['문체부 돋음체', 'sans', true],
    ['돋움체', 'mono', true],
    ['굴림체', 'mono', true],
    ['바탕체', 'mono', true],
    ['KoPub바탕체', 'serif', true],
    ['HCI Poppy', 'sans', false],
    ['HCI Tulip', 'serif', false],
    ['Times New Roman', 'serif', false],
  ] as const)('%s → %s (hangul %s)', (name, cls, hangul) => {
    expect(classifyFont(name)).toMatchObject({ cls, hangul })
  })

  it('treats an unknown Hangul-lettered name as Hangul serif (Hangul default face)', () => {
    expect(classifyFont('새로운글꼴')).toMatchObject({ cls: 'serif', hangul: true })
  })
})

describe('font table of the converted document', () => {
  const NS =
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"' +
    ' xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"'
  const header = `<?xml version="1.0"?><hh:head ${NS}><hh:refList>
    <hh:fontfaces itemCnt="2">
      <hh:fontface lang="HANGUL"><hh:font id="0" face="휴먼명조"/><hh:font id="1" face="HY헤드라인M"/></hh:fontface>
      <hh:fontface lang="LATIN"><hh:font id="0" face="휴먼명조"/><hh:font id="1" face="HCI Poppy"/></hh:fontface>
    </hh:fontfaces>
    <hh:charProperties itemCnt="2">
      <hh:charPr id="0" height="1000"><hh:fontRef hangul="0" latin="0"/></hh:charPr>
      <hh:charPr id="1" height="2000"><hh:fontRef hangul="1" latin="1"/><hh:bold/></hh:charPr>
    </hh:charProperties>
    <hh:paraProperties itemCnt="1"><hh:paraPr id="0"><hh:align horizontal="LEFT"/></hh:paraPr></hh:paraProperties>
  </hh:refList></hh:head>`
  const body = `<?xml version="1.0"?><hs:sec ${NS}>
    <hp:p paraPrIDRef="0"><hp:run charPrIDRef="1"><hp:t>제목 Title</hp:t></hp:run></hp:p>
    <hp:p paraPrIDRef="0"><hp:run charPrIDRef="0"><hp:t>본문 text</hp:t></hp:run></hp:p></hs:sec>`

  async function convert(): Promise<JSZip> {
    const zip = new JSZip()
    zip.file('mimetype', 'application/hwp+zip')
    zip.file('Contents/header.xml', header)
    zip.file('Contents/section0.xml', body)
    const result = await convertHwpxToDocx(await zip.generateAsync({ type: 'uint8array' }))
    return JSZip.loadAsync(result.docx)
  }

  it('keeps only the original font names on the runs (downloads open in the real faces)', async () => {
    const docx = await convert()
    const document = await docx.file('word/document.xml')!.async('string')
    const names = new Set(
      [...document.matchAll(/w:(?:ascii|hAnsi|eastAsia)="([^"]+)"/g)].map((m) => m[1]),
    )
    expect([...names].sort()).toEqual(['HCI Poppy', 'HY헤드라인M', '휴먼명조'])
  })

  it('writes Word substitution hints per face, without altName', async () => {
    const docx = await convert()
    const table = await docx.file('word/fontTable.xml')!.async('string')
    expect(table).toContain(
      '<w:font w:name="휴먼명조"><w:panose1 w:val="02030600000101010101"/><w:charset w:val="81"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>',
    )
    expect(table).toMatch(
      /<w:font w:name="HY헤드라인M">.*?<w:charset w:val="81"\/><w:family w:val="swiss"\/>/,
    )
    expect(table).toMatch(/<w:font w:name="HCI Poppy">.*?<w:charset w:val="00"\/>/)
    expect(table).not.toContain('w:altName')
    // the part is wired into the package and parses back through docx-engine
    expect(await docx.file('word/_rels/document.xml.rels')!.async('string')).toContain(
      'Target="fontTable.xml"',
    )
    expect(await docx.file('[Content_Types].xml')!.async('string')).toContain('/word/fontTable.xml')
    const parsed = await parseDocx(await docx.generateAsync({ type: 'uint8array' }))
    expect(parsed.fontTable?.map((e) => e.name).sort()).toEqual([
      'HCI Poppy',
      'HY헤드라인M',
      '휴먼명조',
    ])
    expect(parsed.fontTable?.find((e) => e.name === '휴먼명조')).toMatchObject({
      charset: '81',
      family: 'roman',
    })
  })
})
