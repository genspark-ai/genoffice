import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { convertHwpToTiptap, type TiptapNode } from '../src/index'

const NS =
  'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"' +
  ' xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"'

const HEADER = `<?xml version="1.0"?><hh:head ${NS}><hh:refList>
  <hh:fontfaces itemCnt="2">
    <hh:fontface lang="HANGUL"><hh:font id="0" face="휴먼명조"/></hh:fontface>
    <hh:fontface lang="LATIN"><hh:font id="0" face="휴먼명조"/></hh:fontface>
  </hh:fontfaces>
  <hh:borderFills itemCnt="1"><hh:borderFill id="1"><hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/><hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/><hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/><hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/><hc:fillBrush><hc:winBrush faceColor="#D9D9D9" hatchColor="#000000" alpha="0"/></hc:fillBrush></hh:borderFill></hh:borderFills>
  <hh:charProperties itemCnt="2">
    <hh:charPr id="0" height="1000" textColor="#000000"><hh:fontRef hangul="0" latin="0"/></hh:charPr>
    <hh:charPr id="1" height="1600" textColor="#0000FF"><hh:fontRef hangul="0" latin="0"/><hh:bold/></hh:charPr>
  </hh:charProperties>
  <hh:paraProperties itemCnt="2">
    <hh:paraPr id="0"><hh:align horizontal="JUSTIFY"/></hh:paraPr>
    <hh:paraPr id="1"><hh:align horizontal="CENTER"/></hh:paraPr>
  </hh:paraProperties>
</hh:refList></hh:head>`

const cell = (col: number, row: number, text: string, span = { col: 1, row: 1 }) =>
  `<hp:tc borderFillIDRef="1"><hp:subList><hp:p paraPrIDRef="0"><hp:run charPrIDRef="0"><hp:t>${text}</hp:t></hp:run></hp:p></hp:subList>` +
  `<hp:cellAddr colAddr="${col}" rowAddr="${row}"/><hp:cellSpan colSpan="${span.col}" rowSpan="${span.row}"/>` +
  `<hp:cellSz width="${7500 * span.col}" height="1500"/></hp:tc>`

async function hwpx(body: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('mimetype', 'application/hwp+zip')
  zip.file('Contents/header.xml', HEADER)
  zip.file('Contents/section0.xml', `<?xml version="1.0"?><hs:sec ${NS}>${body}</hs:sec>`)
  return zip.generateAsync({ type: 'uint8array' })
}

const texts = (node: TiptapNode): string =>
  node.text ?? (node.content ?? []).map(texts).join(node.type === 'doc' ? '\n' : '')

describe('convertHwpToTiptap', () => {
  it('emits Tiptap JSON from the same docx the GenOffice filter writes', async () => {
    const table =
      `<hp:tbl rowCnt="2" colCnt="2" borderFillIDRef="1"><hp:sz width="15000" height="3000"/>` +
      `<hp:tr>${cell(0, 0, '세로 병합', { col: 1, row: 2 })}${cell(1, 0, '위')}</hp:tr>` +
      `<hp:tr>${cell(1, 1, '아래')}</hp:tr></hp:tbl>`
    const result = await convertHwpToTiptap(
      await hwpx(
        `<hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:t>제목</hp:t></hp:run></hp:p>` +
          `<hp:p paraPrIDRef="0"><hp:run charPrIDRef="0"><hp:t>첫 줄<hp:lineBreak/>둘째 줄</hp:t>${table}</hp:run></hp:p>`,
      ),
    )
    const [title, body, tbl] = result.doc.content!
    expect(result.doc.type).toBe('doc')
    expect(result.docx.byteLength).toBeGreaterThan(0)
    expect(title).toMatchObject({ type: 'paragraph', attrs: { textAlign: 'center' } })
    expect(title.content![0]).toMatchObject({
      type: 'text',
      text: '제목',
      marks: [
        { type: 'bold' },
        {
          type: 'textStyle',
          attrs: {
            color: '#0000FF',
            fontFamily: '휴먼명조',
            fontSize: '16pt',
          },
        },
      ],
    })
    expect(body.attrs).toMatchObject({ textAlign: 'justify' })
    // the first section's sheet (A4) for hosts that lay out pages
    expect(result.page).toMatchObject({ widthPt: 595.3, heightPt: 841.9 })
    expect(body.content!.map((n) => n.type)).toEqual(['text', 'hardBreak', 'text'])
    // vertical merge: rowspan on the origin, no cell in the covered slot
    expect(tbl.type).toBe('table')
    const [first, second] = tbl.content!
    expect(first.content!.map((c) => [texts(c), c.attrs?.rowspan])).toEqual([
      ['세로 병합', 2],
      ['위', 1],
    ])
    expect(second.content!.map(texts)).toEqual(['아래'])
    expect(first.content![0].attrs).toMatchObject({
      colspan: 1,
      colwidth: [100],
      backgroundColor: '#D9D9D9',
    })
    // form lines travel as inline CSS for hosts that declare a style attribute
    expect(first.content![0].attrs!.style).toMatch(/^border-top:1px solid #000000;border-right:/)
  })
})
