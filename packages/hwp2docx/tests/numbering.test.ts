import { describe, expect, it } from 'vitest'
import { formatNumber, HeadingCounter, readHeadings } from '../src/numbering'
import { parseXml } from '../src/xml'

const HEADER = `<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"><hh:refList>
  <hh:bullets itemCnt="1"><hh:bullet id="2" char="-"><hh:paraHead level="0"/></hh:bullet></hh:bullets>
  <hh:numberings itemCnt="1"><hh:numbering id="1" start="0">
    <hh:paraHead start="1" level="1" numFormat="DIGIT">^1.</hh:paraHead>
    <hh:paraHead start="1" level="2" numFormat="HANGUL_SYLLABLE">^2.</hh:paraHead>
    <hh:paraHead start="1" level="3" numFormat="CIRCLED_DIGIT">^3</hh:paraHead>
  </hh:numbering><hh:numbering id="2" start="1">
    <hh:paraHead start="1" level="1" autoIndent="1" textOffsetType="PERCENT" textOffset="50">^  □</hh:paraHead>
  </hh:numbering></hh:numberings>
</hh:refList></hh:head>`

describe('paragraph heads', () => {
  it('writes bullets and running numbers the way Hangul shows them', () => {
    const counter = new HeadingCounter(readHeadings(parseXml(HEADER)!))
    expect(counter.next('bullet', '2', 0)?.text).toBe('-')
    const seq = [0, 1, 1, 2, 0, 1].map((level) => counter.next('number', '1', level)?.text)
    // a shallower number restarts the deeper levels
    expect(seq).toEqual(['1.', '가.', '나.', '①', '2.', '가.'])
    // a caret without a level shows nothing; the head keeps its indent rule
    expect(counter.next('number', '2', 0)).toEqual({
      text: '  □',
      autoIndent: true,
      textOffset: { percent: true, value: 50 },
    })
    expect(counter.next('outline', '9', 0)).toBeUndefined()
  })

  it('formats Hangul number styles', () => {
    expect(
      ['DIGIT', 'HANGUL_JAMO', 'ROMAN_SMALL', 'LATIN_CAPITAL', 'CIRCLED_HANGUL_SYLLABLE'].map((f) =>
        formatNumber(4, f),
      ),
    ).toEqual(['4', 'ㄹ', 'iv', 'D', '㉱'])
  })
})
