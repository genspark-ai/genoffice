import { describe, expect, it } from 'vitest'
import { fieldNumberText, isFieldNumberFormat } from '../src/field-number-format'
import { SeqCounter, parseSeqInstruction } from '../src/seq-fields'

/** the characters of a result as code points, so the source stays ASCII */
const points = (text: string | null): number[] | null =>
  text === null ? null : [...text].map((c) => c.codePointAt(0)!)

describe('field number formats (ECMA-376 Part 1, 17.16.4.3)', () => {
  it('writes the examples the standard gives for the East Asian formats', () => {
    // 10 \* CHINESENUM1 -> U+5341
    expect(points(fieldNumberText(10, 'CHINESENUM1'))).toEqual([0x5341])
    // 123 \* CHINESENUM2 -> the legal numerals for 1, hundred, 2, ten, 3
    expect(points(fieldNumberText(123, 'CHINESENUM2'))).toEqual([
      0x58f9, 0x4f70, 0x8d30, 0x62fe, 0x53c1,
    ])
    // CHINESENUM3 -> one hundred two ten three
    expect(points(fieldNumberText(123, 'CHINESENUM3'))).toEqual([
      0x4e00, 0x767e, 0x4e8c, 0x5341, 0x4e09,
    ])
    // 12 \* DBNUM1, KANJINUM1, DBNUM4 -> the digits one, two
    for (const format of ['DBNUM1', 'KANJINUM1', 'DBNUM4'])
      expect(points(fieldNumberText(12, format))).toEqual([0x4e00, 0x4e8c])
    // 12 \* KANJINUM2 -> ten two
    expect(points(fieldNumberText(12, 'KANJINUM2'))).toEqual([0x5341, 0x4e8c])
    // 1 \* ZODIAC1, ZODIAC2, ZODIAC3
    expect(points(fieldNumberText(1, 'ZODIAC1'))).toEqual([0x7532])
    expect(points(fieldNumberText(1, 'ZODIAC2'))).toEqual([0x5b50])
    expect(points(fieldNumberText(1, 'ZODIAC3'))).toEqual([0x7532, 0x5b50])
    // the pairs go on as the standard lists them: the second and the third
    expect(points(fieldNumberText(2, 'ZODIAC3'))).toEqual([0x4e59, 0x4e11])
    expect(points(fieldNumberText(3, 'ZODIAC3'))).toEqual([0x4e19, 0x5bc5])
    // 12 \* GANADA, 1 \* CHOSUNG
    expect(points(fieldNumberText(12, 'GANADA'))).toEqual([0xd0c0])
    expect(points(fieldNumberText(1, 'CHOSUNG'))).toEqual([0x3131])
    // aiueo is the half-width sequence of ST_NumberFormat
    expect(points(fieldNumberText(1, 'AIUEO'))).toEqual([0xff71])
  })

  it('leaves a value the examples do not settle to Word', () => {
    expect(fieldNumberText(100, 'CHINESENUM1')).toBeNull()
    expect(fieldNumberText(101, 'CHINESENUM3')).toBeNull()
    expect(fieldNumberText(100, 'KANJINUM2')).toBeNull()
    expect(fieldNumberText(11, 'ZODIAC1')).toBeNull()
    expect(fieldNumberText(61, 'ZODIAC3')).toBeNull()
    expect(fieldNumberText(0, 'DBNUM1')).toBeNull()
    expect(fieldNumberText(47, 'AIUEO')).toBeNull()
  })

  it('does not know the formats whose glyphs the standard does not list', () => {
    for (const format of ['KANJINUM3', 'DBNUM2', 'DBNUM3', 'IROHA', 'HEBREW1', 'THAILETTER']) {
      expect(isFieldNumberFormat(format)).toBe(false)
      expect(fieldNumberText(12, format)).toBeNull()
    }
  })

  it('writes the worded and lettered formats', () => {
    expect(fieldNumberText(123, 'ArabicDash')).toBe('- 123 -')
    expect(fieldNumberText(32, 'Ordinal')).toBe('32nd')
    expect(fieldNumberText(123, 'CardText')).toBe('one hundred twenty-three')
    expect(fieldNumberText(1235, 'OrdText')).toBe('one thousand two hundred thirty-fifth')
    expect(fieldNumberText(1234.567, 'DollarText')).toBe(
      'one thousand two hundred thirty-four and 57/100',
    )
    expect(fieldNumberText(255, 'Hex')).toBe('FF')
    expect(fieldNumberText(123, 'Roman')).toBe('CXXIII')
    expect(fieldNumberText(123, 'ROMAN')).toBe('CXXIII')
    expect(fieldNumberText(123, 'roman')).toBe('cxxiii')
    expect(fieldNumberText(54, 'ALPHABETIC')).toBe('BBB')
    expect(fieldNumberText(52, 'alphabetic')).toBe('zz')
    expect(fieldNumberText(4000, 'roman')).toBeNull()
    expect(isFieldNumberFormat('Upper')).toBe(false)
  })

  it('a SEQ field counts in them', () => {
    const counter = new SeqCounter()
    const next = (instr: string) => counter.next(parseSeqInstruction(instr)!).text
    expect(next('SEQ Figure \\* CardText')).toBe('one')
    expect(next('SEQ Figure \\* Ordinal')).toBe('2nd')
    expect(next('SEQ Figure \\* Roman')).toBe('III')
    expect(points(next('SEQ Figure \\* CHINESENUM3'))).toEqual([0x56db])
    // a format with no rule here leaves the field to Word
    expect(next('SEQ Figure \\* KANJINUM3')).toBeNull()
    expect(parseSeqInstruction('SEQ Figure \\* Ordinal')).toMatchObject({
      format: null,
      formatName: 'Ordinal',
    })
  })
})
