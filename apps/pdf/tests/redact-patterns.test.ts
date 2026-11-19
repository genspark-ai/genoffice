import { describe, expect, it } from 'vitest'
import { IBAN_LENGTHS, findPattern } from '../src/renderer/redact-patterns'
import type { RedactPattern } from '../src/renderer/redact-patterns'
import { findPatternsInIndex, patternMarks } from '../src/renderer/search'
import type { SearchIndex } from '../src/renderer/search'

const found = (p: RedactPattern, s: string): string[] =>
  findPattern(p, s).map(([a, b]) => s.slice(a, b))

describe('redact patterns — phone', () => {
  it('matches grouped numbers, not plain digit runs', () => {
    expect(
      found(
        'phone',
        'Call (555) 123-4567 or 555.987.6543, +1 555 222 3333; not 12345 or 5551234567890',
      ),
    ).toEqual(['(555) 123-4567', '555.987.6543', '+1 555 222 3333'])
    expect(found('phone', 'Tel. +49 2151 123456.')).toEqual(['+49 2151 123456'])
    expect(
      found(
        'phone',
        '+44 20 7946 0018, +33 1 23 45 67 89, +49 (0) 2151 123456, (0211) 123456, 030 12345678, 01 23 45 67 89',
      ),
    ).toEqual([
      '+44 20 7946 0018',
      '+33 1 23 45 67 89',
      '+49 (0) 2151 123456',
      '(0211) 123456',
      '030 12345678',
      '01 23 45 67 89',
    ])
  })

  it('skips dates, decimals, unseparated, too short or too long', () => {
    expect(
      found(
        'phone',
        'on 01.10.2026 or 09-10-2026, ratio 0.1234567, id 0211123456, +4921511234567, +1 234 567, 0211 12, +49 1234 5678 9012 3456',
      ),
    ).toEqual([])
  })

  it('accepts slashed and labelled unseparated numbers', () => {
    const lines = [
      'A tel. 0426 123456',
      'B tel. +39 0426 123456',
      'C cell. 345 1234567',
      'D cell. 3451234567',
      'E tel. 0426/123456',
      'F cell. 345.123.4567',
      'G tel. 0426.123456',
      'H cell. +393451234567',
      'Fax: 0211123456',
      'Telefono 0426/123456',
    ]
    expect(found('phone', lines.join('\n'))).toEqual([
      '0426 123456',
      '+39 0426 123456',
      '345 1234567',
      '3451234567',
      '0426/123456',
      '345.123.4567',
      '0426.123456',
      '+393451234567',
      '0211123456',
      '0426/123456',
    ])
    expect(
      found(
        'phone',
        'id 0211123456, +4921511234567, hotel 3451234567, prot. 2024/123456, on 01/10/2026, ' +
          'tel. 12345, tel 0426/ 123456, +39 0426/123456, tel 0426/12/3456, tel 0426 12/3456',
      ),
    ).toEqual([])
  })

  it('labelled Italian international slash numbers match whole', () => {
    for (const [text, num] of [
      ['tel. +39 0426/123456', '+39 0426/123456'],
      ['Fax: +390426/123456', '+390426/123456'],
      ['Telefono +39 02/12345678', '+39 02/12345678'],
      ['tel. +39 02/1234', '+39 02/1234'],
    ]) {
      expect(found('phone', text!), text).toEqual([num])
    }
    for (const text of [
      '+39 0426/123456',
      'tel. +39 2024/123456',
      'tel. +39 0426/12/3456',
      'tel. +39 02/123',
      'tel. +39-0426/123456',
      'tel. +40 0426/123456',
    ]) {
      expect(found('phone', text), text).toEqual([])
    }
  })

  it('the tail of a longer number is not a phone number', () => {
    for (const s of [
      'NIR 1 85 05 78 006 084 36 end',
      'IBAN FR76 3000 6000 0112 3456 7890 189',
      'SIRET 123 456 782 00010',
    ]) {
      expect(found('phone', s), s).toEqual([])
    }
    expect(found('phone', '01 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7')).toEqual([])
  })
})

describe('redact patterns — email', () => {
  it('matches addresses with a real TLD', () => {
    expect(
      found('email', 'Write to ada.lovelace+pdf@example.co.uk, or bob@host (no TLD) or x@y.z.'),
    ).toEqual(['ada.lovelace+pdf@example.co.uk'])
    expect(found('email', 'Write to _john@example.com now')).toEqual(['_john@example.com'])
    expect(found('email', 'Write to ..@example.com or -@example.com')).toEqual([])
  })

  it('handles internationalized addresses and spaceless scripts', () => {
    expect(
      found('email', 'Yazın: ayşe.yılmaz@örnek.com.tr; müller@bücher.de, иван@пример.рф.'),
    ).toEqual(['ayşe.yılmaz@örnek.com.tr', 'müller@bücher.de', 'иван@пример.рф'])
    expect(found('email', 'メールはtaro@example.jpまで、联系support@example.cn谢谢')).toEqual([
      'taro@example.jp',
      'support@example.cn',
    ])
  })
})

describe('redact patterns — credit card', () => {
  it('needs a valid Luhn checksum', () => {
    expect(
      found(
        'credit-card',
        'Visa 4111 1111 1111 1111, bad 4111 1111 1111 1112, amex 3782-822463-10005',
      ),
    ).toEqual(['4111 1111 1111 1111', '3782-822463-10005'])
  })

  it('drops a trailing expiry or CVV but not a longer number', () => {
    for (const tail of ['12/26', '2026', '123', '12 26']) {
      expect(found('credit-card', `Card 4111 1111 1111 1111 ${tail}`), tail).toEqual([
        '4111 1111 1111 1111',
      ])
    }
    expect(found('credit-card', 'Card 4111 1111 1111 1112 12/26')).toEqual([])
    expect(found('credit-card', 'IBAN 4111 1111 1111 1111 0000 0000')).toEqual([])
  })
})

describe('redact patterns — SSN', () => {
  it('matches valid US SSNs only', () => {
    expect(
      found(
        'ssn',
        'SSN 123-45-6789 and 123 45 6789; invalid 000-12-3456, 666-12-3456, 912-12-3456, 1234-56-7890',
      ),
    ).toEqual(['123-45-6789', '123 45 6789'])
  })
})

describe('redact patterns — date', () => {
  it('matches numeric and month-name dates', () => {
    expect(
      found(
        'date',
        'Due 10/01/2026, 2026-10-01, 1.10.26, March 5, 2024, Jan. 7 2025 and 5 June 2023; not 10/2026 or Mayday 12',
      ),
    ).toEqual([
      '10/01/2026',
      '2026-10-01',
      '1.10.26',
      'March 5, 2024',
      'Jan. 7 2025',
      '5 June 2023',
    ])
  })

  it('folds Turkish i and matches German, French, Italian months', () => {
    expect(found('date', '10 Ekim 2026, 3 MAYIS 2025, 1 EKİM 2026, 7 Şub. 2025')).toEqual([
      '10 Ekim 2026',
      '3 MAYIS 2025',
      '1 EKİM 2026',
      '7 Şub. 2025',
    ])
    expect(found('date', 'am 10. Oktober 2026, le 14 juillet 2024, il 25 dicembre 2022')).toEqual([
      '10. Oktober 2026',
      '14 juillet 2024',
      '25 dicembre 2022',
    ])
    expect(found('date', 'Mai 2024, 5 Martı 2024, 5 Oktoberfest 2024, 12 Aralık')).toEqual([])
  })
})

/** One example IBAN per country of the SWIFT registry (release 99). */
const IBAN_EXAMPLES = [
  'AD1200012030200359100100',
  'AE070331234567890123456',
  'AL47212110090000000235698741',
  'AT611904300234573201',
  'AZ21NABZ00000000137010001944',
  'BA391290079401028494',
  'BE68539007547034',
  'BG80BNBG96611020345678',
  'BH67BMAG00001299123456',
  'BI4210000100010000332045181',
  'BR1800360305000010009795493C1',
  'BY13NBRB3600900000002Z00AB00',
  'CH9300762011623852957',
  'CR05015202001026284066',
  'CY17002001280000001200527600',
  'CZ6508000000192000145399',
  'DE89370400440532013000',
  'DJ2100010000000154000100186',
  'DK5000400440116243',
  'DO28BAGR00000001212453611324',
  'EE382200221020145685',
  'EG380019000500000000263180002',
  'ES9121000418450200051332',
  'FI2112345600000785',
  'FK88SC123456789012',
  'FO6264600001631634',
  'FR1420041010050500013M02606',
  'GB29NWBK60161331926819',
  'GE29NB0000000101904917',
  'GI75NWBK000000007099453',
  'GL8964710001000206',
  'GR1601101250000000012300695',
  'GT82TRAJ01020000001210029690',
  'HN88CABF00000000000250005469',
  'HR1210010051863000160',
  'HU42117730161111101800000000',
  'IE29AIBK93115212345678',
  'IL620108000000099999999',
  'IQ98NBIQ850123456789012',
  'IS140159260076545510730339',
  'IT60X0542811101000000123456',
  'JO94CBJO0010000000000131000302',
  'KW81CBKU0000000000001234560101',
  'KZ86125KZT5004100100',
  'LB62099900000001001901229114',
  'LC55HEMM000100010012001200023015',
  'LI21088100002324013AA',
  'LT121000011101001000',
  'LU280019400644750000',
  'LV80BANK0000435195001',
  'LY83002048000020100120361',
  'MC5811222000010123456789030',
  'MD24AG000225100013104168',
  'ME25505000012345678951',
  'MK07250120000058984',
  'MN121234123456789123',
  'MR1300020001010000123456753',
  'MT84MALT011000012345MTLCAST001S',
  'MU17BOMM0101101030300200000MUR',
  'NI45BAPR00000013000003558124',
  'NL91ABNA0417164300',
  'NO9386011117947',
  'OM810180000001299123456',
  'PK36SCBL0000001123456702',
  'PL61109010140000071219812874',
  'PS92PALS000000000400123456702',
  'PT50000201231234567890154',
  'QA58DOHB00001234567890ABCDEFG',
  'RO49AAAA1B31007593840000',
  'RS35260005601001611379',
  'RU0304452522540817810538091310419',
  'SA0380000000608010167519',
  'SC18SSCB11010000000000001497USD',
  'SD2129010501234001',
  'SE4550000000058398257466',
  'SI56263300012039086',
  'SK3112000000198742637541',
  'SM86U0322509800000000270100',
  'SO211000001001000100141',
  'ST68000100010051845310112',
  'SV62CENR00000000000000700025',
  'TL380080012345678910157',
  'TN5910006035183598478831',
  'TR330006100519786457841326',
  'UA213223130000026007233566001',
  'VA59001123000012345678',
  'VG96VPVG0000012345678901',
  'XK051212012345678906',
  'YE15CBYE0001018861234567891234',
]

const printed = (iban: string): string => iban.match(/.{1,4}/g)!.join(' ')

describe('redact patterns — IBAN', () => {
  it('registry examples cover every country length', () => {
    expect(Object.fromEntries(IBAN_EXAMPLES.map((e) => [e.slice(0, 2), e.length]))).toEqual(
      IBAN_LENGTHS,
    )
  })

  it('matches every registry example compact and printed, rejects a changed digit', () => {
    for (const ex of IBAN_EXAMPLES) {
      expect(found('iban', ex), ex).toEqual([ex])
      expect(found('iban', printed(ex)), ex).toEqual([printed(ex)])
      const at = ex.search(/\d(?=\D*$)/)
      const changed = ex.slice(0, at) + ((Number(ex[at]) + 1) % 10) + ex.slice(at + 1)
      expect(found('iban', changed), changed).toEqual([])
    }
  })

  it('rejects unknown codes, wrong length, glued tails and stray spaces', () => {
    for (const s of [
      'XX65000000000000000000',
      'DE8300000000000',
      'DE89370400440532013000ABC',
      'ABCDE89370400440532013000',
      'DEAB 3704 0044 0532 0130 00',
      'DE89 3704 0044 0532 013 000',
      'DE89  3704 0044 0532 0130 00',
      'Flight BA37 leaves at noon from gate four',
    ]) {
      expect(found('iban', s), s).toEqual([])
    }
    expect(found('iban', 'Pay to DE89 3704 0044 0532 0130 00 123 today')).toEqual([
      'DE89 3704 0044 0532 0130 00',
    ])
    expect(found('iban', 'DE893704 0044 0532013000')).toEqual(['DE893704 0044 0532013000'])
  })
})

describe('findPatternsInIndex', () => {
  it('maps hits onto item geometry, across items', () => {
    const text = 'Mail ada@example.com Card 4111 1111 1111 1111'
    const index: SearchIndex = [
      {
        text,
        lower: text.toLowerCase(),
        items: [
          { start: 0, end: 21, x: 0, y: 700, w: 210, h: 10 },
          { start: 21, end: text.length, x: 300, y: 680, w: 240, h: 10 },
        ],
      },
    ]
    const hits = findPatternsInIndex(index, ['email', 'credit-card'])
    expect(hits.map((h) => [h.pattern, h.text])).toEqual([
      ['email', 'ada@example.com'],
      ['credit-card', '4111 1111 1111 1111'],
    ])
    expect(hits[0]!.rects).toEqual([[50, 700, 200, 710]])
    expect(hits[1]!.rects).toHaveLength(1)
    expect(hits[1]!.rects[0]![2]).toBeCloseTo(540)
  })
})

describe('patternMarks', () => {
  const hit = (pageIndex: number, n: number, rotated = false) => ({
    pageIndex,
    pattern: 'email' as const,
    text: 'x',
    rects: Array.from({ length: n }, (_, i): [number, number, number, number] => [
      i * 10,
      100,
      i * 10 + 5,
      110,
    ]),
    ...(rotated ? { rotated: true } : {}),
  })

  it('pads rects, skips rotated text, and stops whole hits at the cap', () => {
    const r = patternMarks([hit(0, 1), hit(0, 1, true), hit(0, 2), hit(0, 1)], 2)
    expect(r.hits).toBe(1)
    expect(r.rotated).toBe(1)
    expect(r.capped).toBe(true)
    expect(r.marks).toEqual([{ pageIndex: 0, rect: [-1, 97.5, 6, 110.5] }])
  })
})

describe('findPatternsInIndex skipPage', () => {
  it('skipped pages do not consume the match cap', () => {
    const busy = 'a@b.co '.repeat(1100)
    const live = 'x@y.co'
    const page = (text: string) => ({
      text,
      lower: text,
      items: [{ start: 0, end: text.length, x: 0, y: 700, w: text.length * 5, h: 10 }],
    })
    const index = [page(busy), page(live)]
    expect(findPatternsInIndex(index, ['email']).some((m) => m.pageIndex === 1)).toBe(false)
    const kept = findPatternsInIndex(index, ['email'], (i) => i === 0)
    expect(kept).toHaveLength(1)
    expect(kept[0]!.text).toBe('x@y.co')
  })
})
