import { Schema, type Node as PMNode } from '@tiptap/pm/model'
import { beforeAll, describe, expect, it } from 'vitest'
import {
  blockHyphenWords,
  breakPoints,
  hyphenatesSomeText,
  hyphenWords,
  hyphenationConfigOf,
  hyphenationDictFor,
  hyphenationLangOf,
  hyphenationSettingsOf,
  hyphenationStyleDicts,
  hyphenationZoneApplies,
  loadHyphenation,
} from '../src/renderer/editor/hyphenation'

const split = (dict: string, word: string) => {
  const points = breakPoints(dict, word)
  return points.reduce<string[]>(
    (parts, p, i) => [...parts, word.slice(p, points[i + 1] ?? word.length)],
    [word.slice(0, points[0] ?? word.length)],
  )
}

describe('hyphenation patterns', () => {
  beforeAll(async () => {
    expect(await loadHyphenation('en-us')).toBe(true)
    expect(await loadHyphenation('de')).toBe(true)
  })

  it('maps a document language to a pattern set', () => {
    expect(hyphenationDictFor('en-US')).toBe('en-us')
    expect(hyphenationDictFor('en-AU')).toBe('en-gb')
    expect(hyphenationDictFor('de-DE')).toBe('de')
    expect(hyphenationDictFor('pt-BR')).toBe('pt')
    // as Word: text that declares no language is not hyphenated
    expect(hyphenationDictFor(undefined)).toBeNull()
    expect(hyphenationDictFor('')).toBeNull()
    expect(hyphenationDictFor('ru-RU')).toBeNull()
    expect(hyphenationDictFor('zh-CN')).toBeNull()
  })

  it('has no break points before the patterns load', () => {
    expect(breakPoints('fr', 'internationalisation')).toEqual([])
  })

  it('breaks words the way the patterns say, whatever the case', () => {
    expect(split('en-us', 'hyphenation')).toEqual(['hy', 'phen', 'ation'])
    expect(split('en-us', 'HYPHENATION')).toEqual(['HY', 'PHEN', 'ATION'])
    expect(split('en-us', 'Hyphenation')).toEqual(['Hy', 'phen', 'ation'])
    expect(split('de', 'Silbentrennung')).toEqual(['Sil', 'ben', 'tren', 'nung'])
  })

  it('leaves short words whole', () => {
    expect(breakPoints('en-us', 'into')).toEqual([])
    expect(hyphenWords('a into the', 'en-us', true)).toEqual([])
  })

  it('finds the words of a line with their offsets', () => {
    const words = hyphenWords('An international table', 'en-us', true)
    expect(words.map((w) => [w.from, w.to, w.text])).toEqual([
      [3, 16, 'international'],
      [17, 22, 'table'],
    ])
    expect(words[0].points).toEqual([2, 5, 7, 11])
  })

  it('skips words in capitals only when told to', () => {
    const text = 'INTERNATIONAL International international'
    expect(hyphenWords(text, 'en-us', true).map((w) => w.text)).toEqual([
      'INTERNATIONAL',
      'International',
      'international',
    ])
    expect(hyphenWords(text, 'en-us', false).map((w) => w.text)).toEqual([
      'International',
      'international',
    ])
  })

  it('keeps addresses, codes and words with an optional hyphen whole', () => {
    expect(hyphenWords('https://example.com/international', 'en-us', true)).toEqual([])
    expect(hyphenWords('someone@international.example', 'en-us', true)).toEqual([])
    expect(hyphenWords('interna\u00adtional', 'en-us', true)).toEqual([])
    expect(hyphenWords('model2international', 'en-us', true)).toEqual([])
  })

  it('breaks the parts of a compound on their own', () => {
    const words = hyphenWords('well-international', 'en-us', true)
    expect(words.map((w) => [w.from, w.text])).toEqual([[5, 'international']])
  })
})

describe('hyphenation settings of a document', () => {
  it('reads Word defaults from a file that says nothing', () => {
    expect(hyphenationSettingsOf({})).toEqual({ caps: true, limit: 0, zoneTwips: 360 })
  })

  it('reads the file values and lets the session edit win', () => {
    const parsed = {
      doNotHyphenateCaps: true,
      consecutiveHyphenLimit: 2,
      hyphenationZoneTwips: 425,
    }
    expect(hyphenationSettingsOf(parsed)).toEqual({ caps: false, limit: 2, zoneTwips: 425 })
    expect(hyphenationSettingsOf(parsed, { caps: true, limit: 0 })).toEqual({
      caps: true,
      limit: 0,
      zoneTwips: 425,
    })
  })

  it('switches our hyphenation off when the document does not hyphenate', () => {
    expect(hyphenationConfigOf({ docDefaults: { lang: 'en-US' } }, false).dict).toBeNull()
    expect(hyphenationConfigOf({ docDefaults: { lang: 'en-US' } }, true).dict).toBe('en-us')
    expect(hyphenationConfigOf({ docDefaults: { lang: 'ru-RU' } }, true).dict).toBeNull()
    expect(hyphenationConfigOf({}, true).dict).toBeNull()
    expect(hyphenationConfigOf({ docDefaults: {} }, true).dict).toBeNull()
    // the default paragraph style's language stands before the document defaults'
    const styles = new Map([
      ['Heading1', { type: 'paragraph', display: { lang: 'fr-FR' } }],
      ['Normal', { isDefault: true, type: 'paragraph', display: { lang: 'de-DE' } }],
    ])
    expect(hyphenationLangOf({ styles })).toBe('de-DE')
    expect(hyphenationConfigOf({ styles, docDefaults: { lang: 'en-US' } }, true).dict).toBe('de')
    expect(hyphenationLangOf({ styles: new Map(), docDefaults: { lang: 'en-GB' } })).toBe('en-GB')
  })

  it('applies the zone only in the layout modes that have one', () => {
    const parsed = { hyphenationZoneTwips: 300, docDefaults: { lang: 'en-US' } }
    expect(hyphenationZoneApplies(parsed)).toBe(true)
    expect(hyphenationConfigOf(parsed, true).zonePx).toBe(20)
    expect(hyphenationZoneApplies({ ...parsed, compatibilityMode: 15 })).toBe(false)
    expect(hyphenationConfigOf({ ...parsed, compatibilityMode: 15 }, true).zonePx).toBe(0)
  })
})

const schema = new Schema({
  nodes: {
    doc: { content: 'paragraph+' },
    paragraph: { content: 'text*', attrs: { styleId: { default: null } } },
    text: {},
  },
  marks: {
    textStyle: { attrs: { lang: { default: null }, styleId: { default: null } } },
  },
})
const WORD = 'internationalization'
const run = (text: string, attrs: { lang?: string; styleId?: string } = {}) =>
  schema.text(text, [
    schema.marks.textStyle.create({
      lang: attrs.lang ? JSON.stringify({ val: attrs.lang }) : null,
      styleId: attrs.styleId ?? null,
    }),
  ])
const para = (styleId: string | null, ...runs: PMNode[]) =>
  schema.nodes.paragraph.create({ styleId }, runs)

describe('the language a run hyphenates by', () => {
  beforeAll(async () => {
    expect(await loadHyphenation('en-us')).toBe(true)
    expect(await loadHyphenation('de')).toBe(true)
  })

  // Word for Mac 16.106, exports of one text: a run, a character style and a paragraph
  // style (own or inherited) naming en-US each hyphenate under defaults that name none
  const styles = new Map([
    ['Normal', { isDefault: true, type: 'paragraph', display: {} }],
    ['EnPara', { type: 'paragraph', display: { lang: 'en-US' } }],
    ['EnChild', { type: 'paragraph', display: { lang: 'en-US' } }],
    ['EnChar', { type: 'character', display: { lang: 'en-US' } }],
    ['ZhChar', { type: 'character', display: { lang: 'zh-CN' } }],
    ['Plain', { type: 'character', display: {} }],
  ])
  const none = hyphenationConfigOf({ styles }, true)
  const words = (block: PMNode, config = none) => blockHyphenWords(block, config).map((w) => w.text)

  it('stays on for a document whose defaults name no language', () => {
    expect(none.on).toBe(true)
    expect(none.dict).toBeNull()
    expect(hyphenationConfigOf({ styles }, false).on).toBe(false)
    expect(hyphenationStyleDicts({ styles })).toEqual(
      new Map([
        ['Normal', null],
        ['EnPara', 'en-us'],
        ['EnChild', 'en-us'],
        ['EnChar', 'en-us'],
        ['ZhChar', null],
      ]),
    )
  })

  it('cuts a run that names a language and leaves the rest of the paragraph whole', () => {
    expect(words(para(null, run(WORD)))).toEqual([])
    expect(words(para(null, run(WORD, { lang: 'en-US' })))).toEqual([WORD])
    expect(words(para(null, run(`${WORD} `), run(WORD, { lang: 'en-US' })))).toEqual([WORD])
    expect(words(para(null, run(WORD, { lang: 'zh-CN' })))).toEqual([])
  })

  it('takes the character style before the paragraph style, and the run before both', () => {
    expect(words(para(null, run(WORD, { styleId: 'EnChar' })))).toEqual([WORD])
    expect(words(para('EnPara', run(WORD)))).toEqual([WORD])
    expect(words(para('EnChild', run(WORD)))).toEqual([WORD])
    expect(words(para('EnPara', run(WORD, { styleId: 'Plain' })))).toEqual([WORD])
    expect(words(para('EnPara', run(WORD, { styleId: 'ZhChar' })))).toEqual([])
    expect(words(para('EnPara', run(WORD, { lang: 'zh-CN' })))).toEqual([])
    expect(words(para('EnPara', run(WORD, { styleId: 'ZhChar', lang: 'en-US' })))).toEqual([WORD])
  })

  it('lets a style or a run name another language than the defaults', () => {
    const facts = {
      docDefaults: { lang: 'en-US' },
      styles: new Map([
        ['ZhPara', { type: 'paragraph', display: { lang: 'zh-CN' } }],
        ['Body', { type: 'paragraph', display: {} }],
      ]),
    }
    const config = hyphenationConfigOf(facts, true)
    expect(words(para(null, run(WORD)), config)).toEqual([WORD])
    expect(words(para('Body', run(WORD)), config)).toEqual([WORD])
    expect(words(para('ZhPara', run(WORD)), config)).toEqual([])
    expect(words(para('ZhPara', run(WORD, { lang: 'en-US' })), config)).toEqual([WORD])
    const german = blockHyphenWords(para(null, run('Silbentrennung', { lang: 'de-DE' })), config)
    expect(german.map((w) => w.points)).toEqual([breakPoints('de', 'Silbentrennung')])
  })

  it('offers hyphenation when some text names a language', () => {
    const doc = (...blocks: PMNode[]) => schema.nodes.doc.create(null, blocks)
    expect(hyphenatesSomeText(doc(para(null, run(WORD))), { styles })).toBe(false)
    expect(hyphenatesSomeText(doc(para(null, run(WORD, { lang: 'en-US' }))), { styles })).toBe(true)
    expect(hyphenatesSomeText(doc(para('EnPara', run(WORD))), { styles })).toBe(true)
    expect(hyphenatesSomeText(doc(para(null, run(WORD, { lang: 'ru-RU' }))), { styles })).toBe(
      false,
    )
    expect(hyphenatesSomeText(null, { docDefaults: { lang: 'en-US' } })).toBe(true)
  })
})
