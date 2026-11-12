/**
 * Dumps cssFontFamily for every known face name (Word face table, system
 * snapshot, bundled fonts.css families, corpus names, branch probes) so the
 * face-chain table refactor can prove identical chains. Skipped unless
 * FACE_CHAINS_OUT names the output file.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import wordFaces from '@genoffice/font-metrics/word-faces.json'
import { cssFontFamily, setDocFontTable } from '../../src/renderer/line-metrics'

const out = process.env.FACE_CHAINS_OUT ?? ''

const PROBES = [
  'Calibri',
  'Calibri Light',
  'Aptos',
  'Aptos Display',
  'Aptos Narrow',
  'Aptos Mono',
  'Aptos Serif',
  'Cambria',
  'Cambria Math',
  'STIX Two Math',
  'Latin Modern Math',
  'Times New Roman',
  'Times',
  'MS Reference Sans Serif',
  'Merriweather',
  'Merriweather Sans',
  'Palatino Linotype',
  'Book Antiqua',
  'Arial',
  'Arial Black',
  'Arial Narrow',
  'Roboto',
  'Roboto Mono',
  'Garamond',
  'Berlin Sans FB Demi',
  'Berlin Sans FB',
  'Segoe UI',
  'Segoe UI Emoji',
  'Consolas',
  'Century Gothic',
  'Lucida Handwriting',
  'Courier New',
  'Menlo',
  'Nunito Sans',
  'Poppins',
  'DM Sans',
  'Montserrat',
  'Montserrat Bold',
  'Montserrat-Bold',
  'Microsoft New Tai Lue',
  'STZhongsong',
  '\u4e2d\u5b8b',
  'SimSun',
  'NSimSun',
  '\u5b8b\u4f53',
  'SimHei',
  '\u9ed1\u4f53',
  '\u7ec6\u9ed1',
  'Microsoft YaHei',
  '\u5fae\u8f6f\u96c5\u9ed1',
  'DengXian',
  '\u7b49\u7ebf',
  'FangSong',
  '\u4eff\u5b8b',
  '\u4eff\u5b8b_GB2312',
  'KaiTi',
  '\u6977\u4f53',
  'LiSu',
  '\u96b6\u4e66',
  '\u65b9\u6b63\u5c0f\u6807\u5b8b',
  '\u65b9\u6b63\u9b4f\u7891_GBK',
  'Heiti SC',
  'Songti SC',
  'PingFang SC',
  'Noto Sans CJK SC',
  'Noto Serif CJK SC',
  'Noto Sans CJK KR',
  'Noto Serif CJK JP',
  'Noto Sans CJK TC',
  'Source Han Sans CN',
  'Noto Sans Mono CJK SC',
  'Meiryo',
  'Meiryo UI',
  '\u30e1\u30a4\u30ea\u30aa',
  'MS Gothic',
  'MS PGothic',
  'MS UI Gothic',
  'MS Mincho',
  'MS PMincho',
  'ＭＳ \u660e\u671d',
  'ＭＳ \u30b4\u30b7\u30c3\u30af',
  'Yu Gothic',
  'Yu Gothic Light',
  'Yu Mincho',
  'Yu Mincho Light',
  'Yu Gothic UI',
  '\u6e38\u660e\u671d',
  'Hiragino Sans',
  'Hiragino Mincho ProN',
  'Kozuka Gothic Pro',
  'BIZ UDPGothic',
  'Osaka',
  'Malgun Gothic',
  '맑은 고딕',
  'Batang',
  '바탕',
  'BatangChe',
  'Gulim',
  '굴림',
  'GulimChe',
  'Dotum',
  'DotumChe',
  'Gungsuh',
  'GungsuhChe',
  'Nanum Gothic',
  'NanumMyeongjo',
  'Apple SD Gothic Neo',
  'HY견고딕',
  'DFKai-SB',
  '\u6a19\u6977\u9ad4',
  'Microsoft JhengHei',
  'PMingLiU',
  'MingLiU',
  'BiauKai',
  'PingFang TC',
  'Songti TC',
  '\u7d30\u660e\u9ad4',
  'Nyala',
  'Ebrima',
  'Abyssinica SIL',
  'Nirmala UI',
  'Latha',
  'Vijaya',
  'Tamil MN',
  'TH SarabunPSK',
  'Cordia New',
  'Angsana New',
  'Tahoma',
  'Traditional Arabic',
  'Simplified Arabic',
  'Sakkal Majalla',
  'Noto Naskh Arabic',
  'Noto Sans Arabic',
  'Noto Kufi Arabic',
  'Dubai',
  'B Nazanin',
  'XB Zar',
  'Akhbar MT',
  'Arabic Typesetting',
  'Georgia',
  'Verdana',
  'Helvetica',
  'Helvetica Neue',
  'Inter',
  'Lato',
  'Open Sans',
  'Ubuntu',
  'Unknown Face',
  'Some Serif Name',
  'Some Song Name',
  'Symbol',
  'Wingdings',
  'Webdings',
]

function fontsCssFamilies(): string[] {
  const css = readFileSync(join(__dirname, '../../src/renderer/fonts/fonts.css'), 'utf8')
  return [...new Set([...css.matchAll(/font-family: '([^']+)'/g)].map((m) => m[1]))]
}

describe.skipIf(!out)('face chains dump', () => {
  it('writes cssFontFamily for every known name', () => {
    const faces = wordFaces as {
      faces: Record<string, { family: string }>
      system?: Record<string, { family?: string }>
    }
    const names = new Set<string>(PROBES)
    for (const f of Object.values(faces.faces)) names.add(f.family)
    for (const [k, v] of Object.entries(faces.system ?? {})) names.add(v.family ?? k)
    for (const f of fontsCssFamilies()) names.add(f)
    const result: Record<string, { plain: string; charset81: string; altArial: string }> = {}
    for (const name of [...names].sort()) {
      setDocFontTable([])
      const plain = cssFontFamily(name)
      setDocFontTable([{ name, charset: '81' }])
      const charset81 = cssFontFamily(name)
      setDocFontTable([{ name, altName: 'Arial' }])
      const altArial = cssFontFamily(name)
      result[name] = { plain, charset81, altArial }
    }
    setDocFontTable([])
    writeFileSync(out, JSON.stringify(result, null, 1))
    expect(Object.keys(result).length).toBeGreaterThan(200)
  }, 120_000)
})
