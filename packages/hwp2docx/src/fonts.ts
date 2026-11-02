/**
 * Font table for the converted document. Runs keep the original Hangul font
 * names (휴먼명조, 한양중고딕, HY헤드라인M, ...), so a downloaded .docx opens in
 * those faces wherever they are installed. Where they are not, Word — and
 * GenOffice, which follows Word's substitution rules — falls back by the
 * word/fontTable.xml hints: charset 81 (Hangul) + family roman → Batang class,
 * charset 81 + swiss/modern → Malgun class. Without an entry every
 * Hangul-lettered name lands in Batang, so gothic/headline faces would render
 * as serif.
 *
 * No w:altName is written: Word for Mac lays a face out in its altName even
 * when the named face is installed, which would hide the real font.
 */

import type JSZip from 'jszip'
import { addDocumentPart } from './parts'
import { escapeXml } from './xml'

export type FontClass = 'serif' | 'sans' | 'mono'

export interface FontHint {
  name: string
  cls: FontClass
  /** a face with Hangul glyphs (charset 81) rather than a Latin-only face */
  hangul: boolean
}

/**
 * Hangul/Hancom font families seen in HWP/HWPX documents, by class. Matched
 * against the NFC name with spaces removed; the first matching rule wins.
 * Fixed-pitch "-체" cuts (돋움체, 굴림체, 바탕체, 궁서체) are their own class.
 */
const RULES: Array<[RegExp, FontClass]> = [
  [/^(돋움|굴림|바탕|궁서|새굴림)체$|^(dotum|gulim|batang|gungsuh)che$|코딩|coding|mono/i, 'mono'],
  [
    /고딕|돋움|굴림|헤드라인|그래픽|맑은|윤고딕|나눔스퀘어|나눔바른|페이퍼로지|프리텐다드|pretendard|gothic|dotum|gulim|malgun|sans|헤드|견출|엽서|목각|둥근|모음|체크|spoqa|본고딕|돋음|튼튼/i,
    'sans',
  ],
  [
    /명조|바탕|신명|견명|궁서|해서|옛한글|serif|myeongjo|myungjo|batang|mincho|gungsuh|본명조|kopub바탕/i,
    'serif',
  ],
]

/** Latin-only faces that ship with Hangul or Office, with their class */
const LATIN_ONLY: Array<[RegExp, FontClass]> = [
  [
    /^(hci |hy-?|hancom )?(tulip|morning glory|hollyhock|centaurea)\b|^(times new roman|cambria|georgia|garamond|book antiqua|palatino)/i,
    'serif',
  ],
  [/^(courier new|consolas|lucida console)/i, 'mono'],
  [
    /^(hci |hy-?|hancom )?(poppy|acacia)\b|^(arial|calibri|tahoma|verdana|helvetica|segoe ui|symbol|wingdings)/i,
    'sans',
  ],
]

const HANGUL_RE = /[가-힣ㄱ-ㆎ]/

export function classifyFont(name: string): FontHint {
  const compact = name.normalize('NFC').replace(/\s+/g, '')
  const latin = LATIN_ONLY.find(([re]) => re.test(name.trim()))
  const latinOnly = Boolean(latin)
  const rule = latin ? undefined : RULES.find(([re]) => re.test(compact))
  const cls: FontClass = latin ? latin[1] : rule ? rule[1] : 'serif'
  // a Hangul-lettered or Hangul-family name carries Hangul glyphs; Hangul's
  // default face (함초롬바탕) is serif, so unknown Korean names go serif
  const hangul = !latinOnly && (HANGUL_RE.test(name) || Boolean(rule))
  return { name, cls, hangul }
}

function entryXml(hint: FontHint): string {
  const family = hint.cls === 'serif' ? 'roman' : hint.cls === 'mono' ? 'modern' : 'swiss'
  const pitch = hint.cls === 'mono' ? 'fixed' : 'variable'
  // PANOSE byte 2: 02 = cove serif, 0B = normal sans; Word reads it when choosing a substitute
  const panose =
    hint.cls === 'serif'
      ? '02030600000101010101'
      : hint.cls === 'mono'
        ? '020B0609000101010101'
        : '020B0600000101010101'
  return (
    `<w:font w:name="${escapeXml(hint.name)}">` +
    `<w:panose1 w:val="${panose}"/>` +
    `<w:charset w:val="${hint.hangul ? '81' : '00'}"/>` +
    `<w:family w:val="${family}"/>` +
    `<w:pitch w:val="${pitch}"/>` +
    '</w:font>'
  )
}

const FONT_TABLE_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable'
const FONT_TABLE_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'

/**
 * Add word/fontTable.xml entries for every font the body uses (an existing
 * part keeps its entries; names already present are not duplicated).
 */
export async function addFontTable(zip: JSZip): Promise<void> {
  const document = await zip.file('word/document.xml')?.async('string')
  if (!document) return
  const names = new Set<string>()
  for (const match of document.matchAll(/w:(?:ascii|hAnsi|eastAsia)="([^"]+)"/g)) {
    names.add(match[1].replace(/&amp;/g, '&').replace(/&quot;/g, '"'))
  }
  if (names.size === 0) return

  const existing = await zip.file('word/fontTable.xml')?.async('string')
  const present = new Set(
    [...(existing ?? '').matchAll(/<w:font w:name="([^"]+)"/g)].map((m) => m[1]),
  )
  const entries = [...names]
    .filter((name) => !present.has(escapeXml(name)))
    .sort()
    .map((name) => entryXml(classifyFont(name)))
    .join('')
  if (!entries) return

  if (existing) {
    zip.file('word/fontTable.xml', existing.replace('</w:fonts>', `${entries}</w:fonts>`))
  } else {
    await addDocumentPart(
      zip,
      'fontTable.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        `${entries}</w:fonts>`,
      FONT_TABLE_REL,
      FONT_TABLE_CT,
    )
  }
}
