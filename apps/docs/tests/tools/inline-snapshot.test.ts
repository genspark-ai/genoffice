/**
 * Inline-layout decision snapshot over a docx corpus (tools/docs-inline-snapshot.mjs).
 * Records what every document-, face- and paragraph-level predicate answers today so
 * the text-layout consolidation waves can prove "same decisions" before any gate run.
 * Skipped unless INLINE_SNAPSHOT_CORPUS names the corpus directories.
 */
import { createHash } from 'node:crypto'
import { readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import { blockRunGroups } from '../../../../packages/docx-engine/src/parse-block-passes'
import type { Run } from '@genoffice/docx-engine'
import {
  docAutospaceOff,
  docBodyFont,
  docCjkFactor,
  docHasCjk,
  docLineFactor,
} from '../../src/renderer/doc-style-css'
import {
  autospaceSeams,
  balancedSpaceOffsets,
  charScaleEm,
  isCjkFontName,
  isCjkNamedFont,
  isJapaneseFontName,
  isKoreanFontName,
  textHasCjk,
  textHasHangul,
} from '../../src/renderer/line-metrics'
import {
  cjkDeclaredLineFactor,
  cssFontFamily,
  isEaLayoutFace,
  isKoreanLayoutFace,
  krCharsetRoute,
  lineHeightFactor,
  missingLatinSubstitute,
  setDocFontTable,
} from '../../src/renderer/text-layout/font-rules'
import { setOpenDocumentText } from '../../src/renderer/text-layout/open-document'
import { isEastAsianFontName } from '../../src/renderer/font-list'
import { layoutParaOfRuns, paraLineFactor } from '../../src/renderer/editor/para-layout-spec'
import { eastAsianSegments } from '../../src/renderer/editor/script-fonts'
import { inlineRunsOf, inlineSegments } from '../../src/renderer/text-layout/inline-segments'
import { usesEastAsianRules } from '../../src/renderer/editor/cjk-punct-shrink'
import {
  documentTextContext,
  eaFaceForAutospace,
  eaFaceForCjkFactor,
  eaFaceForGridStrut,
  eaFaceForKrFactor,
} from '../../src/renderer/text-layout/document-context'

const corpus = (process.env.INLINE_SNAPSHOT_CORPUS ?? '').split(',').filter(Boolean)
const out = process.env.INLINE_SNAPSHOT_OUT ?? ''

// private character classes copied from their modules; the consolidation replaces them
const CPS_CJK_RE = /[⺀-〿぀-ヿㇰ-䶿一-鿿豈-﫿＀-￯]/u
const countBy = (text: string, pred: (cp: number) => boolean): number => {
  let n = 0
  for (const ch of text) if (pred(ch.codePointAt(0)!)) n++
  return n
}
const isHangul = (cp: number) =>
  (cp >= 0xac00 && cp <= 0xd7a3) ||
  (cp >= 0x1100 && cp <= 0x11ff) ||
  (cp >= 0x3130 && cp <= 0x318f) ||
  (cp >= 0xa960 && cp <= 0xa97f) ||
  (cp >= 0xd7b0 && cp <= 0xd7ff)
const isCjk = (cp: number) =>
  isHangul(cp) ||
  (cp >= 0x3000 && cp <= 0x30ff) ||
  (cp >= 0x3400 && cp <= 0x9fff) ||
  (cp >= 0xf900 && cp <= 0xfaff) ||
  (cp >= 0xff00 && cp <= 0xffef)
const passesWide = (cp: number) =>
  (cp >= 0x1100 && cp <= 0x11ff) ||
  (cp >= 0x3040 && cp <= 0x30ff) ||
  (cp >= 0x3130 && cp <= 0x318f) ||
  (cp >= 0x31f0 && cp <= 0x31ff) ||
  (cp >= 0xac00 && cp <= 0xd7af) ||
  (cp >= 0x2e80 && cp <= 0x9fff) ||
  (cp >= 0xf900 && cp <= 0xfaff) ||
  (cp >= 0xff00 && cp <= 0xff60) ||
  cp >= 0x20000
const autospaceSide = (cp: number) =>
  !(cp >= 0x3000 && cp <= 0x303f) && cp !== 0x30fb && !(cp >= 0xff00 && cp <= 0xffef) && isCjk(cp)

const EA_LANG_LM = /^(?:ja|zh|ko)(?:[-_]|$)/i
const EA_LANG_PARSE = /^(zh|ja|ko)(-|$)/i

const hash = (s: string) => createHash('sha1').update(s).digest('hex').slice(0, 10)

function faceRecord(name: string) {
  return {
    cjkName: isCjkFontName(name),
    cjkNamed: isCjkNamedFont(name),
    eaLayout: isEaLayoutFace(name),
    ko: isKoreanFontName(name),
    koLayout: isKoreanLayoutFace(name),
    ja: isJapaneseFontName(name),
    eaName: isEastAsianFontName(name),
    lf: lineHeightFactor(name),
    cjkLf: cjkDeclaredLineFactor(name),
    kr: krCharsetRoute(name),
    missing: missingLatinSubstitute(name),
    chain: cssFontFamily(name),
  }
}

async function snapshotDoc(path: string) {
  const parsed = await parseDocx(new Uint8Array(readFileSync(path)))
  setDocFontTable(parsed.fontTable ?? [])
  const textCtx = documentTextContext(parsed)
  const bodyEa = textCtx.bodyEa
  const viaEa = textCtx.latinViaEaSlot
  const balanced = textCtx.balancedSpaces
  setOpenDocumentText(textCtx)
  const dd = parsed.docDefaults
  const doc = {
    bodyEaFont: bodyEa ?? null,
    bodyFont: docBodyFont(parsed) ?? null,
    cjkFactor: docCjkFactor(parsed),
    lineFactor: docLineFactor(parsed, docHasCjk(parsed)),
    autospaceOff: docAutospaceOff(parsed),
    latinViaEaSlot: viaEa,
    balancedSpaces: balanced,
    ddLang: dd?.lang ?? null,
    ddEaLang: dd?.eastAsiaLang ?? null,
    langEa: {
      lm: EA_LANG_LM.test(dd?.lang ?? ''),
      parseEa: EA_LANG_PARSE.test(dd?.eastAsiaLang ?? ''),
      cps: usesEastAsianRules(dd?.eastAsiaLang),
    },
    normal: (() => {
      const n = [...parsed.styles.values()].find(
        (i) => i.isDefault && i.type === 'paragraph',
      )?.display
      return n
        ? {
            font: n.font ?? null,
            fontAscii: n.fontAscii ?? null,
            ea: n.eastAsiaFont ?? null,
            eaSlotEmpty: !!n.eaSlotEmpty,
          }
        : null
    })(),
    themeMinor: parsed.themeFonts?.minor ?? null,
    policies: {
      unified: textCtx.bodyEaFace,
      bodyEa: textCtx.bodyEa,
      cjkFactor: eaFaceForCjkFactor(textCtx),
      autospace: eaFaceForAutospace(textCtx),
      kr: eaFaceForKrFactor(textCtx) ?? null,
      gridStrut: eaFaceForGridStrut(textCtx) ?? null,
    },
    dd: dd
      ? {
          ascii: dd.asciiFont ?? null,
          ea: dd.eastAsiaFont ?? null,
          eaSlotEmpty: !!dd.eaSlotEmpty,
          eaFromLang: !!dd.eaFromLang,
          eaThemeSlot: !!dd.eaThemeSlot,
        }
      : null,
  }
  const faceNames = new Set<string>()
  const addFace = (n: string | undefined | null) => {
    if (n) faceNames.add(n)
  }
  addFace(dd?.asciiFont)
  addFace(dd?.eastAsiaFont)
  for (const info of parsed.styles.values()) {
    addFace(info.display?.font)
    addFace(info.display?.fontAscii)
    addFace(info.display?.eastAsiaFont)
  }
  for (const e of parsed.fontTable ?? []) addFace(e.name)
  const paras: unknown[] = []
  for (const runs of blockRunGroups(parsed.blocks)) {
    if (!runs || runs.length === 0) continue
    const text = runs.map((r: Run) => r.text).join('')
    for (const r of runs) {
      addFace(r.font)
      addFace(r.fontAscii)
      addFace(r.eastAsiaFont)
      addFace(r.csFont)
    }
    if (!text) continue
    const para = layoutParaOfRuns(runs, text, undefined)
    let balancedCount = 0
    if (balanced) {
      runs.forEach((r: Run, i: number) => {
        balancedCount += balancedSpaceOffsets(
          r.text,
          runs[i - 1]?.text ?? '',
          runs[i + 1]?.text ?? '',
        ).length
      })
    }
    paras.push({
      t: hash(text),
      n: text.length,
      cjk: textHasCjk(text),
      hangul: textHasHangul(text),
      lf: paraLineFactor(para),
      cls: {
        isCjk: countBy(text, isCjk),
        sf: eastAsianSegments(text).reduce((n, s) => n + s.text.length, 0),
        cps: countBy(text, (cp) => CPS_CJK_RE.test(String.fromCodePoint(cp))),
        wide: countBy(text, passesWide),
        side: countBy(text, autospaceSide),
        scale: charScaleEm(text, 50),
      },
      balanced: balancedCount,
      seams: autospaceSeams(text, true).length,
      segs: (() => {
        const all = inlineSegments({
          runs: inlineRunsOf(runs).runs,
          autospace: true,
          inheritedAscii: docBodyFont(parsed) ?? null,
          inheritedEa: bodyEa ?? null,
          wordWrap: true,
          balancedSpaces: balanced,
          justified: false,
          simsunGapLineHeight: 'x',
          lineMultiple: null,
        })
        const by: Record<string, number> = {}
        for (const sg of all.segments)
          by[sg.cls ?? sg.style ?? '?'] = (by[sg.cls ?? sg.style ?? '?'] ?? 0) + 1
        if (all.joiners.length) by.joiners = all.joiners.length
        return by
      })(),
    })
  }
  const faces: Record<string, unknown> = {}
  for (const name of [...faceNames].sort()) faces[name] = faceRecord(name)
  return { doc, faces, paras }
}

describe.skipIf(!corpus.length || !out)('inline layout snapshot', () => {
  it('records every decision over the corpus', async () => {
    const result: Record<string, unknown> = {}
    const files: string[] = []
    for (const dir of corpus) {
      for (const f of readdirSync(dir)) if (/\.docx$/i.test(f)) files.push(join(dir, f))
    }
    files.sort()
    for (const file of files) {
      try {
        result[basename(file)] = await snapshotDoc(file)
      } catch (e) {
        result[basename(file)] = { error: String(e).slice(0, 200) }
      }
    }
    writeFileSync(out, JSON.stringify(result, null, 1))
    expect(Object.keys(result).length).toBe(files.length)
  }, 600_000)
})
