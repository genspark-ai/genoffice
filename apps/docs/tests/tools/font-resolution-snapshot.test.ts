/**
 * Font-resolution decision snapshot over a docx corpus
 * (tools/docs-font-resolution-snapshot.mjs). For every distinct declared font
 * name and every distinct run slot pair of each document it records what the
 * face / line-factor / chain functions answer today, so the font-rules
 * consolidation waves can prove "same decisions" without a corpus gate.
 * Skipped unless FONT_SNAPSHOT_CORPUS names the corpus directories.
 */
import { readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import type { Run } from '@genoffice/docx-engine'
import { blockRunGroups } from '../../../../packages/docx-engine/src/parse-block-passes'
import * as lineMetrics from '../../src/renderer/line-metrics'
import * as fontRules from '../../src/renderer/text-layout/font-rules'
import { documentTextContext } from '../../src/renderer/text-layout/document-context'
import { setOpenDocumentText } from '../../src/renderer/text-layout/open-document'

const lm = { ...lineMetrics, ...fontRules }

const corpus = (process.env.FONT_SNAPSHOT_CORPUS ?? '').split(',').filter(Boolean)
const out = process.env.FONT_SNAPSHOT_OUT ?? ''

// the entry point under construction: absent on the baseline checkout
type Decide = (name: string) => unknown
async function loadDecide(): Promise<Decide | null> {
  // a variable path keeps vite from resolving the module at transform time
  const path = '../../src/renderer/text-layout/font-rules'
  try {
    const mod = (await import(/* @vite-ignore */ path)) as { resolveFace?: Decide }
    return mod.resolveFace ?? null
  } catch {
    return null
  }
}

function faceRecord(name: string) {
  return {
    lf: lm.lineHeightFactor(name),
    tableLf: lm.tableLineHeightFactor(name),
    cjkLf: lm.cjkDeclaredLineFactor(name),
    krLf: lm.krLineFactor(name),
    latinLf: lm.latinLineFactor(name),
    simsunGap: lm.simsunGapLineFactor(name),
    kr: lm.krCharsetRoute(name),
    eaCharset: lm.installedEaCharsetFace(name),
    eaLayout: lm.isEaLayoutFace(name),
    koLayout: lm.isKoreanLayoutFace(name),
    krBatangAlt: lm.krBatangAltNamed(name),
    krName: lm.krNameLineFactor(name),
    missingLatin: lm.missingLatinSubstitute(name),
    alias: lm.wordMetricAliasFace(name),
    bundled: lm.isBundledFont(name),
    missing: lm.declaredFaceMissing(name),
    gap: lm.latinAutospaceGapEm(name),
    chain: lm.cssFontFamily(name),
    standIn: lm.standInFontFamily(name),
    eaOnly: lm.cssEaOnlyFontFamily(name),
    latinChain: lm.docLatinChainCss(name),
    markerLift: lm.markerFaceLifts(name),
    oversized: lm.oversizedMarkerExtraPt(name, 14, 12),
  }
}

async function snapshotDoc(path: string, decide: Decide | null) {
  const parsed = await parseDocx(new Uint8Array(readFileSync(path)))
  lm.setDocFontTable(parsed.fontTable ?? [])
  const embedded = (parsed.embeddedFonts ?? [])
    .filter((f) => f.lineMetrics)
    .map((f) => ({ family: f.family, ...f.lineMetrics!, styled: f.bold || f.italic }))
    .sort((a, b) => Number(a.styled) - Number(b.styled))
  lm.setEmbeddedLineMetrics(embedded)
  setOpenDocumentText(documentTextContext(parsed))
  const dd = parsed.docDefaults
  const names = new Set<string>()
  const add = (n: string | undefined | null) => {
    if (n) names.add(n)
  }
  add(dd?.asciiFont)
  add(dd?.eastAsiaFont)
  for (const info of parsed.styles.values()) {
    add(info.display?.font)
    add(info.display?.fontAscii)
    add(info.display?.eastAsiaFont)
  }
  for (const e of parsed.fontTable ?? []) add(e.name)
  const pairs = new Map<string, { ascii: string | null; ea: string | null }>()
  const csTriples = new Map<string, { cs: string; ascii?: string; ea?: string }>()
  for (const runs of blockRunGroups(parsed.blocks)) {
    for (const r of runs ?? []) {
      const run = r as Run
      add(run.font)
      add(run.fontAscii)
      add(run.eastAsiaFont)
      add(run.fontCs)
      const ascii = run.fontAscii ?? null
      const ea = run.eastAsiaFont ?? run.font ?? null
      if (ascii || ea) pairs.set(`${ascii} | ${ea}`, { ascii, ea })
      if (run.fontCs)
        csTriples.set(`${run.fontCs} | ${ascii} | ${ea}`, {
          cs: run.fontCs,
          ascii: ascii ?? undefined,
          ea: ea ?? undefined,
        })
    }
  }
  const faces: Record<string, unknown> = {}
  for (const name of [...names].sort()) faces[name] = faceRecord(name)
  const decisions: Record<string, unknown> = {}
  if (decide) for (const name of [...names].sort()) decisions[name] = decide(name)
  const runPairs: Record<string, unknown> = {}
  for (const [k, { ascii, ea }] of [...pairs].sort()) {
    runPairs[k] = {
      ff: ascii && ea ? lm.cssRunFontFamily(ascii, ea) : ea ? lm.cssEaOnlyFontFamily(ea) : null,
      slots: lm.wordSlotFaces(ascii, ea),
      run: lm.runSlotFaces({ fontAscii: ascii, eastAsiaFont: ea, font: ea ?? ascii }),
    }
  }
  const cs: Record<string, string> = {}
  for (const [k, t] of [...csTriples].sort()) cs[k] = lm.cssCsFontFamily(t.cs, t.ascii, t.ea)
  return {
    doc: {
      fontTable: (parsed.fontTable ?? []).map((e) => ({
        name: e.name,
        altName: e.altName ?? null,
        charset: e.charset ?? null,
        family: e.family ?? null,
        panose: e.panose ?? null,
        csb0: e.csb0 ?? null,
      })),
      embedded: embedded.map((e) => e.family),
      inheritedSlot: lm.inheritedSlotFace(),
      inheritedMark: lm.inheritedMarkFace(dd?.asciiFont ?? null),
      latinViaEa: lm.latinViaEastAsiaSlotOn(),
    },
    faces,
    runPairs,
    cs,
    ...(decide ? { decisions } : {}),
  }
}

describe.skipIf(!corpus.length || !out)('font resolution snapshot', () => {
  it('records every face decision over the corpus', async () => {
    const decide = await loadDecide()
    const result: Record<string, unknown> = {}
    const files: string[] = []
    for (const dir of corpus) {
      for (const f of readdirSync(dir)) if (/\.docx$/i.test(f)) files.push(join(dir, f))
    }
    files.sort()
    for (const file of files) {
      const key = `${basename(join(file, '..'))}/${basename(file)}`
      try {
        result[key] = await snapshotDoc(file, decide)
      } catch (e) {
        result[key] = { error: String(e).slice(0, 200) }
      }
    }
    writeFileSync(out, JSON.stringify(result, null, 1))
    expect(Object.keys(result).length).toBe(files.length)
  }, 1_200_000)
})
