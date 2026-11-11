/**
 * I9 plan == scrape: for every pagination-corpus document the model-side
 * BlockPlan must agree with what measureBlocks scrapes from the rendered
 * editor DOM (break classification, table row policy, inline space before).
 * jsdom lays nothing out, so every top-level block gets a stub rect; the
 * compared facts are class / attribute / text facts, not geometry.
 */
import { appendFileSync, readdirSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx, readSections, type Block, type ParsedDoc } from '@genoffice/docx-engine'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  imageParaPaginationMeta,
  paraPaginationMeta,
  tableWidowOff,
} from '../src/renderer/editor/para-flags'
import { measureBlocks, tableRowFlags, type BlockMeta } from '../src/renderer/pagination'
import {
  buildBlockPlan,
  comparePlanToBlocks,
  planDiffHistogram,
  planPropsOf,
  type BlockPlan,
} from '../src/renderer/pagination-plan'
import type { BlockBox } from '../src/renderer/pagination-types'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { floatTableFlipMatrixXml, floatTableMatrixXml } from './helpers/float-table-matrix'

const HERE = dirname(fileURLToPath(import.meta.url))
const CORPUS_DIR = join(HERE, 'pagination-corpus/docx')
const GEOMETRY_DIR = join(HERE, 'pagination-corpus/geometry')
const TRUTHS_DIR = join(HERE, '../../../packages/docx-engine/tests/fixtures/truths')
/** PLAN_SHADOW_CORPUS=<dir>: also run an outside corpus (reported, never pinned) */
const OUTSIDE_DIR = process.env.PLAN_SHADOW_CORPUS

const docxStems = (dir: string) =>
  readdirSync(dir)
    .filter((f) => f.endsWith('.docx'))
    .map((f) => f.slice(0, -'.docx'.length))
    .sort()
const stems = docxStems(CORPUS_DIR)
const geometryStems = docxStems(GEOMETRY_DIR)
/** truth probes carrying anchored objects / floating tables */
const ANCHOR_PROBES = [
  'anchor-align-origin',
  'float-table-spec-origin',
  'render-anchor-own-line-band',
  'render-bottom-float-section-break',
  'render-hf-anchor-reserve',
  'render-hf-side-wrap-flow',
  'render-nested-float-table',
  'render-section-start-clears-float',
]

/** the parse-layer meta App.tsx hands the slicer (footnote bands left out: not a plan field) */
function metaOfFor(parsed: ParsedDoc): (docxIndex: number) => BlockMeta | undefined {
  const byIndex = new Map<number, Block>()
  for (const b of parsed.blocks)
    if (b.docxIndex !== null && !byIndex.has(b.docxIndex)) byIndex.set(b.docxIndex, b)
  const defaultParaStyle = [...parsed.styles.values()].find(
    (s) => s.type === 'paragraph' && s.isDefault,
  )
  const styleOf = (id: string | undefined) =>
    (id ? parsed.styles.get(id) : defaultParaStyle)?.display
  return (docxIndex) => {
    const b = byIndex.get(docxIndex)
    if (!b) return undefined
    if (b.type === 'table') {
      const styleKeepNext = (id: string) => parsed.styles.get(id)?.display?.keepNext === true
      const xml = b.originalXml ?? ''
      const flagged =
        /tblHeader|cantSplit|keepNext|<w:trHeight\b/.test(xml) ||
        [...xml.matchAll(/<w:pStyle w:val="([^"]+)"/g)].some((m) => styleKeepNext(m[1]))
      const modern = (parsed.compatibilityMode ?? 0) >= 15
      if (!flagged && !modern) return undefined
      return {
        ...(flagged ? { tableRowFlags: tableRowFlags(xml, styleKeepNext) } : {}),
        ...(modern ? { modernTableHeaders: true } : {}),
        ...(modern &&
        tableWidowOff(
          xml,
          defaultParaStyle?.display?.widowControl ?? parsed.docDefaults?.widowControl,
          (id) => parsed.styles.get(id)?.display?.widowControl,
        )
          ? { cellWidowOff: true }
          : {}),
      }
    }
    return b.format
      ? paraPaginationMeta(b.format)
      : imageParaPaginationMeta(b.originalXml, styleOf, parsed.docDefaults)
  }
}

const rectOf = (top: number, height: number) =>
  ({
    top,
    height,
    bottom: top + height,
    left: 0,
    right: 500,
    width: 500,
    x: 0,
    y: top,
    toJSON: () => ({}),
  }) as DOMRect

export interface ShadowRun {
  diffs: ReturnType<typeof comparePlanToBlocks>
  blocks: number
  plans: number
  /** the switched facts, scraped (no plan) and plan-driven, per block */
  scraped: unknown[]
  planned: unknown[]
}

const SWITCHED: Array<keyof BlockBox> = [
  'breakBefore',
  'breakBeforeBr',
  'extraBreaksBefore',
  'breakAfter',
  'extraBreaksAfter',
  'leadingBreakCut',
  'hasCellBreak',
  'colBreakBefore',
  'colBreakAfter',
  'breakForce',
  'breakOnlyLineH',
  'pageRelVyPx',
  'pageRelVAnchor',
  'pageRelVSpec',
  'pageRelVApplied',
  'floatShiftPx',
  'liftPx',
  'floatPushWhole',
  'bandKeep',
  'floatFlowed',
]
const switchedFacts = (b: BlockBox) =>
  Object.fromEntries(SWITCHED.map((k) => [k, b[k]]).filter(([, v]) => v !== undefined))

/** jsdom hands back the declared length (pt or px) */
function declaredPx(css: string): number {
  const v = parseFloat(css)
  if (!Number.isFinite(v)) return 0
  return css.trim().endsWith('pt') ? (v * 96) / 72 : v
}

export async function shadowRun(bytes: Uint8Array): Promise<ShadowRun> {
  const parsed = await parseDocx(bytes)
  const sections = readSections(parsed)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks, sections) as never,
  })
  try {
    const pm = editor.view.dom as HTMLElement
    Array.from(pm.children).forEach((el, i) => {
      // a layout stand-in: Blink lays the first block out below its own margin
      const lead = i === 0 ? declaredPx(getComputedStyle(el).marginTop) : 0
      ;(el as HTMLElement).getBoundingClientRect = () => rectOf(i * 100 + lead, 50)
    })
    const { blocks } = measureBlocks(pm, 0, 1)
    const plans = buildBlockPlan(editor.state.doc, planPropsOf(parsed, sections, false))
    const byEl = new Map<HTMLElement, BlockPlan>()
    for (const p of plans.values()) {
      const dom = editor.view.nodeDOM(p.pos)
      if (dom instanceof HTMLElement) byEl.set(dom, p)
    }
    const diffs = comparePlanToBlocks(blocks, (el) => byEl.get(el), metaOfFor(parsed))
    const planOf = (el: HTMLElement) => byEl.get(el)
    const driven = measureBlocks(pm, 0, 1, planOf).blocks
    // the canvas pass: plan-driven with the scrape kept as the shadow twin
    const shadowed = measureBlocks(pm, 0, 1, planOf, true).blocks
    expect(shadowed.some((b) => b.shadowScrape)).toBe(true)
    const shadowDiffs = comparePlanToBlocks(shadowed, planOf, metaOfFor(parsed))
    return {
      diffs: [...diffs, ...shadowDiffs],
      blocks: blocks.length,
      plans: plans.size,
      scraped: blocks.map(switchedFacts),
      planned: driven.map(switchedFacts),
    }
  } finally {
    editor.destroy()
  }
}

function expectAgreement(stem: string, run: ShadowRun): void {
  expect(run.blocks).toBeGreaterThan(0)
  expect(
    run.diffs,
    `${stem}: ${JSON.stringify(planDiffHistogram(run.diffs))} e.g. ${JSON.stringify(run.diffs.slice(0, 3))}`,
  ).toEqual([])
  // the plan-driven measure yields the same block facts as the scrape
  expect(run.planned).toEqual(run.scraped)
}

describe('I9 plan == scrape over the pagination corpus', () => {
  for (const stem of stems) {
    it(stem, async () => {
      expectAgreement(
        stem,
        await shadowRun(new Uint8Array(readFileSync(join(CORPUS_DIR, `${stem}.docx`)))),
      )
    })
  }
})

describe('I9 plan == scrape: float / anchor geometry documents', () => {
  for (const stem of geometryStems) {
    it(stem, async () => {
      expectAgreement(
        stem,
        await shadowRun(new Uint8Array(readFileSync(join(GEOMETRY_DIR, `${stem}.docx`)))),
      )
    })
  }
  for (const probe of ANCHOR_PROBES) {
    it(probe, async () => {
      expectAgreement(
        probe,
        await shadowRun(new Uint8Array(readFileSync(join(TRUTHS_DIR, probe, 'probe.docx')))),
      )
    })
  }
  for (const [name, body] of Object.entries({
    'float-table-matrix': floatTableMatrixXml,
    'float-table-flip-matrix': floatTableFlipMatrixXml,
  })) {
    it(name, async () => {
      expectAgreement(name, await shadowRun(await buildDocx({ bodyXml: body(), withImage: true })))
    })
  }
})

describe.skipIf(!OUTSIDE_DIR)('I9 plan == scrape: outside corpus (reported)', () => {
  const hist: Record<string, number> = {}
  const docs: string[] = []
  for (const stem of OUTSIDE_DIR ? docxStems(OUTSIDE_DIR) : []) {
    it(
      stem,
      async () => {
        const run = await shadowRun(
          new Uint8Array(readFileSync(join(OUTSIDE_DIR!, `${stem}.docx`))),
        )
        for (const [k, v] of Object.entries(planDiffHistogram(run.diffs)))
          hist[k] = (hist[k] ?? 0) + v
        if (run.diffs.length) docs.push(stem.slice(0, 8))
        const floatDiffs = run.diffs.filter((d) => d.cls.startsWith('float.'))
        if (floatDiffs.length && process.env.PLAN_SHADOW_REPORT)
          appendFileSync(
            process.env.PLAN_SHADOW_REPORT,
            `${stem.slice(0, 8)} ${JSON.stringify(floatDiffs.slice(0, 4))}\n`,
          )
        expect(run.planned).toEqual(run.scraped)
      },
      120_000,
    )
  }
  it('summary', () => {
    const line = `plan shadow outside corpus: ${JSON.stringify(hist)} docs ${docs.length}`
    if (process.env.PLAN_SHADOW_REPORT) appendFileSync(process.env.PLAN_SHADOW_REPORT, `${line}\n`)
    else console.log(line)
  })
})
