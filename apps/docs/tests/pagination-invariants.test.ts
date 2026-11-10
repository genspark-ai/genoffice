/**
 * Pagination invariants over the pure slicing engine, run on every corpus
 * document. The engine only warns or bumps a debug counter when one of these
 * breaks; here each one fails the build.
 *
 *   I1 idempotence  — slicing the same measured input again yields identical slices
 *   I2 boundaries   — pages tile the flow: one boundary per page, no gap, no overlap
 *   I3 seam == row  — a cut inside a table lands on a row top (or carries a row split)
 *   I4 convergence  — no safety net fired: column turns advance, row splits finish
 *                     within their fragment budget, the fixed point settles under its bound
 *
 * A corpus document known to break an invariant is listed in KNOWN_FAILURES
 * with the hazard it tracks (X1..X10 in the pagination loop review); the entry
 * runs as `it.fails` so a fix retires it instead of hiding it.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { parseDocx, type ParsedDoc } from '@genoffice/docx-engine'
import {
  computeSectionedSlicesF2,
  sliceWithLineSplit,
  TABLE_SEAM_PX,
  type BlockBox,
  type PageSlice,
  type SectionGeom,
  type SliceOutputs,
} from '../src/renderer/pagination'
import { buildCorpusBlocks } from './helpers/corpus-blocks'

const CORPUS_DIR = join(dirname(fileURLToPath(import.meta.url)), 'pagination-corpus/docx')

type Invariant = 'idempotence' | 'boundaries' | 'seam' | 'convergence'

/** `${invariant}:${stem}` -> hazard id and note */
const KNOWN_FAILURES: Record<string, string> = {}

const stems = readdirSync(CORPUS_DIR)
  .filter((f) => f.endsWith('.docx'))
  .map((f) => f.slice(0, -'.docx'.length))
  .sort()

interface Sliced {
  parsed: ParsedDoc
  blocks: BlockBox[]
  geoms: SectionGeom[]
  totalHeight: number
  slices: PageSlice[]
  out: SliceOutputs
}

const freshOut = (): SliceOutputs => ({
  rowFills: [],
  rowSplits: [],
  floatVShifts: [],
  floatFlows: [],
  floatSplits: [],
  oversizeClips: [],
})

const parsedCache = new Map<string, Promise<ParsedDoc>>()
const parse = (stem: string) => {
  let p = parsedCache.get(stem)
  if (!p) {
    p = parseDocx(new Uint8Array(readFileSync(join(CORPUS_DIR, `${stem}.docx`))))
    parsedCache.set(stem, p)
  }
  return p
}

async function slice(stem: string): Promise<Sliced> {
  const parsed = await parse(stem)
  const { blocks, geoms, totalHeight } = buildCorpusBlocks(parsed)
  const out = freshOut()
  const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight, out)
  return { parsed, blocks, geoms, totalHeight, slices, out }
}

const snapshot = (slices: PageSlice[]) =>
  slices.map((s) => ({
    start: Math.round(s.start * 100) / 100,
    end: Math.round(s.end * 100) / 100,
    section: s.section,
    repeatHeader: s.repeatHeader ?? null,
    regions: s.regions?.length ?? 0,
  }))

function rowTops(table: BlockBox): number[] {
  const tops: number[] = []
  let y = table.top
  for (const r of table.tableRows ?? []) {
    tops.push(y)
    y += r.height
  }
  return tops
}

function check(inv: Invariant, stem: string, body: () => Promise<void>) {
  const known = KNOWN_FAILURES[`${inv}:${stem}`]
  if (known) it.fails(`${stem} (known: ${known})`, body)
  else it(stem, body)
}

describe('I1 idempotence: slicing the same measured input twice is identical', () => {
  for (const stem of stems) {
    check('idempotence', stem, async () => {
      const first = await slice(stem)
      const again = computeSectionedSlicesF2(
        first.blocks,
        first.geoms,
        first.totalHeight,
        freshOut(),
      )
      expect(snapshot(again)).toEqual(snapshot(first.slices))
      const rebuilt = await slice(stem)
      expect(snapshot(rebuilt.slices)).toEqual(snapshot(first.slices))
    })
  }
})

describe('I2 boundaries: pages tile the flow with exactly one boundary each', () => {
  for (const stem of stems) {
    check('boundaries', stem, async () => {
      const { slices, totalHeight } = await slice(stem)
      expect(slices.length).toBeGreaterThan(0)
      expect(slices[0].start).toBe(0)
      expect(slices[slices.length - 1].end).toBeGreaterThanOrEqual(totalHeight - 0.01)
      for (let i = 0; i < slices.length; i++) {
        const s = slices[i]
        expect(s.end, `page ${i + 1} end >= start`).toBeGreaterThanOrEqual(s.start)
        if (i > 0) {
          expect(s.start, `page ${i + 1} opens where page ${i} closed`).toBe(slices[i - 1].end)
        }
      }
      const starts = new Set(slices.slice(1).map((s) => s.start))
      expect(starts.size, 'no two page boundaries share a flow position').toBe(slices.length - 1)
    })
  }
})

describe('I3 seam == row top: a cut inside a table lands on a row boundary', () => {
  for (const stem of stems) {
    check('seam', stem, async () => {
      const { blocks, slices, out } = await slice(stem)
      const tables = blocks.filter((b) => b.tableRows?.length)
      for (let i = 1; i < slices.length; i++) {
        const y = slices[i].start
        const table = tables.find((b) => b.top < y && y < b.top + b.height - 0.01)
        if (!table) continue
        const tops = rowTops(table)
        const onRow = tops.some((t) => Math.abs(t - y) <= TABLE_SEAM_PX)
        const split = (out.rowSplits ?? []).some(
          (p) =>
            Math.abs(p.blockTop - table.top) < 0.5 &&
            p.rules.some((r) => Math.abs(table.top + r.from - y) <= TABLE_SEAM_PX),
        )
        expect(
          onRow || split,
          `page ${i + 1} opens at ${y.toFixed(2)} inside the table at ${table.top.toFixed(2)}; row tops ${tops
            .map((t) => t.toFixed(2))
            .join(' ')}`,
        ).toBe(true)
      }
    })
  }
})

describe('I4 convergence: no safety net fires', () => {
  for (const stem of stems) {
    check('convergence', stem, async () => {
      const { blocks, geoms, totalHeight, out } = await slice(stem)
      expect(out.runaway, 'column-turn limit').toBe(false)
      expect(out.rowSplitOverflows, 'row splits past the fragment budget').toBe(0)
      const fixed = freshOut()
      const slices = sliceWithLineSplit(blocks, geoms, totalHeight, 1, undefined, fixed)
      const bound = Math.max(3, Math.min(24, slices.length))
      expect(fixed.iterationsCapped, 'fixed-point loop hit its pass bound').toBe(false)
      expect(fixed.iterations ?? 0).toBeLessThanOrEqual(bound)
      expect(fixed.runaway).toBe(false)
      expect(fixed.rowSplitOverflows).toBe(0)
    })
  }
})
