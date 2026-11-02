/**
 * Header/footer strip geometry pins (geometry consolidation, stage 3).
 *
 * Four places size or draw a strip: the line-box estimate, the hidden DOM
 * probe behind `hfReservedHeightPx`, the canvas gap strip and the preview /
 * HeaderFooterArea strip. jsdom lays nothing out, so this golden pins what is
 * computable without layout for every (section, kind, variant) strip of the
 * pagination corpus, the synthetic geometry documents and the header truth
 * probes: the estimate and its floating-drawing floor, the width each source
 * lays the part out at, and the `disagree` tags naming where the sources still
 * read different inputs (the first-page/last-page edge strips draw at the
 * page's width; footer obstacles anchored to the page or margin have no known
 * strip position because the footer strip grows upward). The rendered
 * equality (canvas gap strip == preview strip) is asserted in
 * e2e/docs-pagination-invariants.spec.ts.
 */
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { parseDocx, readSections, type HeaderFooter, type HfImage } from '@genoffice/docx-engine'
import { effectiveHfRefs, hfPartWidthPx, type HfVariant } from '../src/renderer/pagination-hf'
import { estimateHfHeight, hfHeaderGeom } from '../src/renderer/line-metrics'

const HERE = dirname(fileURLToPath(import.meta.url))
const CORPUS = join(HERE, 'pagination-corpus/docx')
const GEOMETRY = join(HERE, 'pagination-corpus/geometry')
const TRUTHS = join(HERE, '../../../packages/docx-engine/tests/fixtures/truths')
const GOLDEN_DIR = join(HERE, 'fixtures/hf-strip-geometry-golden')
const UPDATE = process.env.HF_STRIP_GOLDEN_UPDATE === '1'

const docs: Array<{ name: string; file: string }> = [
  ...readdirSync(CORPUS)
    .filter((f) => f.endsWith('.docx'))
    .map((f) => ({ name: f.slice(0, -5), file: join(CORPUS, f) })),
  ...readdirSync(GEOMETRY)
    .filter((f) => f.endsWith('.docx'))
    .map((f) => ({ name: f.slice(0, -5), file: join(GEOMETRY, f) })),
  ...[
    'render-hf-anchor-reserve',
    'render-inherited-header-width',
    'parity-blank-restart',
    'render-hf-inline-picture',
    'render-hf-side-wrap-flow',
  ].map((id) => ({ name: `truth-${id}`, file: join(TRUTHS, id, 'probe.docx') })),
].sort((a, b) => a.name.localeCompare(b.name))

interface StripPin {
  section: number
  kind: 'header' | 'footer'
  variant: HfVariant
  /** width the capacity / preview lay the part out at (declaring section) */
  ownerWidthPx: number
  /** width the canvas gap strip draws it at (the page's own section) */
  pageWidthPx: number
  estimatePx: number
  /** floating-drawing floor the estimate adds on top of the DOM probe */
  floorPx: number
  floats: string[]
  disagree: string[]
}

const r1 = (n: number) => Math.round(n * 10) / 10

async function pins(file: string): Promise<StripPin[]> {
  const parsed = await parseDocx(new Uint8Array(readFileSync(file)))
  const sections = readSections(parsed)
  const refs = effectiveHfRefs(sections)
  const out: StripPin[] = []
  sections.forEach((s, i) => {
    const set = s.settings
    const pageWidthPx = ((set.pageWidth - set.marginLeft - set.marginRight) / 1440) * 96
    const variants: HfVariant[] = ['default']
    if (s.titlePg) variants.push('first')
    if (parsed.evenAndOddHeaders) variants.push('even')
    for (const kind of ['header', 'footer'] as const) {
      for (const variant of variants) {
        const rId = refs[i][kind][variant]
        const part = rId ? parsed.hfParts?.[rId] : undefined
        if (!part) continue
        const value: HeaderFooter = {
          text: part.text,
          pageNumber: part.hasPageNumber,
          paras: part.paras,
        }
        const images: HfImage[] = part.images ?? []
        const ownerWidthPx = hfPartWidthPx(sections, i, kind, variant)
        const geom =
          kind === 'header'
            ? { ...hfHeaderGeom(set), marginLeftPx: (set.marginLeft / 1440) * 96 }
            : undefined
        const floats = images
          .filter((im) => im.floating)
          .map(
            (im) =>
              `${im.wrap ?? 'none'}${im.behind ? '/behind' : ''}@${im.posHRel ?? '-'}:${im.posVRel ?? '-'}`,
          )
        // inputs the sources still read differently after the stage-3 unification
        const disagree: string[] = []
        if (Math.abs(ownerWidthPx - pageWidthPx) > 0.5) disagree.push('inherited-width-edge-strip')
        if (
          kind === 'footer' &&
          images.some(
            (im) =>
              im.floating &&
              !im.behind &&
              im.wrap &&
              im.wrap !== 'none' &&
              im.wrap !== 'topBottom' &&
              im.posVRel !== 'paragraph',
          )
        )
          disagree.push('footer-page-anchored-obstacle-not-shaped')
        out.push({
          section: i,
          kind,
          variant,
          ownerWidthPx: r1(ownerWidthPx),
          pageWidthPx: r1(pageWidthPx),
          estimatePx: r1(estimateHfHeight(value, ownerWidthPx, images, geom)),
          floorPx: r1(estimateHfHeight(null, ownerWidthPx, images, geom)),
          floats,
          disagree,
        })
      }
    }
  })
  return out
}

describe('header/footer strip geometry golden', () => {
  const tally: Record<string, number> = {}
  for (const { name, file } of docs) {
    it(name, async () => {
      const got = await pins(file)
      for (const p of got) for (const d of p.disagree) tally[d] = (tally[d] ?? 0) + 1
      const path = join(GOLDEN_DIR, `${name}.json`)
      if (UPDATE || !existsSync(path)) {
        mkdirSync(GOLDEN_DIR, { recursive: true })
        writeFileSync(path, `${JSON.stringify(got, null, 2)}\n`)
      }
      expect(got).toEqual(JSON.parse(readFileSync(path, 'utf8')))
    })
  }

  it('lists where the four strip sources read different inputs', () => {
    const path = join(GOLDEN_DIR, '_disagreements.json')
    const got = Object.fromEntries(Object.entries(tally).sort())
    if (UPDATE || !existsSync(path)) writeFileSync(path, `${JSON.stringify(got, null, 2)}\n`)
    expect(got).toEqual(JSON.parse(readFileSync(path, 'utf8')))
  })
})
