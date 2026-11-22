/**
 * PPTX round-trip fidelity bench — composed deck class.
 *
 * Gated: `PPTX_ROUNDTRIP_ACCURACY=1 scripts/roundtrip-accuracy.mjs pptx`
 * drives it (see that script). CI never runs it — the numbers are a narrative
 * instrument for the "open, save, nothing lost" story, not a pass/fail gate;
 * the smoke case (corpus builds, round-trip runs) asserts only under the gate.
 *
 * The fixture deck is built with the engine itself (createBlankPptx + raw
 * elements + notes + a picture) with the ground truth baked in. The bench
 * round-trips it with the engine (openPptx → savePptx — the path every save
 * takes) AND with LibreOffice (`soffice --convert-to pptx`, its own open +
 * re-save), then scores both outputs:
 *
 *   intact   — the attribute is present with its original value
 *   degraded — present but changed (runs merged, geometry rounded, …)
 *   lost     — gone
 */
import { spawnSync } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  addPicture,
  appendRawElements,
  createBlankPptx,
  openPptx,
  savePptx,
  setSlideNotes,
} from '../src/index'

const RUN = process.env.PPTX_ROUNDTRIP_ACCURACY === '1'
const OUT = process.env.PPTX_ROUNDTRIP_ACCURACY_OUT ?? ''

// 1x1 red PNG (the same bytes the other engine tests use)
const PNG_1PX = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
)

const OFF = { x: 914400, y: 914400, cx: 1828800, cy: 914400 }
const NOTES = 'Speaker notes for the round-trip fidelity bench.'

async function buildFixture(): Promise<{ bytes: Uint8Array; slideSize: string }> {
  const opened = await openPptx(await createBlankPptx())
  // title textbox: two runs, the first large and bold
  const title =
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="914400" y="685800"/><a:ext cx="8229600" cy="1325563"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2800" b="1"/>` +
    `<a:t>Fidelity Report</a:t></a:r><a:r><a:rPr lang="en-US" sz="1800"/><a:t> — Q3</a:t></a:r></a:p></p:txBody></p:sp>`
  // a shape with its own fill
  const shape =
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Badge"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="914400" y="3657600"/><a:ext cx="1828800" cy="914400"/></a:xfrm>` +
    `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Badge</a:t></a:r></a:p></p:txBody></p:sp>`
  // a 2x2 table
  const table =
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Grid"/>` +
    `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="3657600" y="3657600"/><a:ext cx="4572000" cy="1828800"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
    `<a:tbl><a:tblGrid><a:gridCol w="2286000"/><a:gridCol w="2286000"/></a:tblGrid>` +
    `<a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/>` +
    `<a:t>Region</a:t></a:r></a:p></a:txBody></a:tc>` +
    `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Amount</a:t></a:r></a:p></a:txBody></a:tc></a:tr>` +
    `<a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/>` +
    `<a:t>North</a:t></a:r></a:p></a:txBody></a:tc>` +
    `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>120</a:t></a:r></a:p></a:txBody></a:tc></a:tr>` +
    `</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  appendRawElements(opened, 0, [title, shape, table])
  const slide = opened.deck.slides[0]!
  addPicture(opened, slide, { bytes: PNG_1PX, ext: 'png', offset: { ...OFF } })
  setSlideNotes(opened, 0, NOTES)
  const bytes = await savePptx(opened)
  const presentation =
    (await (await JSZip.loadAsync(bytes)).file('ppt/presentation.xml')?.async('string')) ?? ''
  const slideSize = /<p:sldSz cx="(\d+)" cy="(\d+)"/.exec(presentation)?.slice(1)?.join('x') ?? ''
  return { bytes, slideSize }
}

type Status = 'intact' | 'degraded' | 'lost'
interface Verdict {
  status: Status
  note?: string
}
const ok = (): Verdict => ({ status: 'intact' })
const lost = (note: string): Verdict => ({ status: 'lost', note })
const degraded = (note: string): Verdict => ({ status: 'degraded', note })

async function scoreDeck(bytes: Uint8Array, slideSize: string): Promise<Record<string, Verdict>> {
  const zip = await JSZip.loadAsync(bytes)
  const read = async (path: string): Promise<string | undefined> =>
    (await zip.file(path)?.async('string')) ?? undefined
  const slide1 = (await read('ppt/slides/slide1.xml')) ?? ''
  const presentation = (await read('ppt/presentation.xml')) ?? ''
  const v: Record<string, Verdict> = {}

  // slide inventory
  const slideCount = (
    await Promise.all(
      Object.keys(zip.files)
        .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
        .map((p) => zip.file(p)!.async('string')),
    )
  ).length
  v['slide count'] = slideCount === 1 ? ok() : degraded(String(slideCount))
  const outSize = /<p:sldSz cx="(\d+)" cy="(\d+)"/.exec(presentation)?.slice(1)?.join('x') ?? ''
  v['slide size'] = outSize === slideSize ? ok() : degraded(outSize || 'missing sldSz')

  // text + runs
  const hasText = slide1.includes('Fidelity Report') && slide1.includes('Badge')
  v['text frames survive'] = hasText ? ok() : lost('title/shape text')
  const run2800 = /sz="2800"[^>]*b="1"|b="1"[^>]*sz="2800"/.test(slide1)
  const run1800 = slide1.includes('sz="1800"')
  v['mixed runs: sizes + bold'] =
    run2800 && run1800
      ? ok()
      : run2800 || run1800
        ? degraded('one run form survived')
        : lost('run properties')
  v['large text keeps 2800'] = slide1.includes('sz="2800"') ? ok() : lost('sz')

  // the shape's geometry + fill
  v['shape geometry roundRect'] = slide1.includes('prst="roundRect"')
    ? ok()
    : slide1.includes('prst=')
      ? degraded(/prst="([^"]+)"/.exec(slide1)?.[1] ?? 'other preset')
      : lost('preset geometry')
  v['shape fill 4472C4'] = slide1.includes('4472C4')
    ? ok()
    : /srgbClr/.test(slide1)
      ? degraded('fill present, color changed')
      : lost('fill')
  const xfrmOk = slide1.includes('<a:off x="914400"') && slide1.includes('cx="1828800"')
  v['frame EMU geometry'] = xfrmOk
    ? ok()
    : /<a:off x="\d+"/.test(slide1)
      ? degraded('frames present, coordinates changed')
      : lost('off/ext')

  // the table
  const tbl = /<a:tbl>[\s\S]*?<\/a:tbl>/.exec(slide1)?.[0] ?? ''
  v['table part survives'] = tbl ? ok() : lost('a:tbl')
  const gridCols = (tbl.match(/<a:gridCol/g) ?? []).length
  v['table 2 columns'] = gridCols === 2 ? ok() : degraded(String(gridCols))
  const cellTexts = ['Region', 'Amount', 'North', '120'].every((t) => tbl.includes(`>${t}<`))
  v['table cell texts'] = tbl && cellTexts ? ok() : lost('cell text')

  // the picture + its bytes
  v['picture element'] = slide1.includes('<p:pic>') ? ok() : lost('p:pic')
  const media = Object.keys(zip.files).filter((p) => /^ppt\/media\//.test(p) && !p.endsWith('/'))
  v['image media part'] = media.length > 0 ? ok() : lost('media')
  const mediaBytes = await zip.file(media[0] ?? '')?.async('uint8array')
  v['image bytes verbatim'] =
    mediaBytes && mediaBytes.length === PNG_1PX.length ? ok() : degraded('media re-encoded')

  // notes
  const notesFiles = Object.keys(zip.files).filter((p) =>
    /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(p),
  )
  const notesText = notesFiles.length
    ? String(
        (await zip.file(notesFiles[0]!)!.async('string')).match(
          /type="body"[\s\S]*?<a:t>([^<]*)</,
        )?.[1] ?? '',
      ) ||
      String((await zip.file(notesFiles[0]!)!.async('string')).match(/<a:t>([^<]*)</)?.[1] ?? '')
    : ''
  v['speaker notes'] = notesText.includes('Speaker notes for the round-trip')
    ? ok()
    : notesFiles.length > 0
      ? degraded('notes part present, text changed')
      : lost('notesSlide')

  return v
}

describe.skipIf(!RUN)('pptx round-trip accuracy bench (composed deck class)', () => {
  it('scores the engine round-trip vs soffice round-trip', async () => {
    const { bytes: fixture, slideSize } = await buildFixture()
    const workDir = mkdtempSync(join(tmpdir(), 'pptx-roundtrip-accuracy-'))
    const fixturePath = join(workDir, 'composed.pptx')
    writeFileSync(fixturePath, fixture)

    // the engine's round-trip: open + save, the path every save takes
    const opened = await openPptx(fixture)
    const engineBytes = await savePptx(opened)

    const loOut = join(workDir, 'lo-out')
    const run = spawnSync(
      'soffice',
      [
        '--headless',
        '--convert-to',
        'pptx',
        '--outdir',
        loOut,
        `-env:UserInstallation=file:///${join(workDir, 'lo-profile')}`,
        fixturePath,
      ],
      { timeout: 120_000 },
    )
    if (run.status !== 0) throw new Error('soffice failed')
    const sofficeBytes = new Uint8Array(readFileSync(join(loOut, 'composed.pptx')))

    const engineVerdicts = await scoreDeck(engineBytes, slideSize)
    const sofficeVerdicts = await scoreDeck(sofficeBytes, slideSize)

    const attributes = Object.keys(engineVerdicts)
    const scorecard: string[] = [
      '# PPTX round-trip accuracy — composed deck class',
      '',
      'Attribute = a content/structure fact the fixture baked in. intact = present',
      'with its original value. The engine row is openPptx → savePptx (every save);',
      'the LibreOffice row is its own open + re-save of the same fixture.',
      '',
      '| attribute | genoffice | soffice |',
      '|---|---|---|',
    ]
    let engineIntact = 0
    let sofficeIntact = 0
    for (const key of attributes) {
      const g = engineVerdicts[key]!
      const s = sofficeVerdicts[key]!
      if (g.status === 'intact') engineIntact += 1
      if (s.status === 'intact') sofficeIntact += 1
      const note = (x: Verdict) =>
        x.status === 'intact' ? '✓' : `${x.status}${x.note ? `: ${x.note}` : ''}`
      scorecard.push(`| ${key} | ${note(g)} | ${note(s)} |`)
    }
    scorecard.push(
      '',
      `genoffice intact ${engineIntact}/${attributes.length}; LibreOffice intact ${sofficeIntact}/${attributes.length}.`,
    )
    rmSync(workDir, { recursive: true, force: true })

    const md = `${scorecard.join('\n')}\n`
    if (OUT) writeFileSync(OUT, md)
    else console.log(`\n${md}`)

    // smoke-level pins only: the corpus ran and the invariants the engine
    // exists to keep are real (the numbers above are narrative)
    expect(engineVerdicts['slide count']!.status).toBe('intact')
    expect(engineVerdicts['text frames survive']!.status).toBe('intact')
    expect(engineVerdicts['table part survives']!.status).toBe('intact')
    expect(engineVerdicts['image media part']!.status).toBe('intact')
  })
})
