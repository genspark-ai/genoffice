/**
 * Recorded narration (Slide Show > Record): the audio picture, its timing and presProps'
 * showNarration, as PowerPoint writes them, through save, reopen, re-record and clear.
 */
import { readFileSync } from 'node:fs'
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { validatePptx, xmllintAvailable } from '../../../tools/ooxml-validate/validate-pptx.mjs'
import {
  addNarration,
  clearNarration,
  addElement,
  createBlankPptx,
  getSlideAnimations,
  narrationSpids,
  openPptx,
  savePptx,
  setShowNarration,
  setSlideAnimations,
  NARRATION_NAME,
} from '../src/index'

const schemaGate = xmllintAvailable() || !!process.env.CI
const FIXTURE = new URL('./fixtures/01_standard_business.pptx', import.meta.url)

/** A 16-bit mono WAV of `ms` milliseconds of silence. */
function wav(ms: number, rate = 8000): Uint8Array {
  const samples = Math.round((rate * ms) / 1000)
  const buf = Buffer.alloc(44 + samples * 2)
  buf.write('RIFF', 0)
  buf.writeUInt32LE(36 + samples * 2, 4)
  buf.write('WAVEfmt ', 8)
  buf.writeUInt32LE(16, 16)
  buf.writeUInt16LE(1, 20)
  buf.writeUInt16LE(1, 22)
  buf.writeUInt32LE(rate, 24)
  buf.writeUInt32LE(rate * 2, 28)
  buf.writeUInt16LE(2, 32)
  buf.writeUInt16LE(16, 34)
  buf.write('data', 36)
  buf.writeUInt32LE(samples * 2, 40)
  return new Uint8Array(buf)
}

const narration = (ms: number) => ({ bytes: wav(ms), ext: 'wav', durationMs: ms })

async function reopen(bytes: Uint8Array) {
  const opened = await openPptx(bytes)
  return { opened, zip: await JSZip.loadAsync(bytes) }
}

const mediaParts = (zip: JSZip) =>
  Object.keys(zip.files).filter((p) => /^ppt\/media\/media\d+\./.test(p))

describe('recorded narration', () => {
  it('stores a hidden audio picture that plays from the slide start, as PowerPoint does', async () => {
    const opened = await openPptx(await createBlankPptx())
    const added = addNarration(opened, 0, narration(1500))
    expect(added).not.toBeNull()
    setShowNarration(opened)
    const saved = await savePptx(opened)
    const { opened: back, zip } = await reopen(saved)

    const slide = back.deck.slides[0]!
    const [spid] = narrationSpids(slide)
    expect(spid).toBeGreaterThan(0)
    const pic = slide.elements.find((el) => el.anchor.originalXml.includes(`id="${spid}"`))!
    expect(pic.anchor.originalXml).toContain(`name="${NARRATION_NAME}"`)
    expect(pic.transform.offset.x + pic.transform.offset.cx).toBe(back.deck.size.cx)
    expect(pic.transform.offset.y + pic.transform.offset.cy).toBe(back.deck.size.cy)

    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    expect(xml).toMatch(/<p:audio isNarration="1"><p:cMediaNode showWhenStopped="0"/)
    expect(xml).toContain('cmd="playFrom(0.0)"')
    expect(getSlideAnimations(slide)[0]).toMatchObject({
      spid,
      effect: 'mediaPlay',
      trigger: 'afterPrev',
    })

    expect(mediaParts(zip)).toEqual(['ppt/media/media1.wav'])
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
      '<Default Extension="wav" ContentType="audio/wav"/>',
    )
    const props = await zip.file('ppt/presProps.xml')!.async('string')
    expect(props).toContain('<p:showPr showNarration="1" useTimings="1"/>')
    expect(await zip.file('ppt/_rels/presentation.xml.rels')!.async('string')).toContain(
      'Target="presProps.xml"',
    )
    if (schemaGate) expect(await validatePptx(saved)).toEqual([])
  })

  it("plays before the slide's own animations, which it keeps", async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const rect = addElement(slide, { kind: 'rect', offset: { x: 0, y: 0, cx: 100, cy: 100 } })
    const rectSpid = Number(/<p:cNvPr id="(\d+)"/.exec(rect!.anchor.originalXml)![1])
    setSlideAnimations(slide, [
      { spid: rectSpid, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ])
    addNarration(opened, 0, narration(800))
    addNarration(opened, 0, narration(900))
    const anims = getSlideAnimations(opened.deck.slides[0]!)
    expect(anims.map((a) => [a.effect, a.trigger])).toEqual([
      ['mediaPlay', 'afterPrev'],
      ['fade', 'onClick'],
    ])
  })

  it('replaces the previous narration of a slide, media part included', async () => {
    const opened = await openPptx(await createBlankPptx())
    addNarration(opened, 0, narration(500))
    addNarration(opened, 0, narration(700))
    const { opened: back, zip } = await reopen(await savePptx(opened))
    expect(narrationSpids(back.deck.slides[0]!)).toHaveLength(1)
    expect(mediaParts(zip)).toHaveLength(1)
  })

  it('clears a narration: picture, timing and media part', async () => {
    const opened = await openPptx(await createBlankPptx())
    addNarration(opened, 0, narration(500))
    expect(clearNarration(opened, 0)).toBe(1)
    const { opened: back, zip } = await reopen(await savePptx(opened))
    const slide = back.deck.slides[0]!
    expect(narrationSpids(slide)).toEqual([])
    expect(slide.elements.some((el) => el.anchor.originalXml.includes(NARRATION_NAME))).toBe(false)
    expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).not.toContain('<p:audio')
    expect(mediaParts(zip)).toEqual([])
  })

  it("sets showNarration in an existing presProps and leaves the deck's other parts untouched", async () => {
    const original = readFileSync(FIXTURE)
    const before = await JSZip.loadAsync(original)
    const opened = await openPptx(original)
    addNarration(opened, 0, narration(600))
    setShowNarration(opened)
    const saved = await savePptx(opened)
    const after = await JSZip.loadAsync(saved)

    const props = await after.file('ppt/presProps.xml')!.async('string')
    expect(props).toMatch(/<p:showPr showNarration="1" useTimings="1"\/><p:extLst>/)
    expect(props).toContain('p14:discardImageEditData')

    const touched = new Set([
      'ppt/slides/slide1.xml',
      'ppt/slides/_rels/slide1.xml.rels',
      'ppt/presProps.xml',
      '[Content_Types].xml',
    ])
    for (const path of Object.keys(before.files)) {
      if (touched.has(path) || before.files[path]!.dir) continue
      const a = await before.file(path)!.async('uint8array')
      const b = await after.file(path)?.async('uint8array')
      expect(b && Buffer.compare(Buffer.from(a), Buffer.from(b)), path).toBe(0)
    }
    if (schemaGate) expect(await validatePptx(saved)).toEqual([])
  })

  it('updates the flags of an existing showPr and keeps its content', async () => {
    const opened0 = await openPptx(await createBlankPptx())
    setShowNarration(opened0)
    const withProps = await JSZip.loadAsync(await savePptx(opened0))
    const showPr = '<p:showPr loop="1" showNarration="0"><p:present/><p:sldAll/></p:showPr>'
    const props = (await withProps.file('ppt/presProps.xml')!.async('string')).replace(
      /<p:showPr[^>]*\/>/,
      showPr,
    )
    withProps.file('ppt/presProps.xml', props)
    const opened = await openPptx(await withProps.generateAsync({ type: 'uint8array' }))
    setShowNarration(opened)
    const saved = await JSZip.loadAsync(await savePptx(opened))
    expect(await saved.file('ppt/presProps.xml')!.async('string')).toContain(
      '<p:showPr useTimings="1" loop="1" showNarration="1"><p:present/><p:sldAll/></p:showPr>',
    )
  })
})
