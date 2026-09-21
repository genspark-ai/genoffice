import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import {
  addMedia,
  createBlankPptx,
  elementSpid,
  openPptx,
  savePptx,
  setSlideAnimations,
} from '@genoffice/pptx-engine'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * A deck whose video starts "In Click Sequence" (PowerPoint's default for an inserted
 * video, written as a mediacall playFrom step): the show must show the video paused on
 * entry and the first click must start it — the media call is an animation step, not
 * an entrance effect that hides the shape and eats the click.
 */
async function clickSequenceVideoDeck(): Promise<string> {
  const opened = await openPptx(await createBlankPptx())
  const EMU = 914400
  const r = addMedia(opened, 0, {
    kind: 'video',
    bytes: new Uint8Array(await readFile(resolve(__dirname, 'assets/click-sequence-video.mp4'))),
    ext: 'mp4',
    offset: { x: 2 * EMU, y: 1.2 * EMU, cx: 6 * EMU, cy: 3.375 * EMU },
    name: 'Clip',
  })
  if (!r) throw new Error('addMedia failed')
  const slide = opened.deck.slides[0]!
  const spid = elementSpid(slide.elements.find((e) => e.id === r.elementId)!)!
  setSlideAnimations(slide, [
    {
      spid,
      effect: 'mediaPlay',
      trigger: 'onClick',
      durationMs: 0,
      delayMs: 0,
      mediaKind: 'video',
    },
  ])
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-media-show-'))
  const file = join(dir, 'click-video.pptx')
  await writeFile(file, await savePptx(opened))
  return file
}

test('click-sequence video plays on the first click of the show', async () => {
  const deck = await clickSequenceVideoDeck()
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'slides-show-media-play',
    openFile: deck,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://slides/')
    await page.waitForSelector('.stage-wrap canvas', { timeout: 30_000 })
    const anims = await page.evaluate(() =>
      (
        window as unknown as { slidesApi: { getAnimations(i: number): Promise<unknown[]> } }
      ).slidesApi.getAnimations(0),
    )
    expect(anims).toMatchObject([{ effect: 'mediaPlay', trigger: 'onClick' }])

    await page.locator('.stage-wrap').click({ position: { x: 5, y: 5 } })
    await page.keyboard.press('F5')
    const video = page.locator('.slideshow video')
    await expect(video).toBeVisible({ timeout: 20_000 })
    expect(await video.evaluate((v: HTMLVideoElement) => v.paused)).toBe(true)

    // First click = the media call step: the video starts and the page does not turn
    await page.locator('.slideshow').click({ position: { x: 8, y: 8 } })
    await expect
      .poll(() => video.evaluate((v: HTMLVideoElement) => !v.paused && v.currentTime > 0), {
        timeout: 10_000,
      })
      .toBe(true)
    await expect(page.locator('.slideshow .ss-counter')).toHaveText('1 / 1')
    await page.keyboard.press('Escape')
  } finally {
    await closeAndSaveVideo(launched, 'slides-show-media-play')
  }
})
