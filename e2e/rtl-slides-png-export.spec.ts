import { test, expect, _electron as electron } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import { cp, mkdtemp, readFile, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { createRequire } from 'node:module'

/**
 * An RTL UI must not re-author a deck that gets exported (genoffice#1861).
 *
 * The six document surfaces pin `direction: ltr`, but Konva hosts in Slides sit
 * outside the pinned ones: the thumbnail stages in `.slide-list`, the slideshow,
 * the master view, and the offscreen containers `export-render.tsx` and
 * `selection-image.tsx` append to `document.body` for every raster export. A
 * `div` on the body inherits `dir` from <html>, and a canvas inside it draws
 * with that direction — measured, a canvas on a bare body div under `ar` paints
 * different pixels than the same canvas in a `direction: ltr` div.
 *
 * So this asserts the export bytes, not a computed style: a style assertion
 * cannot see hosts that were never pinned in the first place.
 *
 * ## Why there is a control, and what it does and does not prove
 *
 * A byte-equality test that passes for the wrong reason is worse than none, so
 * each run also draws the same string twice — once into a bare body div and once
 * into a `direction: ltr` div — and *requires those two to differ*. Under `ar`
 * that control is the whole point: it proves this harness can see a direction
 * difference, so the deck export coming out equal means equal, not "cannot tell".
 *
 * What it does **not** prove: that `NodeBody.tsx`'s `direction={g.direction ?? 'ltr'}`
 * is what makes it pass. Reverting that to `'inherit'`, and even forcing
 * `direction="rtl"` outright, leaves both exports byte-identical on the selection
 * -copy path this fixture drives — so the pin is not exercised here. The control
 * is kept because it makes the assertion honest; the line itself is argued in the
 * PR comment, not by this test.
 */

interface Run {
  bytes: Buffer
  /** the control: same text, bare body div vs a pinned-ltr div, under this UI language */
  controlDiffers: boolean
  hostDir: string
}

/** One deck, one language: the PNG bytes a selection copy puts on the clipboard. */
async function run(lang: string): Promise<Run> {
  const dir = await mkdtemp(join(tmpdir(), `genoffice-rtl-png-${lang}-`))
  try {
    const fixture = join(dir, 'fixture')
    await cp(resolve('e2e/assets/font-manager-rubik'), fixture, { recursive: true })
    const xmlPath = join(fixture, 'ppt/slides/slide1.xml')
    // the bundled face is not installed here, and pinning one face keeps the two
    // runs from differing by whatever font the machine happens to resolve
    await writeFile(xmlPath, (await readFile(xmlPath, 'utf8')).replaceAll('Rubik', 'Arial'))
    const pptx = join(dir, 'deck.pptx')
    execFileSync('zip', ['-X', '-q', '-r', pptx, '.'], { cwd: fixture })

    const require = createRequire(resolve('apps/slides/package.json'))
    const { ELECTRON_RUN_AS_NODE: _node, ...env } = process.env
    const app = await electron.launch({
      executablePath: require('electron'),
      args: [
        ...(process.platform === 'linux' ? ['--no-sandbox', '--disable-gpu'] : []),
        resolve('apps/slides'),
        pptx,
      ],
      env: { ...env, GENOFFICE_USER_DATA: join(dir, 'user-data'), GENOFFICE_LANG: lang },
    })
    try {
      const page = await app.firstWindow()
      await expect(page.locator('html')).toHaveAttribute('dir', lang === 'ar' ? 'rtl' : 'ltr')
      await expect(page.locator('.stage-wrap canvas').first()).toBeVisible()

      const control = await page.evaluate(() => {
        const paint = (dir: string | null): string => {
          const host = document.createElement('div')
          host.style.cssText = `position:fixed;left:-100000px;top:0;${dir ? `direction:${dir};` : ''}`
          document.body.appendChild(host)
          const canvas = document.createElement('canvas')
          canvas.width = 320
          canvas.height = 80
          host.appendChild(canvas)
          const ctx = canvas.getContext('2d')!
          ctx.font = '24px sans-serif'
          ctx.fillStyle = '#000'
          ctx.fillText('Hello, world! (rtl) 123', 10, 44)
          return canvas.toDataURL()
        }
        const bare = paint(null)
        const pinned = paint('ltr')
        return {
          differs: bare !== pinned,
          hostDir: getComputedStyle(document.body).direction,
        }
      })

      await page
        .locator('.stage-wrap')
        .first()
        .click({ position: { x: 10, y: 10 } })
      const mod = process.platform === 'darwin' ? 'Meta' : 'Control'
      await page.keyboard.press(`${mod}+a`)
      await page.keyboard.press(`${mod}+c`)
      await expect
        .poll(() => app.evaluate(({ clipboard }) => !clipboard.readImage().isEmpty()))
        .toBe(true)
      const bytes = Buffer.from(
        await app.evaluate(({ clipboard }) => clipboard.readImage().toPNG().toString('base64')),
        'base64',
      )
      return { bytes, controlDiffers: control.differs, hostDir: control.hostDir }
    } finally {
      const child = app.process()
      const timer = setTimeout(() => child.kill('SIGKILL'), 10000)
      try {
        await app.close()
      } finally {
        clearTimeout(timer)
      }
    }
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

test('a PNG export under ar is byte-equal to the same export under en', async ({
  browserName: _browserName,
}, info) => {
  test.setTimeout(300_000)

  const en = await run('en')
  expect(en.bytes.length).toBeGreaterThan(0)
  await info.attach('export-en.png', { body: en.bytes, contentType: 'image/png' })

  const ar = await run('ar')
  await info.attach('export-ar.png', { body: ar.bytes, contentType: 'image/png' })

  // The control first: if this harness cannot see a direction difference at all,
  // the equality below would be vacuous, and a vacuous test is worse than none.
  expect(
    ar.controlDiffers,
    'the bare-canvas control did not differ under ar, so this harness cannot ' +
      'observe direction and the byte comparison below would prove nothing',
  ).toBe(true)

  // Not `toEqual` on a decoded image: the claim is that the two bitmaps are the
  // same file, so compare the bytes and let the attachments explain a failure.
  expect(ar.bytes.equals(en.bytes)).toBe(true)
})
