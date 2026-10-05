import { test, expect, _electron as electron } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import { mkdtemp, readFile, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { createRequire } from 'node:module'
import JSZip from 'jszip'
import type { SlidesApi } from '../apps/slides/src/shared/ipc'

declare global {
  interface Window {
    slidesApi: SlidesApi
  }
}

/**
 * Withholding a picture from the model, driven through the real app.
 *
 * The picture is centred on the slide so a click at the stage's centre lands on
 * it, which keeps the test off the coordinate arithmetic that a Konva canvas
 * would otherwise demand. Everything after the click is the reader's own path:
 * right-click, name the span, commit.
 *
 * What is asserted afterwards is the point of the whole feature: the words and
 * the picture's own bytes survive in the file, while the model is told a
 * placeholder is there and is handed no reference to fetch.
 */

const LABEL = 'client logo'
const SECRET = '13800138000'

// a 1x1 transparent PNG, so the picture is a real picture part
const PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  'base64',
)

const W = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'

/** centred on the slide, so the stage centre hits it */
const PIC_OFF = { x: 2514600, y: 1143000 }
const PIC_EXT = { x: 4114800, y: 2857500 }

async function buildPptx(path: string): Promise<void> {
  const slide =
    `<?xml version="1.0"?><p:sld ${P} ${R} ${W}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr/><p:grpSpPr/>' +
    // the picture, centred
    '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Logo"/><p:cNvPicPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr><a:xfrm><a:off x="${PIC_OFF.x}" y="${PIC_OFF.y}"/><a:ext cx="${PIC_EXT.x}" cy="${PIC_EXT.y}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"/></p:spPr></p:pic>' +
    // a text box below it, carrying the secret this test is about
    '<p:sp><p:nvSpPr><p:cNvPr id="5" name="Contact"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="1000000" y="4200000"/><a:ext cx="7144000" cy="700000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"/></p:spPr>' +
    '<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r>' +
    `<a:t>Call ${SECRET} now to confirm the order.</a:t></a:r></a:p></p:txBody></p:sp>` +
    '</p:spTree></p:cSld></p:sld>'

  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/>' +
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
      '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'ppt/presentation.xml',
    '<?xml version="1.0"?><p:presentation ' +
      `${P} ${R} ${W}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>` +
      '<p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>' +
      '</Relationships>',
  )
  zip.file('ppt/slides/slide1.xml', slide)
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>' +
      '</Relationships>',
  )
  zip.file('ppt/media/image1.png', PNG)
  const buf = await zip.generateAsync({ type: 'nodebuffer' })
  await writeFile(path, buf)
}

test('a reader can withhold a picture from the model without losing it', async () => {
  test.setTimeout(180_000)
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-redact-slides-'))
  const pptx = join(dir, 'deck.pptx')
  await buildPptx(pptx)

  const require = createRequire(resolve('apps/slides/package.json'))
  const { ELECTRON_RUN_AS_NODE: _node, ...env } = process.env
  const app = await electron.launch({
    executablePath: require('electron'),
    args: [
      ...(process.platform === 'linux' ? ['--no-sandbox', '--disable-gpu'] : []),
      resolve('apps/slides'),
      pptx,
    ],
    env: { ...env, GENOFFICE_USER_DATA: join(dir, 'user-data'), GENOFFICE_LANG: 'en' },
  })
  try {
    const page = await app.firstWindow()
    await expect(page.locator('.stage-wrap canvas').first()).toBeVisible({ timeout: 30_000 })
    const shot = (name: string) =>
      page.screenshot({ path: `e2e/artifacts/screenshots/slides-${name}.png` })

    // the picture is centred, so a click at the stage centre selects it
    const stage = page.locator('.stage-wrap').first()
    await stage.click()
    await shot('1-deck-opened')

    await stage.click({ button: 'right' })
    const menu = page.locator('.ctx-menu')
    await expect(menu).toBeVisible({ timeout: 15_000 })
    const item = menu.locator('.ctx-item', { hasText: 'Hide the selection from AI' })
    await expect(item).toHaveCount(1)
    await shot('2-context-menu')
    await item.click()

    const dialog = page.locator('.modal', { hasText: 'Hide the selection from AI' })
    await expect(dialog).toBeVisible({ timeout: 15_000 })
    await shot('3-dialog')

    const input = dialog.locator('input').first()
    await input.fill(LABEL)
    // the marker the model will be shown, live
    await expect(dialog.locator('.redact-preview')).toHaveText(`{{${LABEL}}}`)
    await shot('4-dialog-preview')

    const confirm = dialog.locator('.modal-actions button').last()
    await expect(confirm).toBeEnabled()
    await confirm.click()
    await expect(dialog).toHaveCount(0)
    await shot('5-marked')

    // save through the app's own API rather than the keyboard shortcut: a
    // native save dialog would block the renderer and the shortcut is flaky
    // once focus has moved through a right-click and a modal
    const saved = await page.evaluate(() => window.slidesApi.save())
    expect(saved.ok, saved.error).toBe(true)

    // what the file now says
    const z = await JSZip.loadAsync(await readFile(pptx))
    const xml = await z.file('ppt/slides/slide1.xml')!.async('string')
    expect(xml).toContain('go:redact')
    expect(xml).toContain(`w:label="${LABEL}"`)
    // the picture's own bytes are untouched — nothing was removed
    expect(xml).toContain('r:embed="rId2"')
    expect(await z.file('ppt/media/image1.png')!.async('nodebuffer')).toEqual(PNG)
    // the secret in the text box is still there, and no placeholder reached the file
    expect(xml).toContain(SECRET)
    expect(xml).not.toContain('{{')
    await shot('6-saved')
  } finally {
    await app.close().catch(() => {})
    await rm(dir, { recursive: true, force: true }).catch(() => {})
  }
})
