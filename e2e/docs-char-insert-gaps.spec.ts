/**
 * Character and insert features that were left simplified: Text Effects has
 * presets and a reflection, the text of an existing text box can be turned,
 * and Font Substitution lets the stand-in be picked and converts for good.
 * Each is driven through the real menu, dialog or ribbon and checked on the
 * page and in the saved file.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-char-insert-gaps'

/** clicks a menu bar item once the renderer has told the menu it is available */
async function clickMenu(app: ElectronApplication, path: string[]) {
  const click = () =>
    app.evaluate(({ Menu }, path) => {
      let items = Menu.getApplicationMenu()?.items
      let item: Electron.MenuItem | undefined
      for (const label of path) {
        item = items?.find((candidate) => candidate.label === label)
        items = item?.submenu?.items
      }
      if (!item?.enabled) return false
      item.click()
      return true
    }, path)
  await expect.poll(click, { timeout: 15_000 }).toBe(true)
}

const partXml = async (docPath: string, part = 'word/document.xml') =>
  (await (await JSZip.loadAsync(readFileSync(docPath))).file(part)?.async('string')) ?? ''

/** selects document positions: the first paragraph's text starts at 1 */
const select = (page: Page, from: number, to: number) =>
  page.evaluate(
    ([from, to]) => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } } }
          }
        }
      ).__aidocs
      editor.chain().focus().setTextSelection({ from, to }).run()
    },
    [from, to],
  )

test('Text Effects: a preset and a reflection are set from the dialog, drawn, saved and read back', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-gaps-fx-')))
  const docPath = join(dir, 'effects.docx')
  writeFileSync(docPath, await docxWith(para('Plain fancy words') + para('Last')))
  /** the mirrored piece of the first paragraph, as the page draws it */
  const mirrored = (page: Page) =>
    page.evaluate(() => {
      const span = Array.from(document.querySelectorAll<HTMLElement>('.doc-page p span')).find(
        (el) => getComputedStyle(el).getPropertyValue('-webkit-box-reflect') !== 'none',
      )
      if (!span) return null
      const style = getComputedStyle(span)
      return {
        text: span.textContent,
        reflect: style.getPropertyValue('-webkit-box-reflect'),
        color: style.color,
        stroke: style.getPropertyValue('-webkit-text-stroke-color'),
        glow: style.textShadow !== 'none',
      }
    })
  const expected = {
    text: 'fancy',
    // below the text, touching it; the mask is clear down to 45% and half opaque at the foot
    reflect: expect.stringMatching(
      /^below 0px linear-gradient\(rgba\(0, 0, 0, 0\) 45%, rgba\(0, 0, 0, 0\.5\)\)/,
    ),
    color: 'rgb(255, 255, 255)',
    stroke: 'rgb(68, 114, 196)',
    glow: true,
  }
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Plain fancy words' })).toBeVisible({
      timeout: 30_000,
    })
    await select(page, 7, 12)
    await clickMenu(app, ['Format', 'Text Effects…'])
    const dialog = page.locator('.text-fx-dialog')
    await expect(dialog).toBeVisible()
    await expect(dialog.locator('.text-fx-preset')).toHaveCount(6)
    await expect(dialog.locator('[data-fx="reflection"] .text-fx-choice.active')).toHaveText('None')

    // a preset fills in the groups below it
    await dialog.locator('.text-fx-preset[data-preset="white-outline-glow"]').click()
    await expect(dialog.locator('[data-fx="outline"] .text-fx-choice.active')).toHaveText('0.75 pt')
    await expect(dialog.locator('[data-fx="glow"] .text-fx-choice.active')).toHaveText('5 pt')
    await expect(dialog.locator('[data-fx="shadow"] .text-fx-choice.active')).toHaveText('None')
    // its colour stays through a change that keeps the preset's effects
    await dialog.locator('.text-fx-preset[data-preset="blue-reflection"]').click()
    await expect(dialog.locator('[data-fx="reflection"] .text-fx-choice.active')).toHaveText(
      'Half Reflection',
    )
    await dialog.locator('.text-fx-preset[data-preset="white-outline-glow"]').click()
    await expect(dialog.locator('.text-fx-preset.active')).toHaveCount(1)
    await page.screenshot({ path: test.info().outputPath('text-effects-preset.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const span = document.querySelector<HTMLElement>('.doc-page p span[style*="stroke"]')
          return span ? [span.textContent, getComputedStyle(span).color] : null
        }),
      )
      .toEqual(['fancy', 'rgb(255, 255, 255)'])

    // the reflection is added to what the text has; the colour of the preset stays on it
    await select(page, 7, 12)
    await clickMenu(app, ['Format', 'Text Effects…'])
    await expect(dialog.locator('[data-fx="glow"] .text-fx-choice.active')).toHaveText('5 pt')
    await dialog
      .locator('[data-fx="reflection"] .text-fx-choice', { hasText: /^Half Reflection$/ })
      .click()
    await page.screenshot({ path: test.info().outputPath('text-effects-reflection-dialog.png') })
    await dialog.locator('button.primary').click()
    await expect.poll(() => mirrored(page)).toEqual(expected)
    await page.screenshot({ path: test.info().outputPath('text-effects-reflection.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await partXml(docPath)
          return [
            xml.includes('<w14:reflection w14:blurRad="6350" w14:stA="50000"'),
            xml.includes('w14:endPos="55000"'),
            xml.includes('<w:color w:val="FFFFFF"/>'),
            xml.includes('<w14:glow w14:rad="63500">'),
            xml.includes('<w14:textOutline w14:w="9525"'),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true, true])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }

  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Plain fancy words' })).toBeVisible({
      timeout: 30_000,
    })
    await expect.poll(() => mirrored(page)).toEqual(expected)
    // the dialog reads the reflection back, and None takes it out of the file
    await select(page, 7, 12)
    await clickMenu(app, ['Format', 'Text Effects…'])
    const dialog = page.locator('.text-fx-dialog')
    await expect(dialog.locator('[data-fx="reflection"] .text-fx-choice.active')).toHaveText(
      'Half Reflection',
    )
    await dialog.locator('[data-fx="reflection"] .text-fx-choice', { hasText: /^None$/ }).click()
    await dialog.locator('button.primary').click()
    await expect.poll(() => mirrored(page)).toBeNull()
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await partXml(docPath)
          return [xml.includes('w14:reflection'), xml.includes('<w14:glow w14:rad="63500">')]
        },
        { timeout: 30_000 },
      )
      .toEqual([false, true])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Shape Format > Text Direction turns the text of an existing text box, and the file says so', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-gaps-box-')))
  const docPath = join(dir, 'box.docx')
  writeFileSync(docPath, await docxWith(para('First') + para('Last')))
  /** how the box draws its text: the classes that say its direction, and what CSS makes of them */
  const direction = (page: Page) =>
    page.evaluate(() => {
      const box = document.querySelector<HTMLElement>('.doc-page .doc-textbox')
      if (!box) return null
      const p = box.querySelector('p, .doc-textbox-para') ?? box
      const style = getComputedStyle(p)
      const r = p.getBoundingClientRect()
      return {
        classes: Array.from(box.classList)
          .filter((name) => name.startsWith('doc-textbox-') && name !== 'doc-textbox-filled')
          .sort()
          .join(' '),
        mode: style.writingMode,
        orientation: style.textOrientation,
        tall: r.height > r.width,
        text: box.textContent,
      }
    })
  const pick = async (page: Page, label: string) => {
    await page.locator('[data-ribbon-body] .gs-dd-btn[data-tip="Text Direction"]').click()
    await page.locator('.gs-dd-pop [role="option"]', { hasText: new RegExp(`^${label}$`) }).click()
  }
  const shown = (page: Page) =>
    page.locator('[data-ribbon-body] .gs-dd-btn[data-tip="Text Direction"] .gs-dd-value')
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const docPage = page.locator('.doc-page[contenteditable="true"]')
    await expect(page.locator('.doc-page p', { hasText: 'First' })).toBeVisible({ timeout: 30_000 })
    await clickMenu(app, ['Insert', 'Text Box', 'Text Box'])
    await expect(docPage).toHaveClass(/doc-shape-drawing/)
    const paper = (await docPage.boundingBox())!
    await page.mouse.click(paper.x + 300, paper.y + 260)
    await expect(docPage.locator('.doc-textbox')).toHaveCount(1)
    await expect
      .poll(() => page.evaluate(() => !!document.activeElement?.closest('.doc-textbox-editor')))
      .toBe(true)
    await page.keyboard.type('Turn me', { delay: 10 })
    await expect(docPage.locator('.doc-textbox')).toContainText('Turn me')
    await expect(page.locator('.ribbon-tab.active')).toHaveText('Shape Format')
    await expect(shown(page)).toHaveText('Horizontal')
    expect(await direction(page)).toMatchObject({ classes: '', mode: 'horizontal-tb', tall: false })

    // stacked: upright letters one under the other, lines left to right; the typed text stays
    await pick(page, 'Stacked')
    await expect
      .poll(() => direction(page))
      .toEqual({
        classes: 'doc-textbox-vert doc-textbox-wordartvert',
        mode: 'vertical-lr',
        orientation: 'upright',
        tall: true,
        text: 'Turn me',
      })
    await expect(shown(page)).toHaveText('Stacked')
    await page.screenshot({ path: test.info().outputPath('box-stacked.png') })
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await partXml(docPath)
          return xml.includes('<wps:bodyPr vert="wordArtVert"') && xml.includes('Turn me')
        },
        { timeout: 30_000 },
      )
      .toBe(true)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }

  // the saved box is now a box from a file: its direction is read and can be changed again
  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const box = page.locator('.doc-page .doc-textbox')
    await expect(box).toHaveCount(1, { timeout: 30_000 })
    await expect
      .poll(() => direction(page))
      .toMatchObject({ classes: 'doc-textbox-vert doc-textbox-wordartvert', mode: 'vertical-lr' })
    await box.click()
    await expect(page.locator('.ribbon-tab.active')).toHaveText('Shape Format')
    await expect(shown(page)).toHaveText('Stacked')

    await pick(page, 'Rotate all text 270°')
    await expect
      .poll(() => direction(page))
      .toMatchObject({
        classes: 'doc-textbox-vert doc-textbox-vert270',
        mode: 'sideways-lr',
        tall: true,
        text: 'Turn me',
      })
    await page.screenshot({ path: test.info().outputPath('box-rotate-270.png') })
    await saveFromMenu(app)
    await expect
      .poll(async () => (await partXml(docPath)).includes('<wps:bodyPr vert="vert270"'), {
        timeout: 30_000,
      })
      .toBe(true)

    await pick(page, 'Rotate all text 90°')
    await expect
      .poll(() => direction(page))
      .toMatchObject({ mode: 'vertical-rl', orientation: 'sideways', tall: true })

    // horizontal again: the attribute is gone from the file
    await pick(page, 'Horizontal')
    await expect
      .poll(() => direction(page))
      .toMatchObject({ classes: '', mode: 'horizontal-tb', tall: false, text: 'Turn me' })
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await partXml(docPath)
          return [/<wps:bodyPr\b[^>]*\bvert=/.test(xml), xml.includes('Turn me')]
        },
        { timeout: 30_000 },
      )
      .toEqual([false, true])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Font Substitution: the stand-in for a missing font is picked, and Convert Permanently rewrites the document', async () => {
  test.setTimeout(240_000)
  const MISSING = 'Zzyzx Gone Sans'
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-gaps-font-')))
  const docPath = join(dir, 'fonts.docx')
  const fonts = `<w:rPr><w:rFonts w:ascii="${MISSING}" w:hAnsi="${MISSING}"/></w:rPr>`
  const source = await JSZip.loadAsync(
    await docxWith(
      `<w:p><w:r>${fonts}<w:t>iiii WWWW in a font nobody has</w:t></w:r></w:p>` +
        '<w:p><w:pPr><w:pStyle w:val="Gone"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>' +
        para('Plain'),
    ),
  )
  source.file(
    'word/styles.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      `<w:style w:type="paragraph" w:styleId="Gone"><w:name w:val="Gone"/><w:basedOn w:val="Normal"/>${fonts}</w:style>` +
      '</w:styles>',
  )
  source.file(
    '[Content_Types].xml',
    (await source.file('[Content_Types].xml')!.async('string')).replace(
      '</Types>',
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    ),
  )
  source.file(
    'word/_rels/document.xml.rels',
    (await source.file('word/_rels/document.xml.rels')!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    ),
  )
  writeFileSync(docPath, await source.generateAsync({ type: 'nodebuffer' }))

  /** how far the first paragraph's text, as the page draws it, is from its width in `family` */
  const offFrom = (page: Page, family: string) =>
    page.evaluate((family) => {
      const span = document.querySelector<HTMLElement>('.doc-page p span')!
      const range = document.createRange()
      range.selectNodeContents(span)
      const drawn = range.getBoundingClientRect().width
      const probe = document.createElement('span')
      probe.textContent = span.textContent
      probe.style.cssText = `font-family:"${family}";font-size:${getComputedStyle(span).fontSize};position:absolute;visibility:hidden;white-space:pre`
      document.body.append(probe)
      const inFamily = probe.getBoundingClientRect().width
      probe.remove()
      const zoom = span.getBoundingClientRect().width / span.offsetWidth || 1
      return Math.abs(Math.round(drawn / zoom) - Math.round(inFamily))
    }, family)

  /** the listed fonts this computer has: the chooser also offers Office fonts it may lack */
  const installed = (chooser: Locator) =>
    chooser.evaluate(async (select: HTMLSelectElement) => {
      const out: string[] = []
      for (const { value } of select.options) {
        if (!value) continue
        try {
          await new FontFace('e2e-probe', `local("${value.replace(/["\\]/g, '')}")`).load()
          out.push(value)
        } catch {
          /* not on this computer */
        }
      }
      return out
    })

  let standIn = ''
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Plain' })).toBeVisible({ timeout: 30_000 })
    await clickMenu(app, ['Format', 'Font Substitution…'])
    const dialog = page.locator('.font-substitution-dialog')
    await expect(dialog).toBeVisible()
    const row = dialog.locator(`tr[data-font="${MISSING}"]`)
    await expect(row).toHaveCount(1)
    const chooser = row.locator('select')
    await expect(chooser).toHaveValue('')
    await page.screenshot({ path: test.info().outputPath('font-substitution-dialog.png') })
    // a font of another width than the app's own stand-in, so the redraw shows
    await expect(async () => {
      for (const family of await installed(chooser))
        if ((await offFrom(page, family)) > 5) return void (standIn = family)
      throw new Error('no listed font on this computer differs from the default stand-in')
    }).toPass({ timeout: 30_000 })
    await chooser.selectOption(standIn)
    // the pick waits for the computer's fonts to be listed, which can take a while
    await expect(chooser).toHaveValue(standIn, { timeout: 90_000 })
    // the page is drawn again in the picked font, wherever the missing one is named
    await expect.poll(() => offFrom(page, standIn)).toBeLessThan(2)
    await page.screenshot({ path: test.info().outputPath('font-substitution-picked.png') })
    // the pick is kept for this computer
    expect(
      await page.evaluate(() => localStorage.getItem('genoffice.docs.fontSubstitutes')),
    ).toContain(`"substitute":"${standIn}"`)
    // nothing is written to the document by a pick alone
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    expect(await partXml(docPath)).toContain(MISSING)

    // reopened, the dialog shows the pick; Convert Permanently writes it into the document
    await clickMenu(app, ['Format', 'Font Substitution…'])
    await expect(row.locator('select')).toHaveValue(standIn)
    await dialog.locator('button[name="convert"]').click()
    await expect(dialog).toBeHidden()
    // converted fonts are no longer missing fonts of this document
    await clickMenu(app, ['Format', 'Font Substitution…'])
    await expect(dialog.locator('tr[data-font]')).toHaveCount(0)
    await expect(dialog.locator('.font-substitution-none')).toBeVisible()
    await dialog.locator('button.primary').click()

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const [document, styles] = [
            await partXml(docPath),
            await partXml(docPath, 'word/styles.xml'),
          ]
          return [
            document.includes(MISSING),
            styles.includes(MISSING),
            document.includes(`w:ascii="${standIn}" w:hAnsi="${standIn}"`),
            styles.includes(`w:ascii="${standIn}" w:hAnsi="${standIn}"`),
            document.includes('iiii WWWW in a font nobody has'),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([false, false, true, true, true])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }

  // a fresh profile has no pick: the converted file needs none, its fonts are real
  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Plain' })).toBeVisible({ timeout: 30_000 })
    await expect.poll(() => offFrom(page, standIn)).toBeLessThan(2)
    await clickMenu(app, ['Format', 'Font Substitution…'])
    const dialog = page.locator('.font-substitution-dialog')
    await expect(dialog.locator('.font-substitution-none')).toBeVisible()
    await expect(dialog.locator('button[name="convert"]')).toBeDisabled()
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
