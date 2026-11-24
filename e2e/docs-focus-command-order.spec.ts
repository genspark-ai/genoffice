/**
 * A command that hands the focus back to the page does so at once: what the
 * user does next (a click elsewhere, another dialog, typing) is not undone a
 * frame later. The page's animation frames are held from a dialog's OK and
 * released at a chosen point, so the order that loses is the one that runs.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-focus-command-order'

type W = {
  __aidocs?: {
    editor?: {
      view: { hasFocus(): boolean }
      state: { selection: { $from: { parent: { textContent: string } } } }
    }
  }
  __release?: () => void
}

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

/** holds requestAnimationFrame; released by `trigger` ('mouseup' in the page, 'key:<k>') or by window.__release() */
const hold = (page: Page, trigger: string) =>
  page.evaluate((trigger) => {
    const raf = window.requestAnimationFrame.bind(window)
    const held: FrameRequestCallback[] = []
    let open = false
    window.requestAnimationFrame = ((cb: FrameRequestCallback) =>
      open ? raf(cb) : held.push(cb)) as typeof requestAnimationFrame
    const release = () => {
      if (open) return
      open = true
      window.requestAnimationFrame = raf
      for (const cb of held.splice(0)) cb(performance.now())
    }
    ;(window as unknown as { __release: () => void }).__release = release
    if (trigger === 'mouseup')
      window.addEventListener(
        'mouseup',
        (e) => {
          if ((e.target as HTMLElement).closest?.('.ProseMirror')) release()
        },
        true,
      )
    else if (trigger.startsWith('key:'))
      window.addEventListener(
        'keydown',
        (e) => {
          if (e.key === trigger.slice(4)) release()
        },
        true,
      )
  }, trigger)

const twoFrames = (page: Page) =>
  page.evaluate(
    () => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))),
  )

let dir: string
let app: ElectronApplication
let page: Page
let launched: Awaited<ReturnType<typeof launchShell>>

const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
const caretParagraph = () =>
  page.evaluate(
    () => (window as unknown as W).__aidocs?.editor?.state.selection.$from.parent.textContent ?? '',
  )

test.beforeEach(async () => {
  dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-focus-order-')))
  const docPath = join(dir, 'focus.docx')
  writeFileSync(docPath, await docxWith(para('apple') + para('banana') + para('cherry')))
  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  app = launched.app
  page = await waitForPageWithUrl(app, '://docs/')
  await expect(paragraph('cherry')).toBeVisible({ timeout: 30_000 })
  await page.waitForFunction(
    () => (window as unknown as W).__aidocs?.editor?.view.hasFocus() === true,
    undefined,
    { timeout: 30_000 },
  )
  await paragraph('apple').click()
  await expect.poll(caretParagraph).toContain('apple')
  await page.keyboard.press('ControlOrMeta+d')
  await expect(page.locator('.modal.font-dialog')).toBeVisible()
  await page.locator('.modal.font-dialog').getByLabel('Small caps').check()
})

test.afterEach(async () => {
  await closeAndSaveVideo(launched, VIDEO_DIR)
  rmSync(dir, { recursive: true, force: true })
})

const fontOk = async () => {
  const dialog = page.locator('.modal.font-dialog')
  await dialog.getByRole('button', { name: 'OK' }).click()
  await expect(dialog).toBeHidden()
}

test('OK, then a click elsewhere in the page: the click stays', async () => {
  await hold(page, 'mouseup')
  await fontOk()
  await paragraph('cherry').click()
  await page.keyboard.type('Z')
  await expect(paragraph('cherry')).toContainText('Z')
  await expect(paragraph('apple')).not.toContainText('Z')
  expect(await caretParagraph()).toContain('cherry')
})

test('OK, then another dialog: the focus stays in the second dialog', async () => {
  await hold(page, 'none')
  await fontOk()
  await clickMenu(app, ['Format', 'Paragraph…'])
  const second = page.locator('.modal', { has: page.locator('h2', { hasText: /^Paragraph$/ }) })
  await expect(second).toBeVisible()
  await expect(second.locator(':focus')).toHaveCount(1)
  await page.evaluate(() => (window as unknown as W).__release!())
  await twoFrames(page)
  await expect(second.locator(':focus')).toHaveCount(1)
})

test('OK, then typing at once: every key arrives', async () => {
  await hold(page, 'key:z')
  await fontOk()
  await page.keyboard.type('xyz')
  await expect(paragraph('apple')).toContainText('xyz')
})
