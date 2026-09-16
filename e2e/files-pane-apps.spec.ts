import { test, expect } from '@playwright/test'
import { copyFileSync, mkdirSync, mkdtempSync, realpathSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { launchShell, closeAndSaveVideo, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * The shared Files pane mounts in every editor: opened from the left-edge tab,
 * it shows the tree and highlights the document this tab is editing.
 */
const cases: Array<{ app: string; file: (root: string) => string }> = [
  {
    app: 'docs',
    file: (root) => {
      const p = join(root, 'Team', 'report.docx')
      copyFileSync(resolve(__dirname, 'assets/justify-pagegap-fr.docx'), p)
      return p
    },
  },
  {
    app: 'slides',
    file: (root) => {
      const p = join(root, 'Team', 'deck.pptx')
      copyFileSync(
        resolve(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
        p,
      )
      return p
    },
  },
  {
    app: 'html',
    file: (root) => {
      const p = join(root, 'Team', 'page.html')
      writeFileSync(p, '<!doctype html><html><body><h1>Hi</h1></body></html>')
      return p
    },
  },
]

for (const c of cases) {
  test(`files pane smoke: ${c.app}`, async () => {
    const root = realpathSync(mkdtempSync(join(tmpdir(), 'fp-smoke-')))
    mkdirSync(join(root, 'Team'))
    writeFileSync(join(root, 'notes.md'), '# n')
    const file = c.file(root)
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root },
      videoDir: 'fp-smoke',
      openFile: file,
    })
    try {
      const editor = await waitForPageWithUrl(launched.app, `://${c.app}/`)
      if (c.app === 'slides') {
        // the thumbnail rail owns the left edge by default, so Slides opens the pane from View
        await editor
          .locator('.ribbon-tab', { hasText: /^View$/ })
          .first()
          .click({ timeout: 40_000 })
        await editor.locator('.ribbon-body .rb-check', { hasText: /^Files$/ }).click()
      } else {
        await editor.locator('.files-edge-tab').click({ timeout: 40_000 })
      }
      await expect(editor.locator('.files-pane')).toBeVisible()
      await expect(editor.locator('.fp-row.current .fp-name')).toHaveText(file.split('/').pop()!)
      await editor.screenshot({ path: screenshotPath(`files-pane-${c.app}`) })
    } finally {
      await closeAndSaveVideo(launched, 'fp-smoke')
    }
  })
}

for (const app of ['sheets', 'pdf'] as const) {
  test(`files pane smoke (new file): ${app}`, async () => {
    const root = realpathSync(mkdtempSync(join(tmpdir(), 'fp-smoke-')))
    mkdirSync(join(root, 'Team'))
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root },
      videoDir: 'fp-smoke',
    })
    try {
      await launched.page
        .locator('.quick-card', { hasText: app === 'sheets' ? 'AI Sheets' : 'AI PDF' })
        .click()
      const editor = await waitForPageWithUrl(launched.app, `://${app}/`)
      if (app === 'pdf') {
        // the thumbnail pane owns the left edge by default, so PDF opens the pane from View
        await editor.locator('.ribbon-tab', { hasText: 'View' }).click({ timeout: 40_000 })
        await editor.locator('.rb-big', { hasText: 'Files' }).click()
      } else {
        await editor.locator('.files-edge-tab').click({ timeout: 40_000 })
      }
      await expect(editor.locator('.files-pane')).toBeVisible()
      await expect(editor.locator('.fp-row.current')).toHaveCount(1)
      await editor.screenshot({ path: screenshotPath(`files-pane-${app}`) })
    } finally {
      await closeAndSaveVideo(launched, 'fp-smoke')
    }
  })
}
