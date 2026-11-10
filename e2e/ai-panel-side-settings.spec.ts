import { test, expect } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

test('AI panel side persists across restart and updates an open Docs tab without losing the draft', async () => {
  const fixture = join(__dirname, 'assets/justify-pagegap-fr.docx')
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'ai-panel-side' })
  const { page } = launched
  try {
    await page.getByRole('button', { name: 'Settings', exact: true }).click()
    await page.locator('.set-nav-item').filter({ hasText: 'General' }).click()
    await page.getByRole('button', { name: 'AI sidebar position', exact: true }).click()
    await page.getByRole('option', { name: 'Right', exact: true }).click()
    await expect(
      page.getByRole('button', { name: 'AI sidebar position', exact: true }),
    ).toContainText('Right')
    await page.screenshot({ path: screenshotPath('ai-panel-side-settings') })
    await page.locator('.set-close').click()
    await page.evaluate((path) => window.aiOffice.openPath(path), fixture)
    const editor = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(editor.locator('.ProseMirror').first()).toBeVisible()
    await expect(editor.locator('html')).toHaveAttribute('data-ai-panel-side', 'right')
    const dock = editor.locator('.ai-dock')
    const content = editor.locator('.app-content')
    expect((await dock.boundingBox())!.x).toBeGreaterThan((await content.boundingBox())!.x)
    const draft = editor.locator('.ai-composer textarea')
    await draft.fill('Keep this unsent draft when switching sides')

    await editor.getByRole('button', { name: 'Move AI panel to the left', exact: true }).click()
    await expect(editor.locator('html')).toHaveAttribute('data-ai-panel-side', 'left')
    expect((await dock.boundingBox())!.x).toBeLessThan((await content.boundingBox())!.x)
    await expect(draft).toHaveValue('Keep this unsent draft when switching sides')
    await editor.getByRole('button', { name: 'Move AI panel to the right', exact: true }).click()
    await expect(editor.locator('html')).toHaveAttribute('data-ai-panel-side', 'right')
    const before = (await dock.boundingBox())!.width
    const resizer = editor.locator('.ai-panel-resizer')
    await resizer.hover({ position: { x: 3, y: 100 } })
    const handle = (await resizer.boundingBox())!
    await editor.mouse.move(handle.x + 3, handle.y + 100)
    await editor.mouse.down()
    await expect(editor.locator('.ai-panel')).toHaveClass(/ai-panel-resizing/)
    await editor.mouse.move(handle.x - 47, handle.y + 100, { steps: 5 })
    await editor.mouse.up()
    await expect.poll(async () => (await dock.boundingBox())!.width).toBeGreaterThan(before + 30)
    await editor.getByRole('button', { name: 'Collapse panel', exact: true }).click()
    await expect.poll(async () => Math.round((await dock.boundingBox())!.width)).toBe(34)
    await editor.locator('.ai-rail').click()
    await expect(draft).toBeVisible()
    await expect(draft).toHaveValue('Keep this unsent draft when switching sides')
    await editor.getByRole('button', { name: 'View', exact: true }).click()
    await editor.getByRole('button', { name: 'Navigation Pane', exact: true }).click()
    await expect(editor.locator('.nav-pane')).toBeVisible()
    expect((await editor.locator('.nav-pane').boundingBox())!.x).toBeLessThan(
      (await editor.locator('.editor-area').boundingBox())!.x,
    )
    expect((await dock.boundingBox())!.x).toBeGreaterThan(
      (await editor.locator('.editor-area').boundingBox())!.x,
    )
    await editor.screenshot({ path: screenshotPath('ai-panel-side-docs-right') })
    const saved = JSON.parse(
      await readFile(join(launched.userDataDir, 'app-settings.json'), 'utf8'),
    )
    expect(saved.aiPanelSide).toBe('right')
  } finally {
    await closeAndSaveVideo(launched, 'ai-panel-side')
  }

  const restarted = await launchShell({
    userDataDir: launched.userDataDir,
    videoDir: 'ai-panel-side-restart',
    openFile: fixture,
  })
  try {
    const editor = await waitForPageWithUrl(restarted.app, '://docs/')
    await expect(editor.locator('.ProseMirror').first()).toBeVisible()
    await expect(editor.locator('html')).toHaveAttribute('data-ai-panel-side', 'right')
    expect((await editor.locator('.ai-dock').boundingBox())!.x).toBeGreaterThan(
      (await editor.locator('.app-content').boundingBox())!.x,
    )
  } finally {
    await closeAndSaveVideo(restarted, 'ai-panel-side-restart')
  }
})
