import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'
import type { TabsApi } from '../apps/shell/src/shared/tabs-api'
import type { HtmlApi } from '../apps/html/src/shared/ipc'

test('tab context menu saves current edits from a background tab and tooltip shows path', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-tab-context-'))
  const source = join(dir, 'notes.md')
  const target = join(dir, 'notes-copy.md')
  await writeFile(source, '# Context menu\n\nOriginal paragraph.\n')
  const launched = await launchShell({
    onboardingSeen: true,
    openFile: source,
    videoDir: 'tab-context-actions',
  })
  const { app } = launched
  try {
    const home = await waitForPageWithUrl(app, 'shell/out')
    const editor = await waitForPageWithUrl(app, '://markdown/')
    await expect(editor.locator('.doc-editor h1')).toHaveText('Context menu')
    const content = editor.locator('[contenteditable="true"]').first()
    await content.click()
    await content.press('ControlOrMeta+End')
    await content.press('Enter')
    await content.pressSequentially('Unsaved context-menu edit.')

    // Hover tooltip includes file path
    const tabItem = home.locator('.tab-bar .tab-item:not(.tab-home)').first()
    await expect(tabItem).toHaveAttribute('title', `notes.md\n${source}`)

    const tabId = await home.evaluate(async () => {
      const api = (window as unknown as { aiOfficeTabs: TabsApi }).aiOfficeTabs
      const tab = (await api.list()).find((item) => item.kind === 'markdown')!
      await api.activate('home')
      return tab.id
    })

    await app.evaluate(({ Menu, dialog }, path) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: path })
      Menu.prototype.popup = function () {
        ;(globalThis as typeof globalThis & { __tabContextMenu?: Electron.Menu }).__tabContextMenu =
          this
      }
    }, target)

    await home.evaluate(
      (id) => (window as unknown as { aiOfficeTabs: TabsApi }).aiOfficeTabs.showTabMenu(id, 20, 20),
      tabId,
    )

    const entries = await app.evaluate(() => {
      const menu = (globalThis as typeof globalThis & { __tabContextMenu?: Electron.Menu })
        .__tabContextMenu!
      return menu.items
        .filter((item) => item.type !== 'separator' && item.enabled)
        .map((item) => item.label)
    })
    expect(entries).toEqual(
      expect.arrayContaining([
        'Save',
        'Save As…',
        'Print…',
        'Rename',
        'Copy path',
        'Copy file name',
      ]),
    )

    await app.evaluate(({ BrowserWindow }) => {
      const menu = (globalThis as typeof globalThis & { __tabContextMenu?: Electron.Menu })
        .__tabContextMenu!
      const saveAs = menu.getMenuItemById('saveAs')!
      saveAs.click(undefined as never, BrowserWindow.getFocusedWindow()!, undefined as never)
    })

    await expect
      .poll(() => readFile(target, 'utf8').catch(() => ''))
      .toContain('Unsaved context-menu edit.')

    const active = await home.evaluate(async () =>
      (await (window as unknown as { aiOfficeTabs: TabsApi }).aiOfficeTabs.list()).find(
        (item) => item.active,
      ),
    )
    expect(active?.id).toBe(tabId)
    expect(active?.filePath).toBe(target)
  } finally {
    await closeAndSaveVideo(launched, 'tab-context-actions')
    await rm(dir, { recursive: true, force: true })
  }
})

test('Present tab has no enabled Save, Save As or Print commands', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-tab-present-'))
  const source = join(dir, 'preview.html')
  await writeFile(source, '<!doctype html><html><body><h1>Present preview</h1></body></html>')
  const launched = await launchShell({
    onboardingSeen: true,
    openFile: source,
    videoDir: 'tab-present-menu',
  })
  try {
    const home = await waitForPageWithUrl(launched.app, 'shell/out')
    const editor = await waitForPageWithUrl(launched.app, '://html/')
    await editor.evaluate(() =>
      (window as unknown as { htmlApi: HtmlApi }).htmlApi.presentInNewTab('Preview'),
    )
    await expect
      .poll(async () =>
        home.evaluate(
          async () =>
            (await (window as unknown as { aiOfficeTabs: TabsApi }).aiOfficeTabs.list()).filter(
              (tab) => tab.kind === 'html',
            ).length,
        ),
      )
      .toBe(2)
    const previewId = await home.evaluate(
      async () =>
        (await (window as unknown as { aiOfficeTabs: TabsApi }).aiOfficeTabs.list()).find(
          (tab) => tab.active && tab.kind === 'html',
        )!.id,
    )
    await launched.app.evaluate(({ Menu }) => {
      Menu.prototype.popup = function () {
        ;(globalThis as typeof globalThis & { __tabContextMenu?: Electron.Menu }).__tabContextMenu =
          this
      }
    })
    await home.evaluate(
      (id) => (window as unknown as { aiOfficeTabs: TabsApi }).aiOfficeTabs.showTabMenu(id, 20, 20),
      previewId,
    )
    const commands = await launched.app.evaluate(() => {
      const menu = (globalThis as typeof globalThis & { __tabContextMenu?: Electron.Menu })
        .__tabContextMenu!
      return ['save', 'saveAs', 'print', 'close'].map((id) => ({
        id,
        enabled: menu.getMenuItemById(id)!.enabled,
      }))
    })
    expect(commands).toEqual([
      { id: 'save', enabled: false },
      { id: 'saveAs', enabled: false },
      { id: 'print', enabled: false },
      { id: 'close', enabled: true },
    ])
  } finally {
    await closeAndSaveVideo(launched, 'tab-present-menu')
    await rm(dir, { recursive: true, force: true })
  }
})
