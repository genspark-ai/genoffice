import { expect, test, type Page } from '@playwright/test'
import type { Editor } from '@tiptap/core'
const rebaseSource = '<details>\n<img src="assets/old.png">\n</details>\n'
const source = 'Title\n=====\n\n* item  \n\n\n'

async function openSource(page: Page, enabled: boolean, rebase = false): Promise<void> {
  page.on('pageerror', (error) => console.log('[browser-error]', error.message))
  page.on('console', (message) => {
    if (message.type() === 'error') console.log('[browser-console]', message.text())
  })
  await page.addInitScript(
    ({ text, enabled, rebase }) => {
      if (enabled) localStorage.setItem('mdapp.experimentalRoundTrip', '1')
      localStorage.setItem('mdapp.showAi', '0')
      const off = () => {}
      window.markdownApi = {
        getLanguage: async () => 'en',
        getTheme: async () => 'light',
        onLanguageChanged: () => off,
        onThemeChanged: () => off,
        getAutoSaveDefault: async () => ({ on: false, updatedAt: 0 }),
        onAutoSaveDefaultChanged: () => off,
        getAiPanelPrefs: async () => ({
          fontSize: 'default',
          customFontSize: 14,
          spellcheck: true,
        }),
        onAiPanelPrefsChanged: () => off,
        consumePending: async () => '/fixtures/polish.md',
        readFile: async () => text,
        consumeHeadlessExport: async () => null,
        headlessExportDone: () => {},
        setDirty: () => {},
        save: async ({ text }) => {
          document.body.dataset.saved = text
          return {
            ok: true,
            path: '/fixtures/polish.md',
            imageRewrites: rebase ? [{ from: 'assets/old.png', to: 'assets/new.png' }] : [],
            writtenText: rebase ? text.replace('assets/old.png', 'assets/new.png') : undefined,
          }
        },
        onSaveRequest: (handler) => {
          const save = () => handler('save')
          window.addEventListener('test:save', save)
          return () => window.removeEventListener('test:save', save)
        },
        onReadTextRequest: (handler) => {
          window.addEventListener('test:read-source', handler)
          return () => window.removeEventListener('test:read-source', handler)
        },
        sendReadTextResult: (result) =>
          window.dispatchEvent(new CustomEvent('test:read-source-result', { detail: result })),
        sendSaveRequestAck: () => {},
        onCloseSaveRequest: () => off,
        sendCloseSaveResult: () => {},
        onFileRenamed: () => off,
        pickImage: async () => null,
        saveImage: async () => null,
        readImage: async () => null,
        onExportRequest: () => off,
        onPrintRequest: () => off,
        exportDocx: async () => ({ ok: false, error: 'not used in renderer coverage' }),
        exportPdf: async () => ({ ok: false, error: 'not used in renderer coverage' }),
        onChromePressed: () => off,
        onViewImage: () => off,
        getAiSettings: async () => ({ providers: [] }),
        aiGskStatus: async () => ({ loggedIn: false }),
        aiStream: async () => {},
        aiStreamCancel: async () => {},
        onAiStream: () => off,
        webSearch: async () => ({
          results: [],
          method: 'error',
          error: 'not used in renderer coverage',
        }),
        imageSearch: async () => ({
          images: [],
          method: 'error',
          error: 'not used in renderer coverage',
        }),
        fetchImage: async () => null,
        aiGenerateImage: async () => ({ error: 'not used in renderer coverage' }),
      }
    },
    { text: rebase ? rebaseSource : source, enabled, rebase },
  )
  await page.goto(`http://localhost:${Number(process.env.MARKDOWN_DEV_PORT) || 5177}`)
  await expect(page.locator('.doc-editor')).toBeVisible()
}

for (const enabled of [false, true]) {
  test(`save and MCP read use the ${enabled ? 'opt-in' : 'default'} serializer`, async ({
    page,
  }) => {
    await openSource(page, enabled)
    await expect(page.locator('.doc-editor')).toContainText('Title')
    const read = () =>
      page.evaluate(
        () =>
          new Promise<string>((resolve) => {
            window.addEventListener(
              'test:read-source-result',
              (event) => {
                resolve((event as CustomEvent).detail.text)
              },
              { once: true },
            )
            window.dispatchEvent(new Event('test:read-source'))
          }),
      )
    // block-level splicing already returns an unedited document's source; the
    // opt-in shortcut must agree with it
    const initial = await read()
    expect(initial).toBe(source)
    await page.evaluate(() => window.dispatchEvent(new Event('test:save')))
    await expect(page.locator('body')).toHaveAttribute('data-saved', initial)
    await page.locator('.doc-editor').evaluate((node) => {
      const editor = (node as HTMLElement & { editor: Editor }).editor
      editor.commands.insertContentAt(2, ' edited')
    })
    const edited = await read()
    expect(edited).toContain('edited')
    await page.evaluate(() => window.dispatchEvent(new Event('test:save')))
    await expect(page.locator('body')).toHaveAttribute('data-saved', edited)
    await page.locator('.doc-editor').evaluate((node) => {
      const editor = (node as HTMLElement & { editor: Editor }).editor
      editor.commands.undo()
    })
    // the shortcut restores the loaded bytes; the splice path has already
    // written the edited heading, so undo re-serializes that block and keeps
    // the untouched list verbatim
    const reverted = await read()
    if (enabled) expect(reverted).toBe(initial)
    else {
      expect(reverted).not.toContain('edited')
      expect(reverted).toContain('\n* item  \n')
    }
  })
}

for (const enabled of [false, true]) {
  test(`Save As preserves raw HTML and rebases the ${enabled ? 'snapshot' : 'source map'} for subsequent saves`, async ({
    page,
  }) => {
    await openSource(page, enabled, true)
    await expect(page.locator('.doc-editor img')).toHaveCount(1)
    await page.evaluate(() => window.dispatchEvent(new Event('test:save')))
    await expect(page.locator('body')).toHaveAttribute('data-saved', rebaseSource)
    await page.evaluate(() => window.dispatchEvent(new Event('test:save')))
    await expect(page.locator('body')).toHaveAttribute(
      'data-saved',
      rebaseSource.replace('assets/old.png', 'assets/new.png'),
    )
  })
}
