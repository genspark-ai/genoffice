import { expect, type Page } from '@playwright/test'
export const rebaseSource = '<details>\n<img src="assets/old.png">\n</details>\n'
export const source = 'Title\n=====\n\n* item  \n\n\n'

export async function openSource(
  page: Page,
  enabled: boolean,
  rebase = false,
  text = source,
): Promise<void> {
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
        onExportRequest: (handler) => {
          const exportFile = (event: Event) => handler((event as CustomEvent).detail)
          window.addEventListener('test:export', exportFile)
          return () => window.removeEventListener('test:export', exportFile)
        },
        prepareImageExport: async () => ({ ok: true, canceled: true }),
        writeExportImage: async () => ({ ok: false, error: 'not used in renderer coverage' }),
        finishImageExport: async () => ({ ok: false, error: 'not used in renderer coverage' }),
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
    { text: rebase ? rebaseSource : text, enabled, rebase },
  )
  await page.goto(`http://localhost:${Number(process.env.MARKDOWN_DEV_PORT) || 5177}`)
  await expect(page.locator('.doc-editor')).toBeVisible()
}
