import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Home ▸ Paste split button (Word parity): the icon half pastes with the
 * default mode, the label half opens Keep Source Formatting / Merge
 * Formatting / Keep Text Only, which paste the clipboard in that mode
 * through the same pipeline as Ctrl+V.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown }
}

const WEB_HTML =
  `<meta charset='utf-8'>` +
  `<p style="font-family: Arial, sans-serif; font-size: 14px;">web one <b>strong</b></p>`
const WEB_TEXT = 'web one strong'

test.describe('docs paste split button', () => {
  test('menu modes paste the clipboard in the chosen mode', async () => {
    test.setTimeout(120_000)
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-paste-split' })
    const { app, page } = launched
    try {
      await expect(page.locator('.quick-card').first()).toContainText('AI Docs')
      await page.locator('.quick-card').first().click()
      const editorPage = await waitForPageWithUrl(app, '://docs/')
      await editorPage.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      await editorPage.locator('.doc-page').click()
      await editorPage.keyboard.type('host ', { delay: 10 })
      await app.evaluate(
        ({ clipboard }, payload) => clipboard.write({ html: payload.html, text: payload.text }),
        { html: WEB_HTML, text: WEB_TEXT },
      )

      const runs = async (): Promise<Array<{ text: string; font: string | null; bold: boolean }>> =>
        editorPage.evaluate(() => {
          const ed = (window as unknown as AidocsWindow).__aidocs!.editor! as {
            state: {
              doc: {
                descendants: (
                  cb: (node: {
                    isText: boolean
                    text?: string
                    marks: Array<{ type: { name: string }; attrs: Record<string, unknown> }>
                  }) => boolean,
                ) => void
              }
            }
          }
          const out: Array<{ text: string; font: string | null; bold: boolean }> = []
          ed.state.doc.descendants((node) => {
            if (node.isText) {
              const style = node.marks.find((m) => m.type.name === 'docTextStyle')
              out.push({
                text: node.text ?? '',
                font: (style?.attrs.font as string | null) ?? null,
                bold: node.marks.some((m) => m.type.name === 'bold'),
              })
            }
            return true
          })
          return out
        })

      const menu = editorPage.getByTestId('ribbon-paste-menu')
      await menu.click()
      await editorPage.getByRole('menuitem', { name: 'Keep Text Only' }).click()
      await editorPage.waitForTimeout(400)
      let state = await runs()
      expect(state.some((r) => r.text.includes('web one'))).toBe(true)
      expect(state.some((r) => r.bold)).toBe(false)
      expect(state.find((r) => r.text.includes('web one'))?.font).not.toBe('Arial')

      await menu.click()
      await editorPage.getByRole('menuitem', { name: 'Keep Source Formatting' }).click()
      await editorPage.waitForTimeout(400)
      state = await runs()
      expect(state.find((r) => r.text === 'strong')?.bold).toBe(true)
      expect(state.filter((r) => r.text.includes('web one')).pop()?.font).toBe('Arial')
      await expect(editorPage.locator('[data-testid="paste-options-chip"]')).toBeVisible()
    } finally {
      await closeAndSaveVideo(launched)
    }
  })
})
