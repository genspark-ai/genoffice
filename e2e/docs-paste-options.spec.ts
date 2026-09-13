import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Post-paste "Paste options" chip (Word parity): after pasting foreign web
 * HTML the chip appears at the paste point and re-applies the same payload
 * in the chosen mode —
 *  - Merge Formatting: emphasis stays, text formats like typing (caret font);
 *  - Keep Text Only: plain text, caret font;
 *  - Keep Source Formatting: the site's concrete font again;
 * "Always use this choice" persists the mode for the next paste.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown }
}

const WEB_TWO_PARAS =
  `<meta charset='utf-8'>` +
  `<p style="margin: 0 0 16px; font-family: Arial, sans-serif; font-size: 14px;">web one <b>strong</b></p>` +
  `<p style="margin: 0 0 16px; font-family: Arial, sans-serif; font-size: 14px;">web two</p>`
const WEB_TEXT = 'web one strong\n\nweb two'

test.describe('docs paste options chip', () => {
  test('chip re-applies merge/text/source modes and remembers the default', async () => {
    test.setTimeout(120_000)
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-paste-options' })
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
      await editorPage.keyboard.type('avant', { delay: 10 })
      await editorPage.evaluate(() => {
        const ed = (window as unknown as AidocsWindow).__aidocs!.editor! as {
          chain: () => {
            selectAll: () => {
              setMark: (name: string, attrs: Record<string, unknown>) => { run: () => void }
            }
          }
          state: { doc: { content: { size: number } } }
          commands: { setTextSelection: (pos: number) => void }
        }
        ed.chain()
          .selectAll()
          .setMark('docTextStyle', { font: 'Calibri', fontAscii: 'Calibri' })
          .run()
        ed.commands.setTextSelection(ed.state.doc.content.size - 1)
      })

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

      const paste = async () => {
        await app.evaluate(
          ({ clipboard }, payload) => {
            clipboard.write({ html: payload.html, text: payload.text })
          },
          { html: WEB_TWO_PARAS, text: WEB_TEXT },
        )
        await editorPage.keyboard.press('Control+v')
        await editorPage.waitForTimeout(400)
      }

      // default = keep source: the site's concrete font, chip appears
      await paste()
      let state = await runs()
      expect(state.find((r) => r.text.includes('web one'))?.font).toBe('Arial')
      expect(state.find((r) => r.text === 'strong')?.bold).toBe(true)
      const chip = editorPage.locator('[data-testid="paste-options-chip"]')
      await expect(chip).toBeVisible()

      // merge formatting: caret font, emphasis kept
      await chip.locator('.paste-chip-btn').click()
      await chip.getByRole('menuitemradio', { name: 'Merge Formatting' }).click()
      await editorPage.waitForTimeout(300)
      state = await runs()
      expect(state.find((r) => r.text.includes('web one'))?.font).toBe('Calibri')
      expect(state.find((r) => r.text.includes('web two'))?.font).toBe('Calibri')
      expect(state.find((r) => r.text === 'strong')?.bold).toBe(true)

      // keep text only: plain caret-formatted text, emphasis gone
      await chip.locator('.paste-chip-btn').click()
      await chip.getByRole('menuitemradio', { name: 'Keep Text Only' }).click()
      await editorPage.waitForTimeout(300)
      state = await runs()
      expect(state.find((r) => r.text.includes('web one'))?.font).toBe('Calibri')
      expect(state.some((r) => r.bold)).toBe(false)

      // back to keep source: the concrete site font returns
      await chip.locator('.paste-chip-btn').click()
      await chip.getByRole('menuitemradio', { name: 'Keep Source Formatting' }).click()
      await editorPage.waitForTimeout(300)
      state = await runs()
      expect(state.find((r) => r.text.includes('web one'))?.font).toBe('Arial')

      // remember merge as the default: the NEXT paste applies it directly
      await chip.locator('.paste-chip-btn').click()
      await chip.getByRole('menuitemradio', { name: 'Merge Formatting' }).click()
      await editorPage.waitForTimeout(200)
      await chip.locator('.paste-chip-btn').click()
      await chip.locator('.paste-chip-remember input').check()
      await editorPage.keyboard.press('Escape')
      await editorPage.keyboard.press('Escape')
      await expect(chip).toBeHidden()
      // focus back in the document before driving the caret
      await editorPage.locator('.doc-page').click()
      await editorPage.keyboard.press('Control+End')
      await editorPage.keyboard.press('Enter')
      await paste()
      state = await runs()
      const second = state.filter((r) => r.text.includes('web two'))
      expect(second[second.length - 1]?.font).toBe('Calibri')

      // typing hides the chip
      await expect(chip).toBeVisible()
      await editorPage.keyboard.type('x')
      await expect(chip).toBeHidden()
    } finally {
      await closeAndSaveVideo(launched)
    }
  })
})
