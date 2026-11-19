import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Web-HTML paste fonts (r181): text copied from a website pasted into a
 * document must not land in the theme font.
 *  - a fragment whose family chain has no concrete face (`sans-serif` —
 *    generic-only sites) takes the insertion point's font like typing;
 *  - a whole-paragraph copy carrying the site's styles on the <p> keeps its
 *    concrete font (block styles are pushed down to the spans the mark
 *    parser reads);
 *  - our own clipboard HTML (data-pm-slice) keeps its exact marks.
 */

interface EditorRun {
  text: string
  font: string | null
  color: string | null
}

interface AidocsWindow {
  __aidocs?: {
    editor?: {
      state: { doc: unknown }
      chain: () => unknown
      commands: Record<string, (...args: unknown[]) => unknown>
    }
  }
}

const GENERIC_SPAN = `<meta charset='utf-8'><span style="color: rgb(32, 33, 34); font-family: sans-serif; font-size: 14px;">web generic</span>`
const BLOCK_ROBOTO = `<meta charset='utf-8'><p style="margin: 0px 0px 16px; font-family: Roboto, sans-serif; font-size: 14px;">web roboto block</p>`
const INTERNAL_BARE = `<div data-pm-slice="1 1 []"><p>internal bare</p></div>`

test.describe('docs web-HTML paste fonts', () => {
  test('web pastes follow the caret font when generic and keep concrete site fonts', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-web-paste-font' })
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
      await editorPage.keyboard.type('avant apres', { delay: 10 })
      // the field docs this guards: surrounding text carries explicit Calibri
      await editorPage.evaluate(() => {
        const ed = (window as unknown as AidocsWindow).__aidocs!.editor! as unknown as {
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
      for (let i = 0; i < 'apres'.length; i++) await editorPage.keyboard.press('ArrowLeft')

      const runs = async (): Promise<EditorRun[]> =>
        editorPage.evaluate(() => {
          const ed = (window as unknown as AidocsWindow).__aidocs!.editor! as unknown as {
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
          const out: Array<{ text: string; font: string | null; color: string | null }> = []
          ed.state.doc.descendants((node) => {
            if (node.isText) {
              const style = node.marks.find((m) => m.type.name === 'docTextStyle')
              out.push({
                text: node.text ?? '',
                font: (style?.attrs.font as string | null) ?? null,
                color: (style?.attrs.color as string | null) ?? null,
              })
            }
            return true
          })
          return out
        })

      const paste = async (html: string, text: string): Promise<void> => {
        await app.evaluate(
          ({ clipboard }, payload) => {
            clipboard.write({ html: payload.html, text: payload.text })
          },
          { html, text },
        )
        await editorPage.keyboard.press('Control+v')
        await editorPage.waitForTimeout(400)
      }

      // 1) generic-only chain mid-paragraph: caret font, source color kept
      await paste(GENERIC_SPAN, 'web generic')
      let state = await runs()
      const generic = state.find((r) => r.text === 'web generic')
      expect(generic?.font).toBe('Calibri')
      expect(generic?.color).toBe('202122')

      // 2) whole-paragraph copy on an empty line: the block's concrete font
      await editorPage.keyboard.press('Control+End')
      await editorPage.keyboard.press('Enter')
      await paste(BLOCK_ROBOTO, 'web roboto block')
      state = await runs()
      expect(state.find((r) => r.text.includes('web roboto'))?.font).toBe('Roboto')

      // 3) internal clipboard HTML stays untouched (no caret-font fill)
      await editorPage.keyboard.press('Control+End')
      await editorPage.keyboard.press('Enter')
      await paste(INTERNAL_BARE, 'internal bare')
      state = await runs()
      expect(state.find((r) => r.text.includes('internal bare'))?.font).toBeNull()
    } finally {
      await closeAndSaveVideo(launched)
    }
  })
})
