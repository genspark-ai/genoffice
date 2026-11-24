/**
 * Word's Synonyms on the right-click menu is a submenu: candidate words, then
 * Thesaurus…. Picking a candidate replaces the selected word in place.
 */
import { test, expect } from '@playwright/test'
import type { Page } from '@playwright/test'
import { createServer, type Server } from 'node:http'
import { mkdtemp, writeFile } from 'node:fs/promises'
import type { AddressInfo } from 'node:net'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const POLL = { timeout: 15_000, intervals: [250, 500, 1000] }

function startModelStub(asked: string[]): Promise<Server> {
  const server = createServer((req, res) => {
    let body = ''
    req.on('data', (chunk) => (body += chunk))
    req.on('end', () => {
      asked.push(body)
      res.writeHead(200, { 'content-type': 'application/json' })
      res.end(
        JSON.stringify({
          choices: [
            {
              message: { role: 'assistant', content: '1. rapid\n2. swift\n3. speedy' },
              finish_reason: 'stop',
            },
          ],
        }),
      )
    })
  })
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve(server)))
}

const selectWord = (page: Page, word: string) =>
  page.evaluate((w) => {
    const ed = (window as unknown as { __aidocs: { editor: any } }).__aidocs.editor
    let hit: { from: number; to: number } | null = null
    ed.state.doc.descendants((node: any, pos: number) => {
      if (hit || !node.isText) return
      const idx = (node.text as string).indexOf(w)
      if (idx !== -1) hit = { from: pos + idx, to: pos + idx + w.length }
    })
    if (hit) ed.chain().focus().setTextSelection(hit).run()
    return hit
  }, word)

const selectionCenter = (page: Page) =>
  page.evaluate(() => {
    const r = window.getSelection()!.getRangeAt(0).getBoundingClientRect()
    return { x: r.left + r.width / 2, y: r.top + r.height / 2 }
  })

test('Synonyms opens a submenu of candidates and a pick replaces the word', async () => {
  test.setTimeout(120_000)
  const asked: string[] = []
  const server = await startModelStub(asked)
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-e2e-'))
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = {
    apiKey: 'stub',
    model: 'stub-model',
    baseUrl: `http://127.0.0.1:${(server.address() as AddressInfo).port}/v1`,
  }
  await writeFile(
    join(userDataDir, 'ai-settings.json'),
    JSON.stringify(resolveCustomEndpoints(settings)),
  )
  const launched = await launchShell({ userDataDir, onboardingSeen: true, videoDir: 'synonyms' })
  const { app, page } = launched
  try {
    await page.locator('.quick-card').first().click()
    const editor = await waitForPageWithUrl(app, '://docs/')
    const docPage = editor.locator('.doc-page[contenteditable="true"]')
    await docPage.waitFor()
    await docPage.click()
    await editor.keyboard.type('A quick brown fox.', { delay: 15 })
    await expect(docPage).toContainText('brown fox.')

    expect(await selectWord(editor, 'quick')).not.toBeNull()
    const at = await selectionCenter(editor)
    await editor.mouse.click(at.x, at.y, { button: 'right' })
    const menu = editor.locator('.ctx-menu')
    await menu.waitFor()
    // a right-click alone asks nothing
    expect(asked).toHaveLength(0)

    await menu.locator('.ctx-item', { hasText: 'Synonyms' }).hover()
    const sub = menu.locator('.ctx-submenu')
    await expect(sub.locator('.ctx-label')).toHaveText(
      ['rapid', 'swift', 'speedy', 'Thesaurus…'],
      POLL,
    )
    expect(asked).toHaveLength(1)
    expect(asked[0]).toContain('quick')
    if (process.env.SYNONYMS_SHOT) await editor.screenshot({ path: process.env.SYNONYMS_SHOT })

    await sub.locator('.ctx-item', { hasText: 'swift' }).click()
    await expect(menu).toHaveCount(0)
    await expect(docPage).toContainText('A swift brown fox.')
  } finally {
    await closeAndSaveVideo(launched)
    server.close()
  }
})
