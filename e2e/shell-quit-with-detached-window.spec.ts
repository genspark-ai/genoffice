/**
 * Quit with a document in its own window. The detached window's close guard
 * cancels the close to ask the document whether it is dirty, which also cancels
 * the quit; on macOS the windows then all closed and the app stayed running.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl } from './helpers'
import { docxWith, para } from './docs-objects-fixture'

const RUNS = 30

interface ShellWindow {
  aiOffice: { newDoc(): Promise<void> }
  aiOfficeTabs: {
    list(): Promise<Array<{ id: string; kind: string; title: string }>>
    detach(id: string): Promise<void>
  }
}

test(`quit exits the app with a clean document in a detached window, ${RUNS} times`, async () => {
  test.setTimeout(RUNS * 40_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-quit-detached-')))
  const docPath = join(dir, 'detached.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('Plain paragraph') + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
    ),
  )
  const outcomes: string[] = []
  for (let run = 0; run < RUNS; run++) {
    // every third run keeps a second, docked document, so both guards hold the quit
    const withDockedTab = run % 3 === 2
    const { app } = await launchShell({
      onboardingSeen: true,
      videoDir: 'shell-quit-with-detached-window',
      openFile: docPath,
    })
    const child = app.process()
    const exited = new Promise<void>((resolve) => child.once('exit', () => resolve()))
    try {
      const shell = await waitForPageWithUrl(app, 'shell/out')
      const docs = await waitForPageWithUrl(app, '://docs/')
      await expect(docs.locator('.doc-page', { hasText: 'Plain paragraph' }).first()).toBeVisible({
        timeout: 30_000,
      })
      if (withDockedTab) {
        await shell.evaluate(() => (window as unknown as ShellWindow).aiOffice.newDoc())
        await expect
          .poll(async () => {
            const tabs = await shell.evaluate(() =>
              (window as unknown as ShellWindow).aiOfficeTabs.list(),
            )
            return tabs.filter((tab) => tab.kind === 'docs').length
          })
          .toBe(2)
        // the new document is still being set up when its tab appears
        await expect
          .poll(async () => {
            const pages = app.windows().filter((page) => page.url().includes('://docs/'))
            const ready = await Promise.all(
              pages.map((page) => page.locator('.doc-page').first().isVisible()),
            )
            return ready.length === 2 && ready.every(Boolean)
          })
          .toBe(true)
        await docs.waitForTimeout(1000)
      }
      const tabs = await shell.evaluate(() =>
        (window as unknown as ShellWindow).aiOfficeTabs.list(),
      )
      const tab = tabs.find((entry) => entry.title === 'detached.docx')
      expect(tab, 'the document opened in a tab').toBeTruthy()
      await shell.evaluate(
        (id) => (window as unknown as ShellWindow).aiOfficeTabs.detach(id),
        tab!.id,
      )
      await expect
        .poll(() => app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length))
        .toBe(2)

      await app.evaluate(({ app: electronApp, dialog }) => {
        const asked: string[] = []
        ;(globalThis as { __prompts?: string[] }).__prompts = asked
        dialog.showMessageBox = (async (...args: unknown[]) => {
          const options = args.find(
            (arg): arg is Electron.MessageBoxOptions =>
              !!arg && typeof arg === 'object' && 'message' in arg,
          )
          asked.push(options?.message ?? '')
          return { response: options?.cancelId ?? 0, checkboxChecked: false }
        }) as typeof dialog.showMessageBox
        setTimeout(() => electronApp.quit(), 0)
      })
      const gone = await Promise.race([
        exited.then(() => true),
        new Promise<boolean>((resolve) => setTimeout(() => resolve(false), 15_000)),
      ])
      if (gone) {
        outcomes.push('exited')
        continue
      }
      const state = await app
        .evaluate(({ BrowserWindow }) => ({
          windows: BrowserWindow.getAllWindows().length,
          prompts: (globalThis as { __prompts?: string[] }).__prompts,
        }))
        .catch((err) => String(err))
      outcomes.push(`run ${run}${withDockedTab ? ' (docked tab)' : ''}: ${JSON.stringify(state)}`)
    } finally {
      if (child.exitCode === null && child.signalCode === null) {
        child.kill('SIGKILL')
        await exited
      }
    }
  }
  expect(outcomes).toEqual(Array.from({ length: RUNS }, () => 'exited'))
})
