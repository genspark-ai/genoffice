import { expect, test } from '@playwright/test'
import type { Editor } from '@tiptap/core'
import { openSource } from './helpers'
const rebaseSource = '<details>\n<img src="assets/old.png">\n</details>\n'
const source = 'Title\n=====\n\n* item  \n\n\n'

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
