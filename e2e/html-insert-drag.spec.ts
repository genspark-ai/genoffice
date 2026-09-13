import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

test.describe('html editor: insert, resize, drag', () => {
  test('Insert menu adds elements, handles resize, dragging the grip reorders', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-html-'))
    const htmlPath = join(dir, 'layout.html')
    const source =
      '<!doctype html>\n<html>\n<body>\n<section class="hero">\n  <h1 id="title">Hello</h1>\n  <p class="lead">First line.</p>\n  <p class="note">Second line.</p>\n</section>\n</body>\n</html>\n'
    await writeFile(htmlPath, source)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'html-insert-drag',
      openFile: htmlPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, '://html/')
      await expect(editorPage.locator('.ribbon-body')).toBeVisible()
      const frame = editorPage.frameLocator('.preview-frame')
      const content = editorPage.locator('.source-editor .cm-content')
      await editorPage.locator('.ribbon').getByRole('tab', { name: /Split/ }).click()

      // nothing selected: the new paragraph lands at the end of the body and opens for typing
      const ribbon = editorPage.locator('.ribbon')
      await ribbon.getByRole('button', { name: 'Paragraph', exact: true }).click()
      await expect(content).toContainText(/<\/section>\s*<p>Type your text here\.<\/p>\s*<\/body>/)
      const fresh = frame.locator('body > p')
      await expect(fresh).toHaveText('Type your text here.')
      await expect(fresh).toHaveAttribute('contenteditable', 'plaintext-only')
      await expect(editorPage.locator('.crumb.current')).toHaveText('p')
      await frame.locator('body').press('Escape')
      // the split view halves the stage: park the style panel so it does not sit over the page
      // (synthetic click: the panel itself may cover the toolbar toggle)
      const float = editorPage.locator('.hx-float')
      await float.getByRole('button', { name: /Style panel/ }).dispatchEvent('click')
      await expect(editorPage.locator('.hx-panel')).toHaveCount(0)

      // with a selection: right after it
      await frame.locator('h1#title').click()
      await expect(editorPage.locator('.crumb.current')).toHaveText('h1#title')
      // the less common kinds sit under "More"
      await ribbon.getByRole('button', { name: 'More', exact: true }).click()
      await editorPage.getByRole('menuitem', { name: 'Divider' }).click()
      await expect(content).toContainText(/<h1 id="title">Hello<\/h1>\s*<hr>\s*<p class="lead">/)
      await expect(frame.locator('section.hero > hr')).toHaveCount(1)

      // resize: dragging the right handle writes an inline width
      await frame.locator('p.lead').click()
      await expect(editorPage.locator('.crumb.current')).toHaveText('p.lead')
      const handle = frame.locator('[data-gx-inspector-handle="e"]')
      await expect(handle).toBeVisible()
      const box = (await handle.boundingBox())!
      const startWidth = await frame
        .locator('p.lead')
        .evaluate((el) => el.getBoundingClientRect().width)
      await editorPage.mouse.move(box.x + box.width / 2, box.y + box.height / 2)
      await editorPage.mouse.down()
      await editorPage.mouse.move(box.x + box.width / 2 - 60, box.y + box.height / 2, { steps: 6 })
      await editorPage.mouse.up()
      await expect(content).toContainText(/<p class="lead" style="width: (\d+)px">/)
      const written = Number(/style="width: (\d+)px"/.exec(await content.innerText())![1])
      expect(Math.abs(written - (startWidth - 60))).toBeLessThanOrEqual(2)
      await editorPage.screenshot({ path: screenshotPath('html-resize') })

      // drag: grab the grip beside the selection and drop the note above the lead paragraph
      await expect(frame.locator('p.lead')).toHaveAttribute('style', /width/)
      await frame.locator('p.note').click()
      await expect(editorPage.locator('.crumb.current')).toHaveText('p.note')
      await expect(frame.locator('[data-gx-inspector-label]')).toHaveText('p.note')
      const grip = frame.locator('[data-gx-inspector-grip]')
      await expect(grip).toBeVisible()
      const gripBox = (await grip.boundingBox())!
      const leadBox = (await frame.locator('p.lead').boundingBox())!
      await editorPage.mouse.move(gripBox.x + gripBox.width / 2, gripBox.y + gripBox.height / 2)
      await editorPage.mouse.down()
      await editorPage.mouse.move(leadBox.x + leadBox.width / 2, leadBox.y + 3, { steps: 8 })
      await expect(frame.locator('[data-gx-inspector-drop]')).toBeVisible()
      await editorPage.mouse.up()
      await expect(content).toContainText(
        /<p class="note">Second line\.<\/p>\s*<p class="lead" style="width:/,
      )
      await expect(frame.locator('section.hero > p').first()).toHaveClass('note')
      await expect(editorPage.locator('.crumb.current')).toHaveText('p.note')
      await editorPage.screenshot({ path: screenshotPath('html-drag-reorder') })
    } finally {
      await closeAndSaveVideo(launched, 'html-insert-drag')
    }
  })
})
