import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

test.describe('html editor', () => {
  test('AI HTML quick card opens an html editor tab in preview view with a ribbon', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'new-html-tab' })
    const { app, page } = launched
    try {
      const card = page.locator('.quick-card', { hasText: 'AI HTML' })
      await expect(card).toHaveCount(1)
      await card.click()

      const editorTab = page.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(editorTab).toHaveCount(1)
      await expect(editorTab).toHaveClass(/active/)

      const editorPage = await waitForPageWithUrl(app, 'html/out')
      await expect(editorPage.locator('.ribbon-body')).toBeVisible()
      await expect(editorPage.locator('.preview-frame')).toBeVisible()
      // blank document: the AI panel offers the design / write intent cards and swaps its starters
      const copilot = editorPage.locator('.copilot')
      if (!(await copilot.isVisible())) await editorPage.locator('.ai-rail').click()
      const cards = copilot.getByRole('radio')
      await expect(cards).toHaveCount(2)
      await expect(cards.first()).toHaveAttribute('aria-checked', 'true')
      await cards.nth(1).click()
      await expect(cards.nth(1)).toHaveAttribute('aria-checked', 'true')
      await expect(copilot.locator('.ai-starter').first()).toHaveText(/Write an article/)
      // preview-only by default; the source pane stays mounted but hidden
      await expect(editorPage.locator('.pane-source')).toBeHidden()
      await editorPage.locator('.rb-view', { hasText: /Source/ }).click()
      await expect(editorPage.locator('.source-editor .cm-content')).toBeVisible()
      await editorPage.screenshot({ path: screenshotPath('new-html-editor') })
    } finally {
      await closeAndSaveVideo(launched, 'new-html-tab')
    }
  })

  test('opens an .html file from argv, renders it in the preview, edits and saves it back', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-html-'))
    const htmlPath = join(dir, 'page.html')
    const source =
      '<!doctype html>\n<html>\n<body>\n<h1 class="hero" style="color: rgb(200, 0, 0)">Hello</h1>\n</body>\n</html>\n'
    await writeFile(htmlPath, source)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'open-html-file',
      openFile: htmlPath,
    })
    const { app } = launched
    try {
      const shellPage = await waitForPageWithUrl(app, 'shell/out')
      const editorTab = shellPage.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(editorTab).toHaveCount(1)
      await expect(editorTab).toContainText('page.html')

      const editorPage = await waitForPageWithUrl(app, 'html/out')
      await expect(editorPage.locator('.ribbon-body')).toBeVisible()
      // preview is the default view; the source pane is shown on demand
      await expect(editorPage.locator('.pane-source')).toBeHidden()
      await editorPage.locator('.rb-view', { hasText: /Split/ }).click()
      const content = editorPage.locator('.source-editor .cm-content')
      await expect(content).toContainText('class="hero"')

      // the sandboxed preview really renders the document through html-preview://
      const heading = editorPage.frameLocator('.preview-frame').locator('h1.hero')
      await expect(heading).toHaveText('Hello')
      await expect(heading).toHaveCSS('color', 'rgb(200, 0, 0)')

      // edit in the source pane, the preview follows, save with ⌘/Ctrl+S
      await content.click()
      await editorPage.keyboard.press('ControlOrMeta+End')
      await editorPage.keyboard.type('<!-- appended -->')
      await expect(editorPage.locator('.status-save')).toHaveText(/Unsaved/)
      await editorPage.keyboard.press('ControlOrMeta+s')
      await expect(editorPage.locator('.status-save')).toHaveText(/Saved/)
      await editorPage.screenshot({ path: screenshotPath('open-html-saved') })

      const saved = await readFile(htmlPath, 'utf8')
      expect(saved).toBe(source.trimEnd() + '<!-- appended -->\n')

      // the source editor survives a round trip through preview-only, undo history included
      await editorPage.locator('.rb-view', { hasText: /Preview/ }).click()
      await expect(editorPage.locator('.pane-source')).toBeHidden()
      await editorPage.locator('.rb-view', { hasText: /Split/ }).click()
      await content.click()
      await editorPage.keyboard.press('ControlOrMeta+z')
      await expect(content).not.toContainText('appended')
    } finally {
      await closeAndSaveVideo(launched, 'open-html-file')
    }
  })

  test('relative stylesheets and images next to the file load in the preview', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-html-'))
    await writeFile(join(dir, 'site.css'), 'h1 { color: rgb(0, 0, 200); }\n')
    // 1x1 PNG
    const png = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
      'base64',
    )
    await writeFile(join(dir, 'dot.png'), png)
    const htmlPath = join(dir, 'assets-page.html')
    await writeFile(
      htmlPath,
      '<!doctype html><html><head><link rel="stylesheet" href="site.css"></head><body><h1>Styled</h1><img id="dot" src="dot.png"></body></html>\n',
    )

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'html-relative-assets',
      openFile: htmlPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, 'html/out')
      const frame = editorPage.frameLocator('.preview-frame')
      await expect(frame.locator('h1')).toHaveCSS('color', 'rgb(0, 0, 200)')
      await expect
        .poll(() => frame.locator('#dot').evaluate((el) => (el as HTMLImageElement).naturalWidth))
        .toBe(1)

      // an <img> gets the image toolbar (replace / flip / rotate), not the text one
      await frame.locator('#dot').dispatchEvent('click')
      const float = editorPage.locator('.hx-float')
      await expect(float.getByRole('button', { name: /Flip horizontally/ })).toBeVisible()
      await expect(float.getByRole('button', { name: /^Bold$/ })).toHaveCount(0)
      const panel = editorPage.locator('.preview-stage .hx-panel')
      await expect(panel.locator('.hx-panel-tag')).toHaveText('<img>')
      await expect(panel.locator('.hx-panel-section', { hasText: /Typography/ })).toHaveCount(0)
      await expect(panel.locator('.hx-panel-section', { hasText: /^Image$/ })).toHaveCount(1)
      await editorPage.screenshot({ path: screenshotPath('html-image-toolbar') })

      // flip then rotate back-to-back: the second click composes on the first poke (not on the
      // snapshot still round-tripping through the frame) and both land as one inline transform
      await float.getByRole('button', { name: /Flip horizontally/ }).click()
      await float.getByRole('button', { name: /Rotate right/ }).click()
      await expect(frame.locator('#dot')).toHaveAttribute('style', /rotate\(90deg\) scale\(-1, 1\)/)
      await editorPage.locator('.rb-view', { hasText: /Split/ }).click()
      await expect(editorPage.locator('.source-editor .cm-content')).toContainText(
        'style="transform: rotate(90deg) scale(-1, 1)"',
      )

      // crop reads the picture next to the document, bakes a PNG into assets/ and repoints src
      await float.getByRole('button', { name: /^Crop$/ }).click()
      const cropDialog = editorPage.getByRole('dialog', { name: /Crop/ })
      await expect(cropDialog).toBeVisible()
      await editorPage.screenshot({ path: screenshotPath('html-image-crop') })
      await cropDialog.getByRole('button', { name: /^Apply$/ }).click()
      await expect(cropDialog).toHaveCount(0)
      await expect(frame.locator('#dot')).toHaveAttribute('src', /^assets\/.+\.png$/)
      await expect(editorPage.locator('.source-editor .cm-content')).toContainText('src="assets/')
    } finally {
      await closeAndSaveVideo(launched, 'html-relative-assets')
    }
  })

  test('saving without edits keeps BOM, CRLF and the missing trailing newline byte-identical', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-html-'))
    const htmlPath = join(dir, 'legacy.htm')
    const source = '﻿<html>\r\n<body>\r\n<p>café &amp; <b>bold</p>\r\n</body>\r\n</html>'
    await writeFile(htmlPath, source, 'utf8')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'html-identity-save',
      openFile: htmlPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, 'html/out')
      await expect(editorPage.locator('.ribbon-body')).toBeVisible()
      await editorPage.locator('.rb-view', { hasText: /Source/ }).click()
      await editorPage.locator('.source-editor .cm-content').click()
      await editorPage.keyboard.press('ControlOrMeta+s')
      await expect(editorPage.locator('.status-save')).toHaveText(/Saved/)
      expect(await readFile(htmlPath, 'utf8')).toBe(source)
    } finally {
      await closeAndSaveVideo(launched, 'html-identity-save')
    }
  })

  test('AI panel toggles from the toolbar and accepts an instruction', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-html-'))
    const htmlPath = join(dir, 'ai.html')
    await writeFile(htmlPath, '<html><body><h1>Topic</h1><p>Body.</p></body></html>\n')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'html-ai-panel',
      openFile: htmlPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, 'html/out')
      await expect(editorPage.locator('.ribbon-body')).toBeVisible()
      const toggle = editorPage.locator('.ribbon .ai-entry').first()
      // the panel state is remembered; normalize to closed first
      if (await editorPage.locator('.copilot').isVisible()) await toggle.click()
      await expect(editorPage.locator('.copilot')).toBeHidden()
      await editorPage.locator('.ai-rail').click()
      await expect(editorPage.locator('.copilot')).toBeVisible()
      // the panel's right border is the divider to the canvas: it must not be clipped by the dock
      await expect
        .poll(() =>
          editorPage.evaluate(() => {
            const dock = document.querySelector('.ai-dock')!.getBoundingClientRect()
            const panel = document.querySelector('.copilot')!.getBoundingClientRect()
            return panel.right <= dock.right
          }),
        )
        .toBe(true)

      const composer = editorPage.locator('.copilot textarea')
      // a pasted bitmap becomes an image attachment chip; sending echoes it on the bubble
      await expect(editorPage.locator('.copilot .ai-attach-btn')).toBeVisible()
      await composer.evaluate((ta) => {
        const b64 =
          'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAEklEQVR4nGP4z8DwHwyBNAMDAB9dB/8dB8xVAAAAAElFTkSuQmCC'
        const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))
        const dt = new DataTransfer()
        dt.items.add(new File([bytes], 'shot.png', { type: 'image/png' }))
        ta.dispatchEvent(
          new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }),
        )
      })
      await expect(editorPage.locator('.copilot .ai-attachments .ai-attachment-thumb')).toHaveCount(
        1,
      )
      await composer.fill('Rename the heading to Summary')
      await composer.press('Enter')
      // the instruction lands as a sent user message (the model reply itself needs credentials)
      await expect(editorPage.locator('.ai-msg-user')).toContainText('Rename the heading')
      await expect(editorPage.locator('.ai-msg-user .ai-msg-attachments img')).toHaveCount(1)
      await expect(editorPage.locator('.copilot .ai-attachments')).toHaveCount(0)
      await editorPage.screenshot({ path: screenshotPath('html-ai-panel') })

      // collapsing hides the panel but keeps it mounted: the live transcript is still there on reopen
      await toggle.click()
      await expect(editorPage.locator('.copilot')).toBeHidden()
      await toggle.click()
      await expect(editorPage.locator('.ai-msg-user')).toContainText('Rename the heading')

      // page-wide AI entries: Theme opens a direction menu, a pick sends the instruction
      // (or parks it in the composer while a run is still busy); Summarize sends right away
      await editorPage
        .locator('.ribbon')
        .getByRole('button', { name: /^Theme$/ })
        .click()
      const themeMenu = editorPage.locator('.rb-menu')
      await expect(themeMenu.getByRole('menuitem')).toHaveCount(5)
      await themeMenu.getByRole('menuitem', { name: /Minimal/ }).click()
      await expect(themeMenu).toBeHidden()
      const instructionText = async () =>
        (await editorPage.locator('.ai-msg-user').last().textContent()) +
        (await editorPage.locator('.copilot textarea').inputValue())
      await expect.poll(instructionText).toMatch(/Minimal/)
      await editorPage
        .locator('.ribbon')
        .getByRole('button', { name: /^AI Summarize$/ })
        .click()
      await expect.poll(instructionText).toMatch(/Summarize/)
    } finally {
      await closeAndSaveVideo(launched, 'html-ai-panel')
    }
  })

  test('preview inspector: click selects, double-click edits text, toolbar deletes, Ask AI drafts', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-html-'))
    const htmlPath = join(dir, 'inspect.html')
    const source =
      '<!doctype html>\n<html>\n<body>\n<section class="hero">\n  <h1 id="title">Hello</h1>\n  <p class="lead">First line.</p>\n  <p class="note">Second line.</p>\n</section>\n</body>\n</html>\n'
    await writeFile(htmlPath, source)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'html-inspector',
      openFile: htmlPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, 'html/out')
      await expect(editorPage.locator('.ribbon-body')).toBeVisible()
      const frame = editorPage.frameLocator('.preview-frame')
      // the breadcrumb is a source-side affordance: it only renders in the split and source views
      await expect(editorPage.locator('.crumbs')).toHaveCount(0)
      await editorPage.locator('.ribbon').getByRole('tab', { name: /Split/ }).click()
      // click → selection: breadcrumb shows the chain and the source pane selects the element
      await frame.locator('h1#title').click()
      await expect(editorPage.locator('.crumb.current')).toHaveText('h1#title')
      await expect(editorPage.locator('.crumb')).toHaveCount(2)
      const float = editorPage.locator('.hx-float')
      await expect(float.getByRole('button', { name: /Delete element/ })).toBeEnabled()
      // the split view halves the stage: park the style panel so it does not sit over the page
      const panelToggle = float.getByRole('button', { name: /Style panel/ })
      await expect(editorPage.locator('.hx-panel')).toHaveCount(1)
      await panelToggle.click()
      await expect(editorPage.locator('.hx-panel')).toHaveCount(0)
      await editorPage.screenshot({ path: screenshotPath('html-inspector-select') })

      // breadcrumb click selects the ancestor; body/head/html never appear as crumbs
      await editorPage.locator('.crumb', { hasText: 'section.hero' }).click()
      await expect(editorPage.locator('.crumb.current')).toHaveText('section.hero')
      await expect(editorPage.locator('.crumb', { hasText: /^body$/ })).toHaveCount(0)

      // Cmd/Ctrl+B on a DOM text selection wraps the right characters even across an entity
      await frame.locator('p.note').click()
      await frame.locator('p.note').evaluate((el) => {
        const node = el.firstChild as Text
        const sel = el.ownerDocument.getSelection()!
        sel.setBaseAndExtent(node, 7, node, 12)
      })
      await frame.locator('body').press('ControlOrMeta+b')
      await expect(editorPage.locator('.source-editor .cm-content')).toContainText(
        '<p class="note">Second <strong>line.</strong></p>',
      )
      await expect(frame.locator('p.note strong')).toHaveText('line.')

      // move up among same-tag siblings: the selection follows the moved element, not its old slot
      await frame.locator('p.note').click()
      await float.getByRole('button', { name: /Move up/ }).click()
      await expect(editorPage.locator('.source-editor .cm-content')).toContainText(
        '<p class="note">Second <strong>line.</strong></p>\n  <p class="lead">First line.</p>',
      )
      await expect(editorPage.locator('.crumb.current')).toHaveText('p.note')
      // wait for the reloaded frame (new order) before clicking into it again
      await expect(frame.locator('section.hero > p').first()).toHaveClass('note')
      await frame.locator('p.note').click()
      await float.getByRole('button', { name: /Move down/ }).click()
      await expect(editorPage.locator('.source-editor .cm-content')).toContainText(
        '<p class="lead">First line.</p>\n  <p class="note">Second <strong>line.</strong></p>',
      )
      await expect(editorPage.locator('.crumb.current')).toHaveText('p.note')
      await expect(frame.locator('section.hero > p').first()).toHaveClass('lead')

      // double-click → inline edit → written back to the source as a text-node patch
      // (select it first via a synthetic click: the previous selection's floating toolbar may sit right over it)
      await frame.locator('p.lead').dispatchEvent('click')
      await expect(editorPage.locator('.crumb.current')).toHaveText('p.lead')
      await frame.locator('p.lead').dblclick()
      await editorPage.waitForTimeout(150)
      await frame.locator('p.lead').press('ControlOrMeta+a')
      await frame.locator('p.lead').pressSequentially('Edited & saved')
      await frame.locator('p.lead').press('ControlOrMeta+Enter')
      await expect(editorPage.locator('.source-editor .cm-content')).toContainText(
        '<p class="lead">Edited &amp; saved</p>',
      )

      // toolbar delete removes the selected element and nothing else
      await frame.locator('p.note').click()
      await expect(editorPage.locator('.crumb.current')).toHaveText('p.note')
      await float.getByRole('button', { name: /Delete element/ }).click()
      await expect(editorPage.locator('.source-editor .cm-content')).not.toContainText(
        'Second line.',
      )
      await expect(editorPage.locator('.source-editor .cm-content')).toContainText(
        'Edited &amp; saved',
      )

      // back to the full-width preview: the floating toolbar follows the selection; style pokes preview
      // live and land in the source as one op; the style panel floats over the right edge of the stage
      await editorPage
        .locator('.ribbon')
        .getByRole('tab', { name: /Preview/ })
        .click()
      await expect(editorPage.locator('.crumbs')).toHaveCount(0)
      await frame.locator('h1#title').click()
      await panelToggle.click()
      await expect(editorPage.locator('.preview-stage .hx-panel')).toBeVisible()
      await editorPage.screenshot({ path: screenshotPath('html-inspector-style-panel') })
      await expect(float).toBeVisible()
      await expect(editorPage.locator('.hx-panel-tag')).toHaveText('<h1>')
      await float.getByRole('button', { name: /Increase font size/ }).click()
      await float.getByRole('button', { name: /^Center$/ }).click()
      await expect(editorPage.locator('.source-editor .cm-content')).toContainText(
        /<h1 id="title" style="font-size: \d+px; text-align: center">Hello<\/h1>/,
      )
      await expect(frame.locator('h1#title')).toBeVisible()
      await expect(editorPage.locator('.hx-panel-tag')).toHaveText('<h1>')
      await expect(frame.locator('h1#title')).toHaveCSS('text-align', 'center')

      // the color button opens the shared Word-style palette; a swatch pick applies and closes it
      await float.getByRole('button', { name: /^Color$/ }).click()
      await editorPage.locator('.hx-color-pop .gcp-swatch[title="Red"]').click()
      await expect(editorPage.locator('.hx-color-pop')).toHaveCount(0)
      await expect(frame.locator('h1#title')).toHaveCSS('color', 'rgb(255, 0, 0)')

      // clicking the bare page (a synthetic click: the fixture's body is only a few pixels of margin)
      // clears the selection and hides the style panel
      await frame.locator('body').dispatchEvent('click')
      await expect(float).toBeHidden()
      await expect(editorPage.locator('.hx-panel')).toHaveCount(0)

      // the floating toolbar's Ask AI (the single canvas entry) opens the element popover;
      // "Add to queue" parks the instruction in the AI panel
      await frame.locator('h1#title').click()
      await expect(editorPage.locator('.hx-panel-tag')).toHaveText('<h1>')
      await float.getByRole('button', { name: /Ask AI/ }).click()
      const askPop = editorPage.locator('.ai-ask-pop')
      await expect(askPop).toBeVisible()
      await expect(askPop.locator('.ai-ask-pop-sub')).toContainText('<h1>')
      await askPop.locator('.ai-ask-pop-input').fill('make it shorter')
      await askPop.locator('.ai-ask-confirm').click()
      await expect(askPop).toBeHidden()
      const queueRow = editorPage.locator('.copilot .ai-queue-row')
      await expect(queueRow).toHaveCount(1)
      await expect(queueRow).toContainText('make it shorter')
      await expect(frame.locator('[data-gx-inspector-pin]')).toHaveText('1')
      // the pin reopens the instruction for editing; removing it clears the queue
      await frame.locator('[data-gx-inspector-pin]').click()
      await expect(askPop.locator('.ai-ask-pop-input')).toHaveValue('make it shorter')
      await askPop.locator('.ai-ask-cancel').click()
      await expect(queueRow).toHaveCount(0)

      // Present → In this tab hides the editing chrome and links work; Exit returns to Edit
      await editorPage
        .locator('.ribbon-body')
        .getByRole('button', { name: /Present/ })
        .click()
      await editorPage
        .locator('.rb-menu')
        .getByRole('menuitem', { name: /In this tab/ })
        .click()
      await expect(editorPage.locator('.ribbon-body')).toBeHidden()
      await expect(editorPage.locator('.present-exit')).toBeVisible()
      await editorPage.locator('.present-exit').click()
      await expect(editorPage.locator('.ribbon-body')).toBeVisible()

      await editorPage.keyboard.press('ControlOrMeta+s')
      await expect(editorPage.locator('.status-save')).toHaveText(/Saved/)
      const saved = await readFile(htmlPath, 'utf8')
      expect(saved).toBe(
        '<!doctype html>\n<html>\n<body>\n<section class="hero">\n  <h1 id="title" style="font-size: 34px; text-align: center; color: #ff0000">Hello</h1>\n  <p class="lead">Edited &amp; saved</p>\n</section>\n</body>\n</html>\n',
      )
    } finally {
      await closeAndSaveVideo(launched, 'html-inspector')
    }
  })

  test('Ctrl+Z inside the preview undoes the last edit; pinch over the frame zooms the stage', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-html-'))
    const htmlPath = join(dir, 'undo.html')
    await writeFile(
      htmlPath,
      '<!doctype html>\n<html>\n<body>\n<h1>Alpha</h1>\n<p class="note">Beta</p>\n</body>\n</html>\n',
    )
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'html-undo-zoom',
      openFile: htmlPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, 'html/out')
      await expect(editorPage.locator('.ribbon-body')).toBeVisible()
      const frame = editorPage.frameLocator('.preview-frame')
      const undoBtn = editorPage.locator('.ribbon').getByRole('button', { name: /^Undo$/ })
      await expect(frame.locator('p.note')).toHaveText('Beta')
      await expect(undoBtn).toBeDisabled()

      // Delete inside the frame removes the element; the frame reloads without it
      await frame.locator('p.note').click()
      await frame.locator('body').press('Delete')
      await expect(frame.locator('p.note')).toHaveCount(0)
      await expect(undoBtn).toBeEnabled()

      // the frame owns keyboard focus, so Cmd/Ctrl+Z must be relayed to the host's history
      await frame.locator('body').press('ControlOrMeta+z')
      await expect(frame.locator('p.note')).toHaveText('Beta')
      await expect(undoBtn).toBeDisabled()
      await frame.locator('body').press('ControlOrMeta+Shift+z')
      await expect(frame.locator('p.note')).toHaveCount(0)
      await frame.locator('body').press('ControlOrMeta+z')
      await expect(frame.locator('p.note')).toHaveText('Beta')

      // trackpad pinch arrives as ctrl+wheel inside the sandboxed frame
      await expect(editorPage.locator('.zoom-value')).toHaveText('100%')
      await frame.locator('body').evaluate((body) => {
        body.dispatchEvent(
          new WheelEvent('wheel', { deltaY: -100, ctrlKey: true, bubbles: true, cancelable: true }),
        )
      })
      await expect(editorPage.locator('.zoom-value')).toHaveText('160%')
      // zoom shortcuts are relayed too while the frame keeps focus
      await frame.locator('body').press('ControlOrMeta+0')
      await expect(editorPage.locator('.zoom-value')).toHaveText('100%')
      await frame.locator('body').press('ControlOrMeta+-')
      await expect(editorPage.locator('.zoom-value')).toHaveText('90%')
    } finally {
      await closeAndSaveVideo(launched, 'html-undo-zoom')
    }
  })

  test('Ctrl+F searches the source pane and Replace All saves the rewritten markup', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-html-'))
    const htmlPath = join(dir, 'find.html')
    const source =
      '<!doctype html>\n<html>\n<body>\n<h1>Alpha</h1>\n<p>alpha and alpha</p>\n</body>\n</html>\n'
    await writeFile(htmlPath, source)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'html-find-replace',
      openFile: htmlPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, 'html/out')
      await expect(editorPage.frameLocator('.preview-frame').locator('h1')).toHaveText('Alpha')
      // preview-only by default: opening find switches to split so the hits are on screen
      await expect(editorPage.locator('.pane-source')).toBeHidden()
      await editorPage
        .locator('.ribbon')
        .getByRole('button', { name: /Find and Replace/ })
        .click()
      const panel = editorPage.locator('.find-panel')
      await expect(panel).toBeVisible()
      await expect(editorPage.locator('.pane-source')).toBeVisible()

      await editorPage.keyboard.type('alpha')
      await expect(panel.locator('.find-count')).toHaveText('1/3')
      await expect(editorPage.locator('.source-editor .search-hit')).toHaveCount(3)
      await editorPage.keyboard.press('Enter')
      await expect(panel.locator('.find-count')).toHaveText('2/3')

      await panel.locator('.find-input').nth(1).fill('omega')
      await panel.locator('.find-action', { hasText: /Replace All/ }).click()
      await expect(panel.locator('.find-count')).toHaveText(/No results/)
      await expect(editorPage.locator('.status-save')).toHaveText(/Unsaved/)
      await expect(editorPage.frameLocator('.preview-frame').locator('h1')).toHaveText('omega')
      await panel.locator('.find-close').click()
      await expect(panel).toHaveCount(0)
      // the shortcut reopens it (CodeMirror's own Mod-F binding is disabled)
      await editorPage.keyboard.press('ControlOrMeta+f')
      await expect(panel).toBeVisible()
      // and pressing it again while open refocuses the query field
      await editorPage.locator('.cm-content').click()
      await editorPage.keyboard.press('ControlOrMeta+f')
      await expect(panel.locator('.find-input').first()).toBeFocused()
      await editorPage.keyboard.press('Escape')
      await expect(panel).toHaveCount(0)

      await editorPage.keyboard.press('ControlOrMeta+s')
      await expect(editorPage.locator('.status-save')).toHaveText(/Saved/)
      expect(await readFile(htmlPath, 'utf8')).toBe(source.replaceAll(/alpha/gi, 'omega'))
    } finally {
      await closeAndSaveVideo(launched, 'html-find-replace')
    }
  })
})
