import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

test.describe('markdown editor', () => {
  test('AI Markdown quick card opens a markdown editor tab', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'new-markdown-tab' })
    const { app, page } = launched
    try {
      const card = page.locator('.quick-card', { hasText: 'AI Markdown' })
      await expect(card).toHaveCount(1)
      await card.click()

      const editorTab = page.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(editorTab).toHaveCount(1)
      await expect(editorTab).toHaveClass(/active/)

      const editorPage = await waitForPageWithUrl(app, 'markdown/out')
      await expect(editorPage.locator('.doc-editor')).toBeVisible()
      await editorPage.screenshot({ path: screenshotPath('new-markdown-editor') })
    } finally {
      await closeAndSaveVideo(launched, 'new-markdown-tab')
    }
  })

  test('legacy fenced divs degrade on open; slash menu inserts a GFM task list', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-md-'))
    const mdPath = join(dir, 'legacy.md')
    await writeFile(mdPath, '# Doc\n\n:::callout {type="info"}\nBe careful.\n:::\n')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'markdown-slash-task',
      openFile: mdPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, 'markdown/out')
      const editor = editorPage.locator('.doc-editor')
      await expect(editor.locator('h1')).toHaveText('Doc')
      // the legacy callout fences are stripped on open; the body text survives
      await expect(editor).toContainText('Be careful.')
      await expect(editor).not.toContainText(':::')

      await editor.click()
      await editorPage.keyboard.press('ControlOrMeta+End')
      await editorPage.keyboard.press('Enter')
      await editorPage.keyboard.type('/')
      await expect(editorPage.locator('.slash-menu')).toBeVisible()
      await editorPage.keyboard.type('task')
      await editorPage.locator('.slash-item', { hasText: /Task list|任务列表/ }).click()
      await expect(editor.locator('ul[data-type="taskList"]')).toBeVisible()
      await editorPage.keyboard.type('Heads up!')
      await editorPage.keyboard.press('ControlOrMeta+s')
      await expect(editorPage.locator('.status-save')).toHaveText(/Saved|已保存/)

      const saved = await readFile(mdPath, 'utf8')
      expect(saved).toContain('- [ ] Heads up!')
      expect(saved).toContain('Be careful.')
      expect(saved).not.toContain(':::')
    } finally {
      await closeAndSaveVideo(launched, 'markdown-slash-task')
    }
  })

  test('opens a .md file from argv, edits and saves it back', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-md-'))
    const mdPath = join(dir, 'note.md')
    await writeFile(mdPath, '---\ntitle: Note\n---\n\n# Hello\n\nSome **bold** text.\n')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'open-markdown-file',
      openFile: mdPath,
    })
    const { app } = launched
    try {
      // with an argv document the first window is the editor view, not the shell
      // home — locate the shell renderer explicitly for tab-strip assertions
      const shellPage = await waitForPageWithUrl(app, 'shell/out')
      const editorTab = shellPage.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(editorTab).toHaveCount(1)
      await expect(editorTab).toContainText('note.md')

      const editorPage = await waitForPageWithUrl(app, 'markdown/out')
      const editor = editorPage.locator('.doc-editor')
      await expect(editor.locator('h1')).toHaveText('Hello')
      await expect(editor.locator('strong')).toHaveText('bold')

      // type at the end of the document, save with ⌘/Ctrl+S
      await editor.click()
      await editorPage.keyboard.press('ControlOrMeta+End')
      await editorPage.keyboard.press('Enter')
      await editorPage.keyboard.type('Appended line.')
      await editorPage.keyboard.press('ControlOrMeta+s')
      await expect(editorPage.locator('.status-save')).toHaveText(/Saved|已保存/)
      await editorPage.screenshot({ path: screenshotPath('open-markdown-saved') })

      const saved = await readFile(mdPath, 'utf8')
      expect(saved.startsWith('---\ntitle: Note\n---\n')).toBe(true)
      expect(saved).toContain('# Hello')
      expect(saved).toContain('**bold**')
      expect(saved).toContain('Appended line.')
      expect(saved.endsWith('\n')).toBe(true)
    } finally {
      await closeAndSaveVideo(launched, 'open-markdown-file')
    }
  })

  test('ribbon AI preset button opens the panel and sends the instruction', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-md-'))
    const mdPath = join(dir, 'summary.md')
    await writeFile(mdPath, '# Topic\n\nSome content worth summarizing.\n')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'markdown-ai-preset',
      openFile: mdPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, 'markdown/out')
      await expect(editorPage.locator('.doc-editor h1')).toHaveText('Topic')

      const summarizeBtn = editorPage.locator('.rb-big.ai-entry', {
        hasText: /AI 总结|AI Summarize/,
      })
      await expect(summarizeBtn).toBeEnabled()
      await summarizeBtn.click()

      await expect(editorPage.locator('.copilot')).toBeVisible()
      // the preset lands as a sent user message (the model reply itself needs credentials)
      await expect(editorPage.locator('.ai-msg-user')).toContainText(/总结|Summarize/)
      await editorPage.screenshot({ path: screenshotPath('markdown-ai-preset') })
    } finally {
      await closeAndSaveVideo(launched, 'markdown-ai-preset')
    }
  })

  test('ribbon bold serializes as GFM; quick-access save and undo work', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-md-'))
    const mdPath = join(dir, 'style.md')
    await writeFile(mdPath, 'Hello style\n')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'markdown-bold-qat',
      openFile: mdPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, 'markdown/out')
      const editor = editorPage.locator('.doc-editor')
      await expect(editor).toContainText('Hello style')

      await editor.click()
      await editorPage.keyboard.press('ControlOrMeta+a')
      await editorPage.getByTitle(/^(加粗|Bold)$/).click()
      await expect(editor.locator('strong')).toHaveText('Hello style')

      // quick-access row: save button writes the file, undo reverts the mark
      const qaButtons = editorPage.locator('.ribbon-tabs .qa-btn')
      await qaButtons.nth(0).click()
      await expect(editorPage.locator('.status-save')).toHaveText(/Saved|已保存/)
      const saved = await readFile(mdPath, 'utf8')
      expect(saved).toContain('**Hello style**')

      await qaButtons.nth(1).click()
      await expect(editor.locator('strong')).toHaveCount(0)

      // save again so the window closes without a dirty-document prompt
      await qaButtons.nth(0).click()
      await expect.poll(() => readFile(mdPath, 'utf8')).not.toContain('**')
    } finally {
      await closeAndSaveVideo(launched, 'markdown-bold-qat')
    }
  })

  test('renders a relative assets/ image through the md-asset protocol', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-md-'))
    const PNG_1PX =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
    await mkdir(join(dir, 'assets'))
    await writeFile(join(dir, 'assets', 'pic.png'), Buffer.from(PNG_1PX, 'base64'))
    const mdPath = join(dir, 'with-image.md')
    await writeFile(mdPath, '# Pics\n\n![a pic](assets/pic.png)\n')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'markdown-image-display',
      openFile: mdPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, 'markdown/out')
      const img = editorPage.locator('.doc-editor img[alt="a pic"]')
      await expect(img).toBeVisible()
      await expect(img).toHaveAttribute('src', /^md-asset:\/\//)
      // naturalWidth > 0 means the protocol handler actually served the bytes
      await expect
        .poll(() => img.evaluate((el) => (el as HTMLImageElement).naturalWidth))
        .toBeGreaterThan(0)
      await editorPage.screenshot({ path: screenshotPath('markdown-image-display') })
    } finally {
      await closeAndSaveVideo(launched, 'markdown-image-display')
    }
  })
})
