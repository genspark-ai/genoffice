import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
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
      await expect(editorPage.locator('.tiptap')).toBeVisible()
      await editorPage.screenshot({ path: screenshotPath('new-markdown-editor') })
    } finally {
      await closeAndSaveVideo(launched, 'new-markdown-tab')
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
      const shellPage = await waitForPageWithUrl(app, 'shell/out')
      const editorTab = shellPage.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(editorTab).toHaveCount(1)
      await expect(editorTab).toContainText('note.md')

      const editorPage = await waitForPageWithUrl(app, 'markdown/out')
      const editor = editorPage.locator('.tiptap')
      await expect(editor.locator('h1')).toHaveText('Hello')
      await expect(editor.locator('strong')).toHaveText('bold')

      await editor.click()
      await editorPage.keyboard.press('ControlOrMeta+End')
      await editorPage.keyboard.press('Enter')
      await editorPage.keyboard.type('Appended line.')
      await editorPage.keyboard.press('ControlOrMeta+s')

      const saved = await readFile(mdPath, 'utf8')
      expect(saved.startsWith('---\ntitle: Note\n---\n')).toBe(true)
      expect(saved).toContain('# Hello')
      expect(saved).toContain('**bold**')
      expect(saved).toContain('Appended line.')
      await editorPage.screenshot({ path: screenshotPath('open-markdown-saved') })
    } finally {
      await closeAndSaveVideo(launched, 'open-markdown-file')
    }
  })

  test('renders markdown headings, bold, lists, and code blocks', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-md-'))
    const mdPath = join(dir, 'formatting.md')
    await writeFile(
      mdPath,
      '# Heading 1\n\n## Heading 2\n\n**bold** and *italic*\n\n- Item 1\n- Item 2\n\n```\ncode block\n```\n',
    )

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'markdown-formatting',
      openFile: mdPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, 'markdown/out')
      const editor = editorPage.locator('.tiptap')
      await expect(editor.locator('h1')).toHaveText('Heading 1')
      await expect(editor.locator('h2')).toHaveText('Heading 2')
      await expect(editor.locator('strong')).toHaveText('bold')
      await expect(editor.locator('em')).toHaveText('italic')
      await expect(editor.locator('ul li')).toHaveCount(2)
      await expect(editor.locator('pre code')).toContainText('code block')
      await editorPage.screenshot({ path: screenshotPath('markdown-formatting') })
    } finally {
      await closeAndSaveVideo(launched, 'markdown-formatting')
    }
  })

  test('ribbon formatting buttons apply marks to selection', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'markdown-ribbon' })
    const { app, page } = launched
    try {
      const card = page.locator('.quick-card', { hasText: 'AI Markdown' })
      await card.click()

      const editorPage = await waitForPageWithUrl(app, 'markdown/out')
      const editor = editorPage.locator('.tiptap')
      await expect(editor).toBeVisible()

      await editor.click()
      await editorPage.keyboard.type('Hello World')

      // Select "World" (last 5 chars)
      await editorPage.keyboard.down('Shift')
      for (let i = 0; i < 5; i++) await editorPage.keyboard.press('ArrowLeft')
      await editorPage.keyboard.up('Shift')

      const boldBtn = editorPage.locator('.ribbon-btn', { hasText: 'B' }).first()
      await boldBtn.click()
      await expect(editor.locator('strong')).toHaveText('World')
      await editorPage.screenshot({ path: screenshotPath('markdown-ribbon-bold') })
    } finally {
      await closeAndSaveVideo(launched, 'markdown-ribbon')
    }
  })

  test('dark mode toggle switches theme', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'markdown-dark' })
    const { app, page } = launched
    try {
      const card = page.locator('.quick-card', { hasText: 'AI Markdown' })
      await card.click()

      const editorPage = await waitForPageWithUrl(app, 'markdown/out')
      await expect(editorPage.locator('.tiptap')).toBeVisible()

      const viewTab = editorPage.locator('.ribbon-tab', { hasText: 'View' })
      await viewTab.click()
      const darkBtn = editorPage.locator('.ribbon-btn', { hasText: 'Dark Mode' })
      await darkBtn.click()

      const theme = await editorPage.evaluate(() =>
        document.documentElement.getAttribute('data-theme'),
      )
      expect(theme).toBe('dark')
      await editorPage.screenshot({ path: screenshotPath('markdown-dark-mode') })
    } finally {
      await closeAndSaveVideo(launched, 'markdown-dark')
    }
  })
})
