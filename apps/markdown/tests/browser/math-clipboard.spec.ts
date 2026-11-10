import { expect, test, type Page } from '@playwright/test'
import type { Editor } from '@tiptap/core'
import { openSource } from './math-clipboard-helpers'

// Format fixtures, not captures from a logged-in AI service. Both MIME types
// deliberately differ so a test cannot pass by always preferring plain text.
const richFormula = {
  text: 'fallback without formula',
  html: '<p><strong>Energy</strong> <span class="katex"><span class="katex-mathml"><math><semantics><mi>E</mi><annotation encoding="application/x-tex">E=mc^2</annotation></semantics></math></span><span class="katex-html">duplicate</span></span> <a href="https://example.com">source</a></p>',
}

async function paste(page: Page, text: string, html?: string, plain = false) {
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.locator('.doc-editor').click()
  await page.evaluate(
    async ({ text, html }) => {
      const data: Record<string, Blob> = { 'text/plain': new Blob([text], { type: 'text/plain' }) }
      if (html) data['text/html'] = new Blob([html], { type: 'text/html' })
      await navigator.clipboard.write([new ClipboardItem(data)])
    },
    { text, html },
  )
  if (plain) {
    // Headless Chromium does not expose the macOS menu accelerator. Dispatch
    // its editing command while keeping the real system clipboard payload.
    const session = await page.context().newCDPSession(page)
    await session.send('Input.dispatchKeyEvent', {
      type: 'keyDown',
      key: 'V',
      code: 'KeyV',
      modifiers: 8,
      commands: ['pasteAndMatchStyle'],
    })
    await session.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'V', code: 'KeyV' })
    await session.detach()
  } else {
    await page.keyboard.press('ControlOrMeta+v')
  }
}

async function formulas(page: Page) {
  return page.locator('.doc-editor').evaluate((node) => {
    const editor = (node as HTMLElement & { editor: Editor }).editor
    const result: { type: string; latex: string }[] = []
    editor.state.doc.descendants((child) => {
      if (child.type.name === 'inlineMath' || child.type.name === 'blockMath')
        result.push({ type: child.type.name, latex: child.attrs.latex })
    })
    return result
  })
}

for (const [text, type, latex] of [
  ['$E=mc^2$', 'inlineMath', 'E=mc^2'],
  [String.raw`\(E=mc^2\)`, 'inlineMath', 'E=mc^2'],
  ['$$\nx^2 + y^2\n$$', 'blockMath', 'x^2 + y^2'],
  ['\\[\nx^2 + y^2\n\\]', 'blockMath', 'x^2 + y^2'],
]) {
  test(`native clipboard: ${text}`, async ({ page }) => {
    await openSource(page, false, false, '')
    await paste(page, text)
    await expect.poll(() => formulas(page)).toEqual([{ type, latex }])
    await expect(page.locator('.doc-editor .katex')).toHaveCount(1)
    await expect(page.locator('.doc-editor .katex-error')).toHaveCount(0)
    await page.evaluate(() => window.dispatchEvent(new Event('test:save')))
    await expect(page.locator('body')).toHaveAttribute('data-saved', /\$/)
    const saved = await page.locator('body').getAttribute('data-saved')
    // New page/editor exercises file loading instead of setContent on the same view.
    const reopened = await page.context().newPage()
    await openSource(reopened, false, false, saved!)
    await expect.poll(() => formulas(reopened)).toEqual([{ type, latex }])
    await reopened.close()
  })
}

test('native rich clipboard keeps marks and recovers TeX once', async ({ page }) => {
  await openSource(page, false, false, '')
  await paste(page, richFormula.text, richFormula.html)
  await expect.poll(() => formulas(page)).toEqual([{ type: 'inlineMath', latex: 'E=mc^2' }])
  await expect(page.locator('.doc-editor strong')).toHaveText('Energy')
  await expect(page.locator('.doc-editor a')).toHaveAttribute('href', 'https://example.com/')
  await expect(page.locator('.doc-editor')).not.toContainText('duplicate')
  await expect(page.locator('.doc-editor .katex')).toHaveCount(1)
})

test('native rich clipboard converts delimited formulas', async ({ page }) => {
  await openSource(page, false, false, '')
  await paste(page, 'fallback', String.raw`<p><strong>Energy</strong> \(E=mc^2\)</p>`)
  await expect.poll(() => formulas(page)).toEqual([{ type: 'inlineMath', latex: 'E=mc^2' }])
  await expect(page.locator('.doc-editor strong')).toHaveText('Energy')
})

test('paste as plain text preserves formula source', async ({ page }) => {
  await openSource(page, false, false, '')
  const text = String.raw`\(E=mc^2\)`
  await paste(page, text, '<p>ignored rich content</p>', true)
  await expect(page.locator('.doc-editor')).toHaveText(text)
  expect(await formulas(page)).toEqual([])
})

test('code, currency and escaped brackets stay literal', async ({ page }) => {
  await openSource(page, false, false, '')
  const text = String.raw`paid $5 and $10; see \[1\] and \(note\)`
  await paste(page, text)
  await expect(page.locator('.doc-editor')).toHaveText(text)
  expect(await formulas(page)).toEqual([])
  await page.locator('.doc-editor').evaluate((node) => {
    const editor = (node as HTMLElement & { editor: Editor }).editor
    editor.commands.setContent('<pre><code></code></pre>')
  })
  await paste(page, String.raw`\(x^2\) $y_1$`)
  await expect(page.locator('.doc-editor pre')).toHaveText(String.raw`\(x^2\) $y_1$`)
  expect(await formulas(page)).toEqual([])
})

for (const html of [
  String.raw`<p>\[<br>x^2 + y^2<br>\]</p>`,
  '<p>$$<br><span>x^2 + y^2</span><br>$$</p>',
]) {
  test(`native rich clipboard recognizes display math across line breaks: ${html}`, async ({
    page,
  }) => {
    await openSource(page, false, false, '')
    await paste(
      page,
      'fallback without math',
      '<p><strong>Before</strong></p>' + html + '<p>After</p>',
    )
    await expect.poll(() => formulas(page)).toEqual([{ type: 'blockMath', latex: 'x^2 + y^2' }])
    await expect(page.locator('.doc-editor strong')).toHaveText('Before')
    await expect(page.locator('.doc-editor')).toContainText('After')
    await expect(page.locator('.doc-editor .katex')).toHaveCount(1)
    await expect(page.locator('.doc-editor .katex-error')).toHaveCount(0)
  })
}
