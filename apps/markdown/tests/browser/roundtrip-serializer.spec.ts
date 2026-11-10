import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { expect, test } from '@playwright/test'
import type { Editor } from '@tiptap/core'
import { openSource, rebaseSource, source } from './helpers'

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
    await expect(page.locator('.doc-editor img:not(.ProseMirror-separator)')).toHaveCount(1)
    await page.evaluate(() => window.dispatchEvent(new Event('test:save')))
    await expect(page.locator('body')).toHaveAttribute('data-saved', rebaseSource)
    await page.evaluate(() => window.dispatchEvent(new Event('test:save')))
    await expect(page.locator('body')).toHaveAttribute(
      'data-saved',
      rebaseSource.replace('assets/old.png', 'assets/new.png'),
    )
  })
}

for (const enabled of [false, true]) {
  for (const [name, markdown, blockType] of [
    ['heading', '# abcdef\n', 'heading'],
    ['paragraph', 'abcdef\n', 'paragraph'],
    ['bullet list', '- abcdef\n', 'paragraph'],
    ['ordered list', '1. abcdef\n', 'paragraph'],
    ['task list', '- [ ] abcdef\n', 'paragraph'],
  ] as const) {
    for (const offset of [3, 6]) {
      test(`${name} Enter at ${offset}, flag ${enabled}: subsequent typing stays in the split block`, async ({
        page,
      }) => {
        await openSource(page, enabled, false, markdown)
        await page.locator('.doc-editor').evaluate(
          (node, { offset, blockType }) => {
            const editor = (node as HTMLElement & { editor: Editor }).editor
            let target: number | undefined
            editor.state.doc.descendants((child, pos) => {
              if (child.type.name === blockType && child.textContent === 'abcdef')
                target = pos + 1 + offset
            })
            if (target === undefined) throw new Error('fixture text block missing')
            editor.commands.setTextSelection(target)
            editor.view.focus()
          },
          { offset, blockType },
        )
        await page.keyboard.type('X')
        await page.keyboard.press('Enter')
        await page.keyboard.type('Y')
        const blocks = await page.locator('.doc-editor').evaluate((node) => {
          const editor = (node as HTMLElement & { editor: Editor }).editor
          editor.state.doc.check()
          const blocks: string[] = []
          editor.state.doc.descendants((child) => {
            if (child.isTextblock) blocks.push(child.textContent)
          })
          return { text: blocks.filter(Boolean), structure: editor.state.doc.toJSON() }
        })
        expect(blocks.text).toEqual(['abcdef'.slice(0, offset) + 'X', 'Y' + 'abcdef'.slice(offset)])
        await page.evaluate(() => window.dispatchEvent(new Event('test:save')))
        await expect(page.locator('body')).toHaveAttribute('data-saved', /Y/)
        const reopened = await page.locator('.doc-editor').evaluate((node) => {
          const editor = (node as HTMLElement & { editor: Editor }).editor
          editor.commands.setContent(document.body.dataset.saved!, { contentType: 'markdown' })
          editor.state.doc.check()
          const blocks: string[] = []
          editor.state.doc.descendants((child) => {
            if (child.isTextblock) blocks.push(child.textContent)
          })
          return { text: blocks.filter(Boolean), structure: editor.state.doc.toJSON() }
        })
        // TrailingNode adds an empty editing paragraph after headings. It has no
        // Markdown representation; compare all other structure and text exactly.
        for (const { structure } of [blocks, reopened]) {
          const last = structure.content?.at(-1)
          if (last?.type === 'paragraph' && !last.content?.length) structure.content!.pop()
        }
        expect(reopened).toEqual(blocks)
      })
    }
  }

  for (const size of [127_000, 317_000]) {
    test(`records insertion-to-frame timings at ${size} bytes, flag ${enabled}`, async ({
      page,
    }) => {
      const seed = readFileSync(resolve(process.cwd(), 'skills/genoffice/SKILL.md'), 'utf8')
      let text = Buffer.from(seed.repeat(Math.ceil(size / Buffer.byteLength(seed)) + 1))
        .subarray(0, size)
        .toString('utf8')
        .replace(/\uFFFD$/, '')
      text += ' '.repeat(size - Buffer.byteLength(text))
      expect(Buffer.byteLength(text)).toBe(size)
      await openSource(page, enabled, false, text)
      const result = await page.locator('.doc-editor').evaluate(async (node) => {
        const editor = (node as HTMLElement & { editor: Editor }).editor
        editor.commands.setTextSelection(3)
        editor.view.focus()
        // Let load-time layout settle before measuring editing work.
        await new Promise<void>((resolve) =>
          requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
        )
        const initialLength = editor.state.doc.textContent.length
        const samples: number[] = []
        for (let i = 0; i < 20; i++) {
          const start = performance.now()
          editor.commands.insertContent('x')
          await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
          samples.push(performance.now() - start)
        }
        editor.state.doc.check()
        return {
          timings: samples,
          insertedCharacters: editor.state.doc.textContent.length - initialLength,
        }
      })
      expect(result.insertedCharacters).toBe(20)
      const { timings } = result
      const sorted = [...timings].sort((a, b) => a - b)
      const p95 = sorted[Math.ceil(sorted.length * 0.95) - 1]
      console.log(
        JSON.stringify({
          bytes: size,
          flag: enabled,
          samples: timings.length,
          p95,
          max: Math.max(...timings),
        }),
      )
    })
  }
}
