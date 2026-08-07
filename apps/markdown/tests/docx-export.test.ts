import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { exportDocxBytes, mapDocToSaveBlocks } from '../src/renderer/export/docxExport'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom teardown
// ("document is not defined" unhandled error) — destroy every editor we create.
const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

function createEditor(md: string): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  editor.commands.setContent(md, { contentType: 'markdown' })
  editors.push(editor)
  return editor
}

const noImages = () => Promise.resolve(null)

/** md → .docx bytes → parseDocx: the exported file must be a valid Word document */
async function exportAndParse(md: string) {
  const editor = createEditor(md)
  const bytes = await exportDocxBytes(editor.getJSON(), noImages)
  expect(bytes.length).toBeGreaterThan(1000)
  return parseDocx(bytes)
}

describe('docx export', () => {
  it('headings, paragraphs and inline marks survive into the docx', async () => {
    const parsed = await exportAndParse('# Title\n\nSome **bold** and *italic* and `code` text.')
    const heading = parsed.blocks.find((b) => b.type === 'heading')
    expect(heading?.level).toBe(1)
    expect(heading?.runs?.map((r) => r.text).join('')).toBe('Title')
    const para = parsed.blocks.find((b) => b.type === 'paragraph')
    const runs = para?.runs ?? []
    expect(runs.find((r) => r.bold)?.text).toBe('bold')
    expect(runs.find((r) => r.italic)?.text).toBe('italic')
    expect(runs.map((r) => r.text).join('')).toContain('code')
  })

  it('bullet and ordered lists become native Word list items', async () => {
    const parsed = await exportAndParse('- one\n- two\n\n1. first\n2. second')
    const items = parsed.blocks.filter((b) => b.type === 'listItem')
    expect(items.length).toBe(4)
    expect(items.filter((b) => b.list?.kind === 'bullet').length).toBe(2)
    expect(items.filter((b) => b.list?.kind === 'ordered').length).toBe(2)
  })

  it('nested list levels carry ilvl', async () => {
    const parsed = await exportAndParse('- top\n  - nested')
    const items = parsed.blocks.filter((b) => b.type === 'listItem')
    expect(items[0]?.list?.ilvl).toBe(0)
    expect(items[1]?.list?.ilvl).toBe(1)
  })

  it('tables become native Word tables with a bold header row', async () => {
    const parsed = await exportAndParse('| Name | Value |\n| --- | --- |\n| a | 1 |')
    const table = parsed.blocks.find((b) => b.type === 'table')
    expect(table?.table?.rows.length).toBe(2)
    expect(table?.table?.rows[0]?.[0]?.paras.join('')).toBe('Name')
    expect(table?.table?.rows[1]?.[1]?.paras.join('')).toBe('1')
  })

  it('task lists render checkbox glyphs', async () => {
    const parsed = await exportAndParse('- [x] done\n- [ ] open')
    const texts = parsed.blocks.map((b) => (b.runs ?? []).map((r) => r.text).join(''))
    expect(texts.some((t) => t.startsWith('☑'))).toBe(true)
    expect(texts.some((t) => t.startsWith('☐'))).toBe(true)
  })

  it('code blocks keep their text in a shaded monospace paragraph', async () => {
    const parsed = await exportAndParse('```js\nconst a = 1\nconst b = 2\n```')
    const code = parsed.blocks.find((b) =>
      (b.runs ?? []).some((r) => r.text.includes('const a = 1')),
    )
    expect(code).toBeDefined()
    expect(code?.runs?.[0]?.font).toBe('Consolas')
    expect(code?.runs?.[0]?.text).toContain('const b = 2')
  })

  it('callout and toggle degrade to styled paragraphs, keeping all text', async () => {
    const md =
      ':::callout {type="warning"}\nBe careful.\n:::\n\n:::toggle {summary="More"}\nHidden.\n:::'
    const parsed = await exportAndParse(md)
    const texts = parsed.blocks.map((b) => (b.runs ?? []).map((r) => r.text).join(''))
    expect(texts).toContain('Be careful.')
    expect(texts).toContain('More')
    expect(texts).toContain('Hidden.')
    const summary = parsed.blocks.find((b) => (b.runs ?? [])[0]?.text === 'More')
    expect(summary?.runs?.[0]?.bold).toBe(true)
  })

  it('links survive as hyperlink runs', async () => {
    const parsed = await exportAndParse('Visit [Genspark](https://genspark.ai) now.')
    const para = parsed.blocks.find((b) => b.type === 'paragraph')
    const link = para?.runs?.find((r) => r.link)
    expect(link?.text).toBe('Genspark')
    expect(link?.link?.href).toBe('https://genspark.ai')
  })

  it('unresolvable images fall back to alt text', async () => {
    const parsed = await exportAndParse('![architecture diagram](assets/missing.png)')
    const texts = parsed.blocks.map((b) => (b.runs ?? []).map((r) => r.text).join(''))
    expect(texts.some((t) => t.includes('architecture diagram'))).toBe(true)
  })

  it('separate ordered lists restart numbering', async () => {
    const editor = createEditor('1. a\n\ntext between\n\n1. b')
    const mapping = await mapDocToSaveBlocks(editor.getJSON(), noImages)
    expect(mapping.options.numbering?.restartNums?.length).toBe(2)
    const numIds = mapping.options.numbering!.restartNums!.map((r) => r.numId)
    expect(new Set(numIds).size).toBe(2)
  })
})

describe('inline mark coverage in docx export', () => {
  it('underline and highlight survive as run properties', async () => {
    const editor = createEditor('plain')
    editor.commands.setContent('some ==marked== and <u>underlined</u> text', {
      contentType: 'markdown',
    })
    const bytes = await exportDocxBytes(editor.getJSON(), noImages)
    const parsed = await parseDocx(bytes)
    const runs = parsed.blocks.flatMap((b) => b.runs ?? [])
    expect(runs.find((r) => r.highlight)?.text).toBe('marked')
    expect(runs.find((r) => r.underline)?.text).toBe('underlined')
  })
})
