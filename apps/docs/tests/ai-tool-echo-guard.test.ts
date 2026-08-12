import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  parseDocx,
} from '@genoffice/docx-engine'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { executeTool } from '../src/renderer/ai/tools'

/**
 * Guard tests for tool-protocol echo (alpha feedback r102): a model that saw
 * gateway-flattened tool results can pass them back as insert/replace html —
 * raw {"index":…} block dumps, literal </tool_response> tags. Those must be
 * rejected with a retryable error instead of landing in the document as text.
 */

const NUM_IDS = { bullet: BLANK_BULLET_NUM_ID, ordered: BLANK_ORDERED_NUM_ID }

async function createBlankEditor() {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const bytes = await buildBlankDocx()
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return editor
}

const BLOCK_DUMP =
  '[{"index":0,"type":"p","content":"Lettre de justification"},{"index":1,"type":"li","content":"Motif"}],"wordCount":237,"charCountNoSpaces":1437}'

async function run(editor: Editor, name: string, input: Record<string, unknown>) {
  return executeTool(editor, { id: 't', name, input }, NUM_IDS)
}

describe('insert/replace html echo guard', () => {
  it('rejects html containing a literal tool_response tag', async () => {
    const editor = await createBlankEditor()
    const exec = await run(editor, 'insert_content', {
      html: `<p>ok</p>${BLOCK_DUMP}</tool_response>`,
    })
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('tool_response')
    editor.destroy()
  })

  it('rejects a raw JSON block dump', async () => {
    const editor = await createBlankEditor()
    const exec = await run(editor, 'replace_blocks', {
      startBlockIndex: 0,
      endBlockIndex: 0,
      html: BLOCK_DUMP,
    })
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('JSON')
    editor.destroy()
  })

  it('rejects a block dump wrapped in a markdown code fence or prose', async () => {
    const editor = await createBlankEditor()
    for (const html of [
      '```\n' + BLOCK_DUMP + '\n```',
      '```json\n' + BLOCK_DUMP + '\n```',
      'Here is the result:\n```json\n' + BLOCK_DUMP + '\n```',
    ]) {
      const exec = await run(editor, 'insert_content', { html })
      expect(exec.isError).toBe(true)
      expect(exec.output).toContain('JSON')
    }
    editor.destroy()
  })

  it('rejects parseable raw JSON without the block-dump shape', async () => {
    const editor = await createBlankEditor()
    const exec = await run(editor, 'insert_content', {
      html: '{"content":"hello","note":"not html"}',
    })
    expect(exec.isError).toBe(true)
    editor.destroy()
  })

  it('accepts normal HTML, fenced HTML, and brace-led plain text', async () => {
    const editor = await createBlankEditor()
    const ok = await run(editor, 'insert_content', { html: '<p>Bonjour</p>' })
    expect(ok.isError).toBeFalsy()
    const fenced = await run(editor, 'insert_content', {
      html: '```html\n<p>Fenced but real content</p>\n```',
    })
    expect(fenced.isError).toBeFalsy()
    const braces = await run(editor, 'insert_content', {
      html: '<p>{placeholder} braces in prose are fine</p>',
    })
    expect(braces.isError).toBeFalsy()
    editor.destroy()
  })

  it('does not reject valid HTML containing an embedded code fence', async () => {
    const editor = await createBlankEditor()
    const exec = await run(editor, 'insert_content', {
      html: '<p>Config example:</p><pre>```json\n{"port": 8080}\n```</pre><p>End.</p>',
    })
    expect(exec.isError).toBeFalsy()
    editor.destroy()
  })
})
