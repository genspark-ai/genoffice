import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom teardown
// ("document is not defined" unhandled error). Editors here are shared per describe,
// so destroy them once at the end of the file.
const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

function createEditor(): Editor {
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
  editors.push(editor)
  return editor
}

/** parse → serialize → parse must be structurally stable */
function roundTrip(editor: Editor, md: string): { out: string; stable: boolean } {
  const manager = editor.markdown!
  const first = manager.parse(md)
  const out = manager.serialize(first)
  const second = manager.parse(out)
  return { out, stable: JSON.stringify(first) === JSON.stringify(second) }
}

describe('markdown round-trip for GFM nodes', () => {
  const editor = createEditor()

  it('task list', () => {
    const md = '- [ ] open item\n- [x] done item'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('- [ ] open item')
    expect(out).toContain('- [x] done item')
    expect(stable).toBe(true)
  })

  it('table', () => {
    const md = '| Name | Value |\n| --- | --- |\n| a | 1 |\n| b | 2 |'
    const { out, stable } = roundTrip(editor, md)
    // the serializer pads cells to align columns
    expect(out).toContain('| Name | Value |')
    expect(out).toMatch(/\| a\s+\| 1\s+\|/)
    expect(stable).toBe(true)
  })

  it('image keeps the authored relative path in markdown', () => {
    const md = '![diagram](assets/diagram.png)'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('![diagram](assets/diagram.png)')
    expect(stable).toBe(true)
  })

  it('plain constructs round-trip', () => {
    const md = '# Title\n\n> quoted\n\n```js\ncode()\n```\n\n---'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('# Title')
    expect(out).toContain('```js')
    expect(stable).toBe(true)
  })

  it('code block language survives the round-trip', () => {
    const { out, stable } = roundTrip(editor, '```python\nprint(1)\n```')
    expect(out).toContain('```python')
    expect(stable).toBe(true)
  })

  it('nested lists serialize with 4-space indents (strict-CommonMark safe)', () => {
    // 2-space indents would be below the ordered item's content column ("1. "
    // = 3 chars), so GitHub would flatten the sub-list when re-parsing the file
    const ordered = roundTrip(editor, '1. one\n    - sub\n2. two')
    expect(ordered.out).toContain('\n    - sub')
    expect(ordered.stable).toBe(true)

    const bullets = roundTrip(editor, '- a\n    - b\n        - c')
    expect(bullets.out).toContain('\n    - b')
    expect(bullets.out).toContain('\n        - c')
    expect(bullets.stable).toBe(true)

    // files saved by earlier versions used 2-space indents — still parsed as nested
    const legacy = roundTrip(editor, '- a\n  - b')
    expect(legacy.out).toContain('\n    - b')
  })
})

describe('only pure markdown syntax is ever produced', () => {
  const editor = createEditor()

  it('underline and highlight commands are gone', () => {
    // both would serialize as non-GFM syntax (`++u++` / `==mark==`)
    expect('toggleUnderline' in editor.commands).toBe(false)
    expect('toggleHighlight' in editor.commands).toBe(false)
  })

  it('alignment and line-height commands are gone', () => {
    expect('setTextAlign' in editor.commands).toBe(false)
    expect('setLineHeight' in editor.commands).toBe(false)
  })

  it('callout and toggle nodes are gone from the schema', () => {
    expect(editor.schema.nodes.callout).toBeUndefined()
    expect(editor.schema.nodes.toggle).toBeUndefined()
  })
})

describe('legacy HTML content degrades to plain markdown, keeping the text', () => {
  const editor = createEditor()

  function parseAndSerialize(md: string): string {
    const manager = editor.markdown!
    return manager.serialize(manager.parse(md))
  }

  it('a styled span drops the styling but keeps the text', () => {
    const out = parseAndSerialize('a <span style="color: #ff0000">red text</span> b')
    expect(out).not.toContain('<span')
    expect(out).toContain('red text')
  })

  it('an aligned paragraph becomes a plain paragraph with marks intact', () => {
    const out = parseAndSerialize(
      '<p style="text-align: center">centered <strong>text</strong></p>',
    )
    expect(out).not.toContain('<p')
    expect(out).toContain('centered **text**')
  })

  it('an aligned heading becomes a plain heading', () => {
    const out = parseAndSerialize('<h2 style="text-align: right">title</h2>')
    expect(out).toBe('## title')
  })

  it('a resized image goes back to pure image syntax', () => {
    const out = parseAndSerialize('<img src="assets/d.png" alt="d" width="300" align="center">')
    expect(out).toBe('![d](assets/d.png)')
  })

  it('u and mark tags drop the tag but keep the text', () => {
    const out = parseAndSerialize('a <u>underlined</u> and <mark>marked</mark> b')
    expect(out).toBe('a underlined and marked b')
  })
})
