import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

const css = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
  'utf8',
)

function runStyle(sizeHalfPoints: number): string {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: null },
          content: [
            {
              type: 'text',
              text: 'x',
              marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints } }],
            },
          ],
        },
      ],
    },
  })
  const style = editor.view.dom.querySelector('p span')!.getAttribute('style') ?? ''
  editor.destroy()
  return style
}

describe('runs smaller than the strut never grow an exact / atLeast line', () => {
  it('gives fixed-rule runs an em-based natural line box, capped at an exact rule', () => {
    expect(css).toMatch(
      /\.doc-lh-fixed span \{\s*line-height: min\(\s*calc\(var\(--doc-line-factor, 1\.2\) \* 1em\),\s*var\(--doc-lh-cap, 9999px\)\s*\);/,
    )
  })

  it('an exact paragraph publishes its rule as the cap, atLeast does not', () => {
    const paraStyle = (attrs: Record<string, unknown>): string => {
      const editor = new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content: {
          type: 'doc',
          content: [
            {
              type: 'docParagraph',
              attrs: { docxIndex: null, ...attrs },
              content: [{ type: 'text', text: 'x' }],
            },
          ],
        },
      })
      const style = editor.view.dom.querySelector('p')!.getAttribute('style') ?? ''
      editor.destroy()
      return style
    }
    expect(paraStyle({ lineRule: 'exact', lineRawTwips: 240 })).toContain('--doc-lh-cap: 12.0pt')
    expect(paraStyle({ lineRule: 'atLeast', lineRawTwips: 240 })).not.toContain('--doc-lh-cap')
    expect(paraStyle({ lineRule: 'auto', lineRawTwips: 360 })).not.toContain('--doc-lh-cap')
  })

  it('collapses the line box of a spacer run of one point or less', () => {
    expect(runStyle(1)).toContain('line-height: 0')
    expect(runStyle(2)).toContain('line-height: 0')
    expect(runStyle(21)).not.toContain('line-height')
  })
})
