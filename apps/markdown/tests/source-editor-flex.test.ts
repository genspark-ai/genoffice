/**
 * The .txt/.json view is a flex row, so its CodeMirror host has to be the flex
 * item. jsdom computes no layout, but element.matches() is real selector
 * matching, so "does the rule the stylesheet actually ships still select the
 * node that carries the layout" is answerable here — and it is the whole bug.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it } from 'vitest'
import { PlainTextEditor } from '../src/renderer/source/PlainTextEditor'

const styles = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

/** every rule that lays a `.source-editor` descendant out as a flex item */
function layoutSelectors(): string[] {
  return [...styles.matchAll(/([^{}]+)\{([^{}]*)\}/g)]
    .filter(([, , body]) => /\bflex\s*:\s*1\b/.test(body))
    .map(([, sel]) => sel.trim())
    .filter((sel) => sel.includes('.source-editor'))
}

function render(): { host: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() =>
    root.render(
      createElement(
        'div',
        { className: 'source-editor' },
        createElement(PlainTextEditor, {
          initialText: '{"key": "value"}',
          mode: 'json',
          onChange: () => {},
          spellcheck: false,
        }),
      ),
    ),
  )
  const cm = container.querySelector('.cm-editor')
  expect(cm, 'CodeMirror mounted').not.toBeNull()
  // the node the layout rule has to reach: .cm-editor's own parent
  const host = cm!.parentElement as HTMLElement
  return { host, unmount: () => void act(() => root.unmount()) }
}

describe('the .txt/.json source view fills its pane', () => {
  it('the CodeMirror host is the flex item the stylesheet actually targets', () => {
    const { host, unmount } = render()
    const matched = layoutSelectors().filter((sel) => host.matches(sel))
    expect(
      matched,
      `no flex:1 rule matches the CodeMirror host <${host.tagName.toLowerCase()} class="${host.className}">; ` +
        `the stylesheet only has ${JSON.stringify(layoutSelectors())}`,
    ).not.toHaveLength(0)
    unmount()
  })

  it('the host is a direct child of .source-editor, not nested under one', () => {
    const { host, unmount } = render()
    expect(host.parentElement?.className).toBe('source-editor')
    unmount()
  })
})
