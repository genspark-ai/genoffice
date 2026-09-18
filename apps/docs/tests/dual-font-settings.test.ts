import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { StyleInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))
function editorWith(runs: Array<Record<string, unknown>>, styleId = 'Normal') {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { styleId },
          content: runs.map((attrs) => ({
            type: 'text',
            text: '中文 English 123',
            marks: [{ type: 'docTextStyle', attrs }],
          })),
        },
      ],
    },
  })
  editors.push(editor)
  editor.commands.selectAll()
  return editor
}
describe('independent font settings', () => {
  it('reports both slots instead of choosing one by selected text', () => {
    const state = computeFormatState(editorWith([{ font: 'SimSun', fontAscii: 'Times New Roman' }]))
    expect(state.fontEastAsia).toBe('SimSun')
    expect(state.fontLatin).toBe('Times New Roman')
  })
  it('reports mixed values independently', () => {
    const state = computeFormatState(
      editorWith([
        { font: 'SimSun', fontAscii: 'Arial' },
        { font: 'KaiTi', fontAscii: 'Arial' },
      ]),
    )
    expect(state.fontEastAsia).toBeNull()
    expect(state.fontLatin).toBe('Arial')
  })
  it('resolves character and paragraph styles before defaults per slot', () => {
    const styles = new Map<string, StyleInfo>([
      [
        'Normal',
        {
          styleId: 'Normal',
          name: 'Normal',
          type: 'paragraph',
          display: { font: 'SimSun', fontAscii: 'Calibri' },
        },
      ],
      [
        'Emphasis',
        {
          styleId: 'Emphasis',
          name: 'Emphasis',
          type: 'character',
          display: { fontAscii: 'Times New Roman' },
        },
      ],
    ])
    const state = computeFormatState(editorWith([{ styleId: 'Emphasis' }]), styles, {
      asciiFont: 'Arial',
      eastAsiaFont: 'KaiTi',
    })
    expect(state.fontEastAsia).toBe('SimSun')
    expect(state.fontLatin).toBe('Times New Roman')
  })
})

import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { vi } from 'vitest'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { ribbonProps } from './helpers/ribbon-props'
it('offers explicit font slots and applies an unclassified family only to the chosen slot', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  vi.stubGlobal(
    'ResizeObserver',
    class {
      observe() {}
      disconnect() {}
      unobserve() {}
    },
  )
  const editor = editorWith([{ font: 'SimSun', fontAscii: 'Times New Roman', sizeHalfPoints: 28 }])
  const host = document.createElement('div')
  document.body.append(host)
  const root = createRoot(host)
  const settings = vi.fn(async () => {})
  try {
    act(() =>
      root.render(
        createElement(Ribbon, {
          ...ribbonProps(editor, computeFormatState(editor)),
          onFontSettings: settings,
        }),
      ),
    )
    const input = host.querySelector<HTMLInputElement>('input[data-font-slot="font"]')
    expect(input).not.toBeNull()
    expect(host.querySelector<HTMLInputElement>('input[data-font-slot="fontAscii"]')?.value).toBe(
      'Times New Roman',
    )
    act(() => {
      input!.focus()
      input!.value = 'Custom CJK Face'
      input!.blur()
    })
    expect(editor.getAttributes('docTextStyle')).toMatchObject({
      font: 'Custom CJK Face',
      fontAscii: 'Times New Roman',
      sizeHalfPoints: 28,
    })
    const scope = host.querySelector('select')!
    act(() => {
      scope.value = 'defaults'
      scope.dispatchEvent(new Event('change', { bubbles: true }))
    })
    await act(async () => {
      const latin = host.querySelector<HTMLInputElement>('input[data-font-slot="fontAscii"]')!
      latin.focus()
      latin.value = 'Georgia'
      latin.blur()
    })
    expect(settings).toHaveBeenCalledWith('defaults', { font: 'Georgia' })
    expect(editor.getAttributes('docTextStyle').fontAscii).toBe('Times New Roman')
  } finally {
    act(() => root.unmount())
    host.remove()
    vi.unstubAllGlobals()
  }
})
it('routes CJK glyphs to the East Asian slot even when the Latin face also covers CJK', () => {
  const editor = editorWith([{ font: 'SimSun', fontAscii: 'Microsoft YaHei' }])
  const cjk = editor.view.dom.querySelector<HTMLElement>('.doc-east-asian-font')
  expect(cjk?.textContent).toBe('中文')
  expect(cjk?.style.fontFamily).toContain('--doc-east-asian-font')
  expect(cjk?.parentElement?.style.getPropertyValue('--doc-east-asian-font')).toContain('SimSun')
  expect(editor.getJSON().content?.[0].content).toHaveLength(1)
})

import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
it('keeps bilingual fonts and bold through the real editor save plan and reopen', async () => {
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml: '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>中文测试 English text 123</w:t></w:r></w:p>',
    }),
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.push(editor)
  editor.commands.selectAll()
  editor.commands.setMark('docTextStyle', {
    font: 'SimSun',
    eastAsiaFont: 'SimSun',
    eaSlotEmpty: false,
  })
  editor.commands.setMark('docTextStyle', { fontAscii: 'Times New Roman' })
  editor.commands.setTextSelection(editor.state.doc.content.size - 1)
  editor.commands.insertContent(' 新增 456')
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
  expect(reopened.blocks[0].runs!.map((r) => r.text).join('')).toBe(
    '中文测试 English text 123 新增 456',
  )
  for (const run of reopened.blocks[0].runs!)
    expect(run).toMatchObject({ font: 'SimSun', fontAscii: 'Times New Roman', bold: true })
  expect(editor.commands.undo()).toBe(true)
})
it('a Latin-only edit keeps East Asian inheritance when import derived font from ascii', () => {
  const editor = editorWith([
    {
      font: 'Arial',
      fontAscii: 'Arial',
      rawRPr: '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>',
    },
  ])
  editor.commands.setMark('docTextStyle', { fontAscii: 'Times New Roman' })
  expect(computeFormatState(editor, undefined, { eastAsiaFont: 'SimSun' }).fontEastAsia).toBe(
    'SimSun',
  )
  const run = editor.view.dom.querySelector<HTMLElement>('[data-doc-style]')!
  expect(run.style.getPropertyValue('--doc-east-asian-font')).toBe('')
})

import { scriptFontHtml } from '../src/renderer/editor/script-fonts'
it('uses the same script font routing in the read-only split pane', () => {
  const editor = editorWith([
    { font: 'SimSun', eastAsiaFont: 'SimSun', fontAscii: 'Microsoft YaHei' },
  ])
  const template = document.createElement('template')
  template.innerHTML = scriptFontHtml(editor.getHTML())
  expect(template.content.querySelector('.doc-east-asian-font')?.textContent).toBe('中文')
  expect(template.content.textContent).toBe('中文 English 123')
  expect(
    template.content
      .querySelector<HTMLElement>('[data-doc-style]')
      ?.style.getPropertyValue('--doc-east-asian-font'),
  ).toContain('SimSun')
})
