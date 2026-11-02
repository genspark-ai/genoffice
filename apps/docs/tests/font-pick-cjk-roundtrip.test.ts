import { afterEach, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { ribbonProps } from './helpers/ribbon-props'
import { executeOps } from '../src/renderer/ai/ops'

const families = [
  'MiSans',
  'LXGW WenKai GB Screen',
  'AR PL UMing CN',
  'AR PL UKai TW MBE',
  'Droid Sans Fallback',
  'Noto Sans Mono CJK SC',
]
const text = '\u4e2d\u6587\u6d4b\u8bd5 English text 123'
const untouched = '<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Untouched</w:t></w:r></w:p>'

afterEach(() => vi.unstubAllGlobals())

it.each([...families, 'Arial'])(
  'saves the script slots selected by a real ribbon pick of %s',
  async (family) => {
    vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
    vi.stubGlobal(
      'ResizeObserver',
      class {
        observe() {}
        disconnect() {}
        unobserve() {}
      },
    )
    const original = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="SimSun"/>' +
        '<w:b/></w:rPr><w:t>' +
        text +
        '</w:t></w:r></w:p>' +
        untouched,
    })
    const parsed = await parseDocx(original)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editor.commands.setTextSelection({ from: 1, to: text.length + 1 })
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)
    try {
      act(() => root.render(createElement(Ribbon, ribbonProps(editor, computeFormatState(editor)))))
      const box = host.querySelector<HTMLInputElement>('input.rb-font-family')!
      const setNative = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
      act(() => {
        box.focus()
        setNative.call(box, family)
        box.dispatchEvent(new Event('input', { bubbles: true }))
        box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
      })
      const eastAsia = family === 'Arial' ? 'SimSun' : family
      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const saved = await saveDocx(parsed, plan.saveBlocks)
      const zip = await JSZip.loadAsync(saved)
      const xml = await zip.file('word/document.xml')!.async('string')
      expect(xml).toContain(`w:eastAsia="${eastAsia}"`)
      expect(xml).toContain(`w:ascii="${family}"`)
      expect(xml).toContain(`w:hAnsi="${family}"`)
      expect(xml).toContain(untouched)
      expect(editor.getAttributes('docTextStyle')).toMatchObject({
        fontAscii: family,
        font: eastAsia,
      })
      const reopened = await parseDocx(saved)
      expect(reopened.blocks[0].runs![0]).toMatchObject({
        text,
        font: eastAsia,
        fontAscii: family,
        bold: true,
      })
      const before = await JSZip.loadAsync(original)
      for (const part of ['word/styles.xml', 'word/_rels/document.xml.rels'])
        expect(await zip.file(part)!.async('uint8array')).toEqual(
          await before.file(part)!.async('uint8array'),
        )
    } finally {
      act(() => root.unmount())
      host.remove()
      editor.destroy()
    }
  },
)

it.each(families)(
  'the shared docs operations route %s to East Asian without changing Latin',
  (family) => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text,
                marks: [{ type: 'docTextStyle', attrs: { font: 'SimSun', fontAscii: 'Calibri' } }],
              },
            ],
          },
        ],
      },
    })
    try {
      expect(
        executeOps(editor, [
          { op: 'setFont', target: { nodeType: 'docParagraph' }, fontFamily: family },
        ]).ok,
      ).toBe(true)
      const mark = editor.state.doc.firstChild!.firstChild!.marks.find(
        (m) => m.type.name === 'docTextStyle',
      )!
      expect(mark.attrs).toMatchObject({ font: family, fontAscii: 'Calibri', eastAsiaFont: family })
    } finally {
      editor.destroy()
    }
  },
)
