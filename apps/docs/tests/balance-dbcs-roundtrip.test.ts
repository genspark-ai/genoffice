import { Editor } from '@tiptap/core'
import { DOMSerializer } from '@tiptap/pm/model'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { executeOps } from '../src/renderer/ai/ops'
import { charSpacingFromAttrs, setTextStyleWithScale } from '../src/renderer/components/FontDialog'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { mountHfEditor } from '../src/renderer/editor/hf-editor'
import { runLetterSpacingCss } from '../src/renderer/line-metrics'

const untouched = '<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Untouched</w:t></w:r></w:p>'

async function input(text: string, spacing: number, balanced = true) {
  return buildDocx({
    bodyXml:
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="SimSun"/>' +
      `<w:spacing w:val="${spacing}"/><w:sz w:val="28"/></w:rPr><w:t>${text}</w:t></w:r></w:p>` +
      untouched,
    extraParts: [
      {
        path: 'word/settings.xml',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
        xml:
          '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          `<w:compat>${balanced ? '<w:balanceSingleByteDoubleByteWidth/>' : ''}</w:compat></w:settings>`,
      },
    ],
  })
}

function editorFor(parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editor.storage.listNumbering.styles = parsed.styles
  editor.storage.listNumbering.docDefaults = parsed.docDefaults
  return editor
}

function displaySpacing(editor: Editor): string {
  const root = document.createElement('div')
  root.append(DOMSerializer.fromSchema(editor.schema).serializeFragment(editor.state.doc.content))
  return root.querySelector<HTMLElement>('span[style*="letter-spacing"]')!.style.letterSpacing
}

describe('balanced character spacing preserves authored formatting', () => {
  it('keeps an unedited document on the original-byte save path', async () => {
    const original = await input('测试字符间距文本', 10)
    const parsed = await parseDocx(original)
    const editor = editorFor(parsed)
    try {
      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      expect(plan.saveBlocks.every((block) => block.kind === 'original')).toBe(true)
      const zip = await JSZip.loadAsync(await saveDocx(parsed, plan.saveBlocks))
      expect(await zip.file('word/document.xml')!.async('string')).toBe(
        await (await JSZip.loadAsync(original)).file('word/document.xml')!.async('string'),
      )
      expect(runLetterSpacingCss(parsed.blocks[0].runs![0])).toBe('1pt')
    } finally {
      editor.destroy()
    }
  })

  for (const balanced of [true, false]) {
    for (const spacing of [10, -10, 0]) {
      it.each(['测试字符间距文本', '测试 ABC 中文'])(
        `keeps native spacing ${spacing} through three text edits and saves (balance=${balanced}, %s)`,
        async (text) => {
          const original = await input(text, spacing, balanced)
          let parsed = await parseDocx(original)
          const editor = editorFor(parsed)
          try {
            const beforeSpacing = displaySpacing(editor)
            for (const [find, replace] of [
              ['测试', '验证'],
              ['验证', '复测'],
              ['复测', '测试'],
            ]) {
              expect(
                executeOps(editor, [
                  {
                    op: 'findReplace',
                    target: { blockIndexes: [0] },
                    find,
                    replace,
                  },
                ]).ok,
              ).toBe(true)
              const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
              const bytes = await saveDocx(parsed, plan.saveBlocks)
              const zip = await JSZip.loadAsync(bytes)
              const xml = await zip.file('word/document.xml')!.async('string')
              expect(xml).toContain(`<w:spacing w:val="${spacing}"/>`)
              expect(xml).toContain(untouched)
              const source = await JSZip.loadAsync(original)
              for (const part of [
                'word/settings.xml',
                'word/styles.xml',
                'word/_rels/document.xml.rels',
              ])
                expect(await zip.file(part)!.async('uint8array')).toEqual(
                  await source.file(part)!.async('uint8array'),
                )
              parsed = await parseDocx(bytes)
              expect(parsed.blocks[0].runs![0].charSpacingTwips).toBe(spacing)
              editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
              expect(displaySpacing(editor)).toBe(beforeSpacing)
            }
          } finally {
            editor.destroy()
          }
        },
      )
    }
  }

  it('recomputes the display adjustment when replacement changes the script mix', async () => {
    const parsed = await parseDocx(await input('测试字符间距文本', 10))
    const editor = editorFor(parsed)
    try {
      expect(displaySpacing(editor)).toBe('1pt')
      expect(
        executeOps(editor, [
          {
            op: 'findReplace',
            target: { blockIndexes: [0] },
            find: '测试字符间距文本',
            replace: 'abcdefgh',
          },
        ]).ok,
      ).toBe(true)
      expect(displaySpacing(editor)).toBe('0.5pt')
      const bytes = await saveDocx(
        parsed,
        pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
      )
      const reopened = editorFor(await parseDocx(bytes))
      try {
        expect(displaySpacing(reopened)).toBe('0.5pt')
      } finally {
        reopened.destroy()
      }
    } finally {
      editor.destroy()
    }
  })

  it.each([20, -20, 0])(
    'saves explicitly edited native spacing %s and compensates only for display',
    async (spacing) => {
      const parsed = await parseDocx(await input('测试字符间距文本', 10))
      const editor = editorFor(parsed)
      try {
        expect(
          charSpacingFromAttrs(
            editor.state.doc.firstChild!.firstChild!.marks.find(
              (m) => m.type.name === 'docTextStyle',
            )!.attrs,
            14,
          ).spacingPt,
        ).toBe(0.5)
        const tr = editor.state.tr
        setTextStyleWithScale(tr, [{ from: 1, to: 9 }], { charSpacingTwips: spacing })
        editor.view.dispatch(tr)
        expect(parseFloat(displaySpacing(editor))).toBe(spacing / 10)
        const bytes = await saveDocx(
          parsed,
          pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
        )
        const reopened = await parseDocx(bytes)
        expect(reopened.blocks[0].runs![0].charSpacingTwips).toBe(spacing)
        const next = editorFor(reopened)
        try {
          expect(displaySpacing(next)).toBe(displaySpacing(editor))
        } finally {
          next.destroy()
        }
      } finally {
        editor.destroy()
      }
    },
  )

  it('uses the new East Asian font gate immediately after a font edit', async () => {
    const parsed = await parseDocx(await input('정보화비전', 10))
    const editor = editorFor(parsed)
    try {
      expect(displaySpacing(editor)).toBe('1pt')
      const tr = editor.state.tr
      setTextStyleWithScale(tr, [{ from: 1, to: 6 }], {
        font: 'Malgun Gothic',
        eastAsiaFont: 'Malgun Gothic',
      })
      editor.view.dispatch(tr)
      expect(displaySpacing(editor)).toBe('0.5pt')
      const saved = await saveDocx(
        parsed,
        pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
      )
      const reopened = editorFor(await parseDocx(saved))
      try {
        expect(displaySpacing(reopened)).toBe(displaySpacing(editor))
      } finally {
        reopened.destroy()
      }
    } finally {
      editor.destroy()
    }
  })

  it('applies compatibility rendering when spacing is newly added in the active document', async () => {
    const original = await input('测试字符间距文本', 0)
    const zip = await JSZip.loadAsync(original)
    zip.file(
      'word/document.xml',
      (await zip.file('word/document.xml')!.async('string')).replace('<w:spacing w:val="0"/>', ''),
    )
    const parsed = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    const editor = editorFor(parsed)
    editor.storage.balancedSpacing.enabled = true
    try {
      const tr = editor.state.tr
      setTextStyleWithScale(tr, [{ from: 1, to: 9 }], { charSpacingTwips: 20 })
      editor.view.dispatch(tr)
      expect(displaySpacing(editor)).toBe('2pt')
      const saved = await saveDocx(
        parsed,
        pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
      )
      expect((await parseDocx(saved)).blocks[0].runs![0].charSpacingTwips).toBe(20)
    } finally {
      editor.destroy()
    }
  })

  for (const operation of ['markup', 'attribute'] as const) {
    it(`updates the inherited font after a paragraph style ${operation} edit`, async () => {
      const bytes = await input('정보화비전', 10)
      const zip = await JSZip.loadAsync(bytes)
      zip.file(
        'word/document.xml',
        (await zip.file('word/document.xml')!.async('string'))
          .replace('<w:p>', '<w:p><w:pPr><w:pStyle w:val="Legacy"/></w:pPr>')
          .replace('<w:rFonts w:ascii="Calibri" w:eastAsia="SimSun"/>', ''),
      )
      zip.file(
        'word/styles.xml',
        (await zip.file('word/styles.xml')!.async('string')).replace(
          '</w:styles>',
          '<w:style w:type="paragraph" w:styleId="Legacy"><w:rPr><w:rFonts w:eastAsia="Batang"/></w:rPr></w:style>' +
            '<w:style w:type="paragraph" w:styleId="Modern"><w:rPr><w:rFonts w:eastAsia="Malgun Gothic"/></w:rPr></w:style></w:styles>',
        ),
      )
      const parsed = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
      const editor = editorFor(parsed)
      try {
        expect(displaySpacing(editor)).toBe('1pt')
        const tr =
          operation === 'markup'
            ? editor.state.tr.setNodeMarkup(0, undefined, {
                ...editor.state.doc.firstChild!.attrs,
                styleId: 'Modern',
              })
            : editor.state.tr.setNodeAttribute(0, 'styleId', 'Modern')
        editor.view.dispatch(tr)
        expect(displaySpacing(editor)).toBe('0.5pt')
      } finally {
        editor.destroy()
      }
    })
  }

  it('retains a character-style East Asian font while editing text', async () => {
    const zip = await JSZip.loadAsync(await input('정보화비전', 10))
    zip.file(
      'word/document.xml',
      (await zip.file('word/document.xml')!.async('string'))
        .replace('<w:p>', '<w:p><w:pPr><w:pStyle w:val="Legacy"/></w:pPr>')
        .replace(
          '<w:rFonts w:ascii="Calibri" w:eastAsia="SimSun"/>',
          '<w:rStyle w:val="ModernChar"/>',
        ),
    )
    zip.file(
      'word/styles.xml',
      (await zip.file('word/styles.xml')!.async('string')).replace(
        '</w:styles>',
        '<w:style w:type="paragraph" w:styleId="Legacy"><w:rPr><w:rFonts w:eastAsia="Batang"/></w:rPr></w:style>' +
          '<w:style w:type="character" w:styleId="ModernChar"><w:rPr><w:rFonts w:eastAsia="Malgun Gothic"/></w:rPr></w:style></w:styles>',
      ),
    )
    const editor = editorFor(await parseDocx(await zip.generateAsync({ type: 'uint8array' })))
    try {
      expect(displaySpacing(editor)).toBe('0.5pt')
      editor.view.dispatch(editor.state.tr.insertText('가', 1, 2))
      expect(displaySpacing(editor)).toBe('0.5pt')
    } finally {
      editor.destroy()
    }
  })

  it('inherits document compatibility when adding spacing in a header editor', async () => {
    const owner = editorFor(await parseDocx(await input('测试', 10)))
    owner.storage.balancedSpacing.enabled = true
    const handle = mountHfEditor(document.createElement('div'), {
      value: { text: '', paras: [{ runs: [{ text: '中文', eastAsiaFont: 'SimSun' }] }] },
      owner,
      pageNo: '1',
      pageTotal: '1',
      spellcheck: false,
      onCommit: () => {},
      onExit: () => {},
    })
    try {
      const tr = handle.editor.state.tr
      setTextStyleWithScale(tr, [{ from: 1, to: 3 }], { charSpacingTwips: 20 })
      handle.editor.view.dispatch(tr)
      expect(displaySpacing(handle.editor)).toBe('2pt')
    } finally {
      handle.discard()
      owner.destroy()
    }
  })
})
