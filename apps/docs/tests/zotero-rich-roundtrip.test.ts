import { Editor } from '@tiptap/core'
import { buildBlankDocx, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { ZoteroDocumentController } from '../src/renderer/zotero/controller'

describe('Zotero rich bibliography DOCX round-trip', () => {
  it('keeps rich runs, hanging indents, and one field across paragraphs', async () => {
    const parsed = await parseDocx(await buildBlankDocx())
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const controller = new ZoteroDocumentController(editor, { get: () => '', set: () => {} })
    const call = (command: string, args: unknown[]) =>
      controller.handle({ requestId: 'test', command, args })

    try {
      const [fieldId] = (await call('Document_insertField', [1, 'ReferenceMark', 0])) as [
        number,
        string,
        number,
      ]
      await call('Field_setCode', [1, fieldId, 'BIBL {} CSL_BIBLIOGRAPHY'])
      await call('Field_setText', [
        1,
        fieldId,
        '{\\rtf1\\pard\\li720\\fi-360 Alpha {\\i Journal} {\\super 2} {\\scaps Press}\\\r\n' +
          '\\pard\\li720\\fi-360 Beta}',
        true,
      ])

      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
      const blocks = reopened.blocks.filter((block) => !block.hidden)
      expect(blocks).toHaveLength(2)
      expect(blocks.map((block) => block.format)).toEqual([
        expect.objectContaining({ indentLeft: 720, indentFirstLine: -360 }),
        expect.objectContaining({ indentLeft: 720, indentFirstLine: -360 }),
      ])
      const runs = blocks.flatMap((block) => block.runs ?? [])
      expect(runs.find((run) => run.text === 'Journal')?.italic).toBe(true)
      expect(runs.find((run) => run.text === '2')?.vertAlign).toBe('superscript')
      expect(runs.find((run) => run.text === 'Press')?.caps).toBe('small')
      expect(new Set(runs.map((run) => run.zoteroFieldId)).size).toBe(1)
      expect(runs[0].zoteroFieldPart).toBe('begin')
      expect(runs[runs.length - 1].zoteroFieldPart).toBe('end')
    } finally {
      editor.destroy()
    }
  })
})
