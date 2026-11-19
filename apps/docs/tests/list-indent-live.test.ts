/**
 * Regression: A list created on a blank document must show its
 * indent immediately - the freshly allocated numbering definition has to be
 * synced into the renderer's numbering table (overlayNumberingDef), not only
 * written to numbering.xml at save. The editing-time geometry must equal the
 * save -> reopen geometry.
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { buildBlankDocx, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocToSavePlan } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { allocateListNumId, type NumberingContext } from '../src/renderer/numbering-actions'

const liStyleOf = (editor: Editor) =>
  (editor.view.dom.querySelector('.doc-li') as HTMLElement | null)?.getAttribute('style') ?? ''

describe('new-list indent is live', () => {
  it('toggling a list on a blank doc applies the numbering indent immediately and it survives save → reopen', async () => {
    const parsed = await parseDocx(await buildBlankDocx())
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editor.storage.listNumbering.defs = parsed.numbering

    // ribbon toggleList path on a blank doc: no reusable numId in the body → allocate
    let pending = { newDefs: [], restartNums: [], styleNums: [], styleDefs: [] } as never as {
      newDefs: unknown[]
      restartNums: unknown[]
    }
    const ctx: NumberingContext = {
      editor,
      doc: { parsed } as never,
      pendingNumbering: pending as never,
      setPendingNumbering: (u) => {
        pending = u(pending as never) as never
      },
      numIdFloorRef: { current: 0 },
      setStatus: () => {},
    }
    const numId = allocateListNumId(ctx, 'bullet')
    expect(numId).toBeTruthy()
    editor.chain().setNode('docListItem', { kind: 'bullet', numId, ilvl: 0 }).run()
    editor.commands.insertContentAt(1, 'first item')

    // the allocated definition is in the renderer's table and the indent shows NOW
    const defs = editor.storage.listNumbering.defs as Map<
      string,
      { levels: Record<number, { indentLeft?: number }> }
    >
    expect(defs.get(String(numId))?.levels[0]?.indentLeft).toBe(720)
    const liveStyle = liStyleOf(editor)
    expect(liveStyle).toContain('--li-left: 36pt')
    expect(liveStyle).toContain('--li-hang: 18pt')

    // save → reopen renders the same geometry (editing view == final layout)
    const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks, { numbering: pending as never })
    const reparsed = await parseDocx(saved)
    const editor2 = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor2.storage.listNumbering.defs = reparsed.numbering
    editor2.commands.setContent(blocksToPmDoc(reparsed.blocks) as never)
    const reopenedStyle = liStyleOf(editor2)
    expect(reopenedStyle).toContain('--li-left: 36pt')
    expect(reopenedStyle).toContain('--li-hang: 18pt')

    editor.destroy()
    editor2.destroy()
  })
})
