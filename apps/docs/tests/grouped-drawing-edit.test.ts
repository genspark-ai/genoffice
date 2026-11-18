import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx, type TextboxDisplay } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  GROUPED_DRAWING,
  GROUPED_DRAWING_WITH_FALLBACK,
} from '../../../packages/docx-engine/tests/helpers/grouped-drawing'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { textboxBoxStyle } from '../src/renderer/editor/protected-render'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

async function open(bodyXml = GROUPED_DRAWING) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.push(editor)
  const boxes = () => editor.state.doc.firstChild!.attrs.textboxes as TextboxDisplay[]
  const update = (index: number, patch: Partial<TextboxDisplay>) => {
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...editor.state.doc.firstChild!.attrs,
        textboxes: boxes().map((b, i) => (i === index ? { ...b, ...patch } : b)),
      }),
    )
  }
  const save = async () => {
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    return parseDocx(await saveDocx(parsed, plan.saveBlocks))
  }
  return { editor, parsed, boxes, update, save }
}

describe('manual grouped drawing edits', () => {
  it('resizes and moves the second child without changing its siblings or group extent', async () => {
    const { boxes, update, save, parsed } = await open()
    const before = boxes()
    update(1, {
      widthPx: 240,
      heightPx: 60,
      offsetXEmu: before[1].offsetXEmu! + 95250,
      offsetYEmu: before[1].offsetYEmu! + 190500,
    })
    const saved = await save()
    const after = saved.blocks[0].textboxes!
    expect(after[1]).toMatchObject({
      widthPx: 240,
      heightPx: 60,
      offsetXEmu: 3143250,
      offsetYEmu: 381000,
    })
    expect(after[0]).toEqual(parsed.blocks[0].textboxes![0])
    expect(after.slice(2)).toEqual(parsed.blocks[0].textboxes!.slice(2))
    expect(saved.blocks[0].originalXml).toContain('<wp:extent cx="9525000" cy="4762500"/>')
  })

  it('saves an attributed child shape style by identity', async () => {
    const { update, save } = await open()
    update(1, { fill: 'FF0000', rotDeg: 90, borderWidthPx: 3 })
    const saved = await save()
    expect(saved.blocks[0].textboxes![1]).toMatchObject({
      fill: 'FF0000',
      rotDeg: 90,
      borderWidthPx: 3,
    })
    expect(saved.blocks[0].textboxes![0].fill).toBeUndefined()
  })

  it('keeps the horizontal and vertical line geometry separate from their click area', async () => {
    const { boxes } = await open()
    const horizontal = textboxBoxStyle(boxes()[2])
    const vertical = textboxBoxStyle(boxes()[3])
    expect(horizontal).toContain('height:8px')
    expect(vertical).toContain('width:8px')
    const svg = (css: string) => decodeURIComponent(/data:image\/svg\+xml,([^"]+)/.exec(css)![1])
    expect(svg(horizontal)).toContain('marker-end=')
    expect(svg(vertical)).toContain('stroke-dasharray=')
  })

  it('exposes resize handles for the actual grouped children', async () => {
    const { editor } = await open()
    expect(editor.view.dom.querySelectorAll('.box-rotator')).toHaveLength(5)
    expect(editor.view.dom.querySelectorAll('.box-resize-handle')).toHaveLength(5)
  })

  it('resizes the pressed child by mouse, saves it, and supports undo', async () => {
    const { editor, boxes, save } = await open()
    const handle = editor.view.dom.querySelectorAll('.box-resize-handle')[1]
    handle.dispatchEvent(
      new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: 200, clientY: 40 }),
    )
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: 240, clientY: 60 }))
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: 240, clientY: 60 }))
    expect(boxes()[1]).toMatchObject({ widthPx: 240, heightPx: 60 })
    expect((await save()).blocks[0].textboxes![1]).toMatchObject({ widthPx: 240, heightPx: 60 })
    editor.commands.undo()
    expect(boxes()[1]).toMatchObject({ widthPx: 200, heightPx: 40 })
  })

  it('moves the pressed child and rotates it about its own centre', async () => {
    const { editor, boxes, save } = await open()
    const el = editor.view.dom.querySelectorAll<HTMLElement>('.doc-textbox')[1]
    el.dispatchEvent(
      new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: 20, clientY: 20 }),
    )
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: 30, clientY: 40 }))
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: 30, clientY: 40 }))
    expect(boxes()[1]).toMatchObject({ offsetXEmu: 3143250, offsetYEmu: 381000 })
    el.getBoundingClientRect = () => new DOMRect(100, 100, 100, 100)
    const handle = editor.view.dom.querySelectorAll('.box-rotate-handle')[1]
    handle.dispatchEvent(
      new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: 200, clientY: 150 }),
    )
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: 150, clientY: 200 }))
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: 150, clientY: 200 }))
    expect(boxes()[1].rotDeg).toBe(90)
    expect((await save()).blocks[0].textboxes![1]).toMatchObject({
      offsetXEmu: 3143250,
      offsetYEmu: 381000,
      rotDeg: 90,
    })
  })

  it('keeps the VML compatibility child geometry and paint synchronized', async () => {
    const { boxes, update, save } = await open(GROUPED_DRAWING_WITH_FALLBACK)
    update(1, {
      widthPx: 240,
      heightPx: 60,
      offsetXEmu: boxes()[1].offsetXEmu! + 95250,
      fill: 'FF0000',
      rotDeg: 90,
    })
    const xml = (await save()).blocks[0].originalXml!
    expect(xml).toContain('left:4300;top:2200;width:2400;height:600;rotation:90')
    expect(xml).toContain('fillcolor="#FF0000"')
    update(2, { widthPx: 120 })
    expect((await save()).blocks[0].originalXml).toContain('from="3200,2400" to="4400,2400"')
  })

  it('edits text inside a child without rebuilding its group', async () => {
    const { editor, save } = await open(GROUPED_DRAWING_WITH_FALLBACK)
    const el = editor.view.dom.querySelectorAll('.doc-textbox')[1]
    el.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))
    const { getActiveSubEditor } = await import('../src/renderer/editor/active-editor')
    const content = el.querySelector('.doc-textbox-editor')!
    content.dispatchEvent(new FocusEvent('focus'))
    getActiveSubEditor()!.commands.setContent('<p>Updated destination</p>')
    window.dispatchEvent(new Event('ai-docs-commit-tables'))
    const xml = (await save()).blocks[0].originalXml!
    expect(xml.match(/Updated destination/g)).toHaveLength(2)
    expect(xml.match(/<w:t>Source<\/w:t>/g)).toHaveLength(2)
    expect(xml).toContain('<a:chOff x="1000" y="2000"/>')
  })

  it('cancels a gesture when the document changes or the editor is destroyed', async () => {
    const { editor, boxes, update } = await open()
    const handle = editor.view.dom.querySelectorAll('.box-resize-handle')[1]
    handle.dispatchEvent(
      new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: 200, clientY: 40 }),
    )
    update(1, { widthPx: 260 })
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: 240, clientY: 60 }))
    expect(boxes()[1].widthPx).toBe(260)
    const removed = vi.spyOn(window, 'removeEventListener')
    handle.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))
    editor.destroy()
    expect(removed.mock.calls.some(([type]) => type === 'mousemove')).toBe(true)
    expect(removed.mock.calls.some(([type]) => type === 'mouseup')).toBe(true)
    removed.mockRestore()
  })

  it('keeps reversed VML endpoint directions when resizing a line', async () => {
    const source = GROUPED_DRAWING_WITH_FALLBACK.replace(
      'from="3200,2400" to="4200,2400"',
      'from="4200,2400" to="3200,2400"',
    )
    const { update, save } = await open(source)
    update(2, { widthPx: 120 })
    expect((await save()).blocks[0].originalXml).toContain('from="4200,2400" to="3000,2400"')
  })

  it('provides enough line viewport for a thick arrow without changing its geometry', async () => {
    const { boxes } = await open()
    const box = { ...boxes()[2], borderWidthPx: 4 }
    const css = textboxBoxStyle(box)
    const svg = decodeURIComponent(/data:image\/svg\+xml,([^"]+)/.exec(css)![1])
    const height = Number(/<svg[^>]*height="([\d.]+)"/.exec(svg)![1])
    expect(height).toBeGreaterThanOrEqual(12)
    expect(box.heightPx).toBe(0)
  })

  it('honors a grouped fixed-height text box while its text is edited', async () => {
    const { editor, boxes, save } = await open()
    const el = editor.view.dom.querySelectorAll<HTMLElement>('.doc-textbox')[1]
    Object.defineProperty(el, 'scrollHeight', { configurable: true, get: () => 180 })
    el.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))
    el.querySelector('.doc-textbox-editor')!.dispatchEvent(new FocusEvent('focus'))
    const { getActiveSubEditor } = await import('../src/renderer/editor/active-editor')
    getActiveSubEditor()!.commands.setContent('<p>Longer destination text</p>')
    await new Promise((resolve) => requestAnimationFrame(resolve))
    window.dispatchEvent(new Event('ai-docs-commit-tables'))
    expect(boxes()[1].heightPx).toBe(40)
    expect((await save()).blocks[0].textboxes![1].heightPx).toBe(40)
  })
})
