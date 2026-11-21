import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { closeHistory, undo } from '@tiptap/pm/history'
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, saveDocx, type Block, type TextboxDisplay } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  blocksToPmDoc,
  keepGroupMembersApart,
  pmDocToSavePlan,
  type PmNode,
} from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  canGroupObjects,
  canUngroupObjects,
  ungroupSelectedObjects,
} from '../src/renderer/editor/object-group'
import {
  isFloatingObject,
  selectedObjectPositions,
  setObjectSelection,
} from '../src/renderer/editor/object-multi-select'

const EMU = 9525
const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

const child = (id: number, x: number, y: number, w: number, h: number, attrs = '', text = '') =>
  `<wps:wsp><wps:cNvPr id="${id}" name="Child ${id}"/><wps:cNvSpPr/><wps:spPr><a:xfrm${attrs}>` +
  `<a:off x="${x * EMU}" y="${y * EMU}"/><a:ext cx="${w * EMU}" cy="${h * EMU}"/></a:xfrm>` +
  `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>` +
  (text
    ? `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>`
    : '') +
  `<wps:bodyPr/></wps:wsp>`

const inner = (id: number, box: number[], kids: string) =>
  `<wpg:grpSp><wpg:cNvPr id="${id}" name="Inner ${id}"/><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm>` +
  `<a:off x="${box[0] * EMU}" y="${box[1] * EMU}"/><a:ext cx="${box[2] * EMU}" cy="${box[3] * EMU}"/>` +
  `<a:chOff x="${box[0] * EMU}" y="${box[1] * EMU}"/><a:chExt cx="${box[2] * EMU}" cy="${box[3] * EMU}"/>` +
  `</a:xfrm></wpg:grpSpPr>${kids}</wpg:grpSp>`

/** a paragraph holding a 320 x 160 group at (60, 20) in front of the text, then `after` */
const group = (attrs: string, kids: string, after = '') =>
  `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="5" ` +
  `behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>` +
  `<wp:positionH relativeFrom="column"><wp:posOffset>${60 * EMU}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${20 * EMU}</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="${320 * EMU}" cy="${160 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
  `<wp:docPr id="50" name="Group 50"/><wp:cNvGraphicFramePr/>` +
  `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">` +
  `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm${attrs}><a:off x="0" y="0"/><a:ext cx="${320 * EMU}" cy="${160 * EMU}"/>` +
  `<a:chOff x="0" y="0"/><a:chExt cx="${320 * EMU}" cy="${160 * EMU}"/></a:xfrm></wpg:grpSpPr>${kids}</wpg:wgp>` +
  `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>${after}</w:p>`

const THREE =
  child(1, 0, 0, 100, 60, '', 'One') +
  child(2, 140, 20, 80, 80) +
  child(3, 60, 110, 120, 40, '', 'Three')

async function open(bodyXml: string): Promise<{ editor: Editor; blocks: Block[] }> {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const element = document.createElement('div')
  document.body.appendChild(element)
  const editor = new Editor({
    element,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, blocks: parsed.blocks }
}

function objects(editor: Editor): number[] {
  const out: number[] = []
  editor.state.doc.forEach((node, pos) => {
    if (node.type.name === 'docProtected' && node.attrs.textboxes) out.push(pos)
  })
  return out
}

const boxesAt = (editor: Editor, pos: number): TextboxDisplay[] =>
  editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[]

const select = (editor: Editor, pos: number): void =>
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))

async function savedFile(editor: Editor, bodyXml: string): Promise<Uint8Array> {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  return saveDocx(parsed, plan.saveBlocks, {
    groups: plan.groups.length > 0 ? plan.groups : undefined,
    carried: plan.carried.length > 0 ? plan.carried : undefined,
  })
}

async function savedBody(editor: Editor, bodyXml: string): Promise<string> {
  const saved = await savedFile(editor, bodyXml)
  const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
  return /<w:body>([\s\S]*)<w:sectPr/.exec(xml)![1]
}

const offsets = (xml: string, axis: 'H' | 'V'): number[] =>
  Array.from(
    xml.matchAll(new RegExp(`<wp:position${axis}[^>]*><wp:posOffset>(-?\\d+)</wp:posOffset>`, 'g')),
  ).map((m) => Number(m[1]) / EMU)

describe('ungrouping a group that came with the file', () => {
  const body = para('Lead') + group('', THREE, '<w:r><w:t>beside</w:t></w:r>') + para('Tail')

  it('leaves one floating object per child, all selected, where the children stood', async () => {
    const { editor, blocks } = await open(body)
    const [at] = objects(editor)
    const before = boxesAt(editor, at).map((b) => [b.offsetXEmu, b.offsetYEmu, b.widthPx])
    select(editor, at)
    expect(canUngroupObjects(editor.state)).toBe(true)
    expect(ungroupSelectedObjects(editor, blocks)).toBe(true)

    const now = objects(editor)
    expect(now).toHaveLength(3)
    expect(selectedObjectPositions(editor.state).sort((a, b) => a - b)).toEqual(now)
    const nodes = now.map((pos) => editor.state.doc.nodeAt(pos)!)
    expect(nodes.every((node) => isFloatingObject(node))).toBe(true)
    expect(nodes.map((node) => node.attrs.carried)).toEqual([false, true, true])
    expect(now.map((pos) => boxesAt(editor, pos).length)).toEqual([1, 1, 1])
    expect(
      now.map((pos) => boxesAt(editor, pos)[0]).map((b) => [b.offsetXEmu, b.offsetYEmu, b.widthPx]),
    ).toEqual(before)
    expect(now.some((pos) => boxesAt(editor, pos)[0].grouped)).toBe(false)
    // the paragraph's own text stays with its first child
    expect(nodes[0].attrs.strayRuns?.[0]?.text).toBe('beside')
    expect(nodes[2].attrs.strayRuns ?? null).toBeNull()
    // nothing left to ungroup; the three can be grouped again
    expect(canUngroupObjects(editor.state)).toBe(false)
    expect(canGroupObjects(editor.state)).toBe(true)
    editor.destroy()
  })

  it('one undo puts the group back', async () => {
    const { editor, blocks } = await open(body)
    const [at] = objects(editor)
    const group = editor.state.doc.nodeAt(at)!
    select(editor, at)
    ungroupSelectedObjects(editor, blocks)
    undo(editor.state, editor.view.dispatch)
    expect(objects(editor)).toEqual([at])
    expect(editor.state.doc.nodeAt(at)!.attrs).toEqual(group.attrs)
    editor.destroy()
  })

  it('saves the children as drawings of the one paragraph the group was in', async () => {
    const { editor, blocks } = await open(body)
    select(editor, objects(editor)[0])
    ungroupSelectedObjects(editor, blocks)
    const xml = await savedBody(editor, body)
    expect(xml.match(/<w:p>/g)).toHaveLength(body.match(/<w:p>/g)!.length)
    expect(xml).not.toContain('wpg:wgp')
    expect(xml.match(/<wp:anchor /g)).toHaveLength(3)
    expect(offsets(xml, 'H')).toEqual([60, 200, 120])
    expect(offsets(xml, 'V')).toEqual([20, 40, 130])
    expect(xml).toContain('<w:t>beside</w:t>')
    expect(xml.indexOf('<w:t>One</w:t>')).toBeLessThan(xml.indexOf('<w:t>Three</w:t>'))
    // the later children follow the paragraph's own text
    expect(xml.indexOf('<w:t>beside</w:t>')).toBeLessThan(xml.indexOf('<w:t>Three</w:t>'))
    editor.destroy()
  })

  it('saves an edit made to one child after the ungroup, and only to that child', async () => {
    const { editor, blocks } = await open(body)
    select(editor, objects(editor)[0])
    ungroupSelectedObjects(editor, blocks)
    const [, second, third] = objects(editor)
    // the second child moved, the third one widened and retyped
    const moved = editor.state.doc.nodeAt(second)!
    const tr = editor.state.tr.setNodeMarkup(second, undefined, {
      ...moved.attrs,
      imageOffsetXEmu: Number(moved.attrs.imageOffsetXEmu) + 30 * EMU,
      imageOffsetYEmu: Number(moved.attrs.imageOffsetYEmu) + 10 * EMU,
      textboxes: boxesAt(editor, second).map((b) => ({
        ...b,
        offsetXEmu: b.offsetXEmu! + 30 * EMU,
        offsetYEmu: b.offsetYEmu! + 10 * EMU,
      })),
    })
    const typed = tr.doc.nodeAt(third)!
    tr.setNodeMarkup(third, undefined, {
      ...typed.attrs,
      textboxes: boxesAt(editor, third).map((b) => ({
        ...b,
        widthPx: 130,
        paras: [{ ...b.paras[0], runs: [{ ...b.paras[0].runs[0], text: 'Third' }] }],
      })),
    })
    editor.view.dispatch(tr)
    const xml = await savedBody(editor, body)
    expect(offsets(xml, 'H')).toEqual([60, 230, 120])
    expect(offsets(xml, 'V')).toEqual([20, 50, 130])
    expect(xml).toContain(`<a:ext cx="${130 * EMU}" cy="${40 * EMU}"/>`)
    expect(xml).toContain('<w:t>One</w:t>')
    expect(xml).toContain('Third')
    expect(xml).not.toContain('<w:t>Three</w:t>')
    editor.destroy()
  })

  it('a deleted child is gone from the file and the others stay', async () => {
    const { editor, blocks } = await open(body)
    select(editor, objects(editor)[0])
    ungroupSelectedObjects(editor, blocks)
    const third = objects(editor)[2]
    const node = editor.state.doc.nodeAt(third)!
    editor.view.dispatch(editor.state.tr.delete(third, third + node.nodeSize))
    const xml = await savedBody(editor, body)
    expect(xml.match(/<wp:anchor /g)).toHaveLength(2)
    expect(xml).not.toContain('<w:t>Three</w:t>')
    expect(xml).toContain('<w:t>beside</w:t>')
    expect(xml.match(/<w:p>/g)).toHaveLength(body.match(/<w:p>/g)!.length - 1)
    editor.destroy()
  })

  it("when the group's paragraph is deleted the next child becomes the paragraph", async () => {
    const { editor, blocks } = await open(body)
    select(editor, objects(editor)[0])
    ungroupSelectedObjects(editor, blocks)
    const [first] = objects(editor)
    const node = editor.state.doc.nodeAt(first)!
    editor.view.dispatch(closeHistory(editor.state.tr))
    editor.view.dispatch(editor.state.tr.delete(first, first + node.nodeSize))
    const nodes = objects(editor).map((pos) => editor.state.doc.nodeAt(pos)!)
    expect(nodes.map((n) => n.attrs.carried)).toEqual([false, true])
    const xml = await savedBody(editor, body)
    // three body paragraphs still: Lead, the two drawings in one, Tail
    expect(xml.match(/<w:p>/g)).toHaveLength(4)
    expect(xml.match(/<wp:anchor /g)).toHaveLength(2)
    expect(offsets(xml, 'H')).toEqual([200, 120])
    expect(xml).not.toContain('<w:t>One</w:t>')
    undo(editor.state, editor.view.dispatch)
    expect(objects(editor).map((pos) => editor.state.doc.nodeAt(pos)!.attrs.carried)).toEqual([
      false,
      true,
      true,
    ])
    editor.destroy()
  })

  it('stay apart when the document is read back from the file after a save', async () => {
    const { editor, blocks } = await open(body)
    select(editor, objects(editor)[0])
    ungroupSelectedObjects(editor, blocks)
    const reread = await parseDocx(await savedFile(editor, body))
    const read = blocksToPmDoc(reread.blocks) as PmNode
    // the file has them as one node of three boxes
    expect(
      read
        .content!.filter((n) => n.attrs?.textboxes)
        .map((n) => (n.attrs!.textboxes as unknown[]).length),
    ).toEqual([3])
    const kept = keepGroupMembersApart(editor.getJSON() as PmNode, read)
    expect(editor.state.doc.eq(editor.schema.nodeFromJSON(kept))).toBe(true)
    // and the next save, told from the new file, writes the same paragraph again
    editor.commands.setContent(kept as never)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, reread.blocks)
    const again = await saveDocx(reread, plan.saveBlocks, { carried: plan.carried })
    const xml = await (await JSZip.loadAsync(again)).file('word/document.xml')!.async('string')
    expect(xml.match(/<wp:anchor /g)).toHaveLength(3)
    expect(offsets(xml, 'H')).toEqual([60, 200, 120])
    expect(xml.match(/<w:t>beside<\/w:t>/g)).toHaveLength(1)
    // a document that does not line up with the file is taken as read
    const other = { ...read, content: read.content!.slice(1) }
    expect(keepGroupMembersApart(editor.getJSON() as PmNode, other)).toBe(other)
    editor.destroy()
  })

  it('grouped again, the children save as one group in the same paragraph', async () => {
    const { editor, blocks } = await open(body)
    select(editor, objects(editor)[0])
    ungroupSelectedObjects(editor, blocks)
    const [a, b] = objects(editor)
    const tr = editor.state.tr
    ;[a, b].forEach((pos, i) => {
      const node = tr.doc.nodeAt(pos)!
      tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        groupId: 'g1',
        groupBox: i === 0 ? { x: 0, y: 0, w: 100, h: 60 } : { x: 140, y: 20, w: 80, h: 80 },
      })
    })
    editor.view.dispatch(setObjectSelection(tr, [a, b]))
    const xml = await savedBody(editor, body)
    expect(xml.match(/<wpg:wgp /g)).toHaveLength(1)
    expect(xml.match(/<wp:anchor /g)).toHaveLength(2)
    expect(xml.match(/<w:p>/g)).toHaveLength(body.match(/<w:p>/g)!.length)
    expect(offsets(xml, 'H')).toEqual([60, 120])
    editor.destroy()
  })

  it('takes one level off a nested group, and the inner group comes apart next', async () => {
    const nested =
      para('Lead') +
      group(
        '',
        child(1, 0, 0, 100, 60, '', 'One') +
          inner(
            9,
            [150, 30, 170, 130],
            child(2, 150, 30, 80, 80, '', 'Two') + child(3, 240, 110, 80, 50),
          ),
      )
    const { editor, blocks } = await open(nested)
    const [at] = objects(editor)
    const before = boxesAt(editor, at).map((b) => [b.offsetXEmu, b.offsetYEmu])
    select(editor, at)
    expect(ungroupSelectedObjects(editor, blocks)).toBe(true)
    let now = objects(editor)
    expect(now.map((pos) => boxesAt(editor, pos).length)).toEqual([1, 2])
    expect(boxesAt(editor, now[1]).every((b) => b.grouped)).toBe(true)
    expect(boxesAt(editor, now[1]).map((b) => b.txbxIndex)).toEqual([0, undefined])
    const once = await savedBody(editor, nested)
    expect(once.match(/<wp:anchor /g)).toHaveLength(2)
    expect(once.match(/<wpg:wgp[ >]/g)).toHaveLength(1)

    select(editor, now[1])
    expect(canUngroupObjects(editor.state)).toBe(true)
    expect(ungroupSelectedObjects(editor, blocks)).toBe(true)
    now = objects(editor)
    expect(now.map((pos) => boxesAt(editor, pos).length)).toEqual([1, 1, 1])
    expect(
      now.map((pos) => boxesAt(editor, pos)[0]).map((b) => [b.offsetXEmu, b.offsetYEmu]),
    ).toEqual(before)
    const twice = await savedBody(editor, nested)
    expect(twice).not.toContain('wpg:')
    expect(twice.match(/<wp:anchor /g)).toHaveLength(3)
    expect(twice.match(/<w:p>/g)).toHaveLength(nested.match(/<w:p>/g)!.length)
    const ids = Array.from(twice.matchAll(/<wp:docPr id="(\d+)"/g)).map((m) => m[1])
    expect(new Set(ids).size).toBe(3)
    editor.destroy()
  })

  it('an inner group that is not the last child comes apart without adding a paragraph', async () => {
    const nested =
      para('Lead') +
      group(
        '',
        inner(9, [150, 30, 170, 130], child(2, 150, 30, 80, 80) + child(3, 240, 110, 80, 50)) +
          child(1, 0, 0, 100, 60),
      )
    const { editor, blocks } = await open(nested)
    select(editor, objects(editor)[0])
    ungroupSelectedObjects(editor, blocks)
    select(editor, objects(editor)[0])
    expect(ungroupSelectedObjects(editor, blocks)).toBe(true)
    const nodes = objects(editor).map((pos) => editor.state.doc.nodeAt(pos)!)
    expect(nodes.map((node) => node.attrs.carried)).toEqual([false, true, true])
    const xml = await savedBody(editor, nested)
    expect(xml.match(/<w:p>/g)).toHaveLength(2)
    expect(xml.match(/<wp:anchor /g)).toHaveLength(3)
    expect(offsets(xml, 'H')).toEqual([210, 300, 60])
    editor.destroy()
  })

  it('a turned group hands its turn to the children, which stay where they showed', async () => {
    const turned = para('Lead') + group(' rot="2700000"', THREE)
    const { editor, blocks } = await open(turned)
    const [at] = objects(editor)
    const before = boxesAt(editor, at).map((b) => [b.offsetXEmu, b.offsetYEmu, b.rotDeg])
    expect(before.map((b) => b[2])).toEqual([45, 45, 45])
    select(editor, at)
    expect(ungroupSelectedObjects(editor, blocks)).toBe(true)
    const now = objects(editor)
    expect(
      now.map((pos) => boxesAt(editor, pos)[0]).map((b) => [b.offsetXEmu, b.offsetYEmu, b.rotDeg]),
    ).toEqual(before)
    const xml = await savedBody(editor, turned)
    expect(xml.match(/<a:xfrm rot="2700000">/g)).toHaveLength(3)
    editor.destroy()
  })

  it('leaves a group with text above and below it as it is', async () => {
    const banded =
      para('Lead') + group('', THREE).replace('<wp:wrapNone/>', '<wp:wrapTopAndBottom/>')
    const { editor, blocks } = await open(banded)
    select(editor, objects(editor)[0])
    expect(ungroupSelectedObjects(editor, blocks)).toBe(false)
    expect(objects(editor)).toHaveLength(1)
    editor.destroy()
  })

  it('leaves a group placed by alignment as it is', async () => {
    const aligned =
      para('Lead') +
      group('', THREE).replace(
        /<wp:positionH[^>]*>.*?<\/wp:positionH>/,
        '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>',
      )
    const { editor, blocks } = await open(aligned)
    select(editor, objects(editor)[0])
    expect(ungroupSelectedObjects(editor, blocks)).toBe(false)
    expect(objects(editor)).toHaveLength(1)
    editor.destroy()
  })
})
