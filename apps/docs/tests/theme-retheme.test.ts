import { Editor } from '@tiptap/core'
import type { Node as PMNode } from '@tiptap/pm/model'
import {
  OFFICE_THEMES,
  parseDocx,
  saveDocx,
  themedDocxBytes,
  type ParsedDocFull,
} from '@genoffice/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { themedDocx } from '../../../packages/docx-engine/tests/helpers/themed-docx'
import { blocksToPmDoc, pmDocOptions, pmDocToSavePlan } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyThemePatch, rethemeBody, themePatch } from '../src/renderer/editor/theme-retheme'
import { retintThemeRuns } from '../src/renderer/editor/theme-retint'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const theme = (id: string) => {
  const t = OFFICE_THEMES.find((x) => x.id === id)!
  return { fonts: t.fonts, colors: t.colors, format: { name: t.name, xml: t.format } }
}
const facet = theme('facet')
const ion = theme('ion')

let editor: Editor | null = null
afterEach(() => {
  editor?.destroy()
  editor = null
})

const pmJson = (parsed: ParsedDocFull) =>
  blocksToPmDoc(parsed.blocks, undefined, pmDocOptions(parsed))

async function open() {
  const bytes = await themedDocx()
  const parsed = await parseDocx(bytes)
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: pmJson(parsed) as never,
  })
  const read = async (edit: typeof facet): Promise<PMNode> =>
    editor!.schema.nodeFromJSON(pmJson(await parseDocx(await themedDocxBytes(bytes, edit))))
  return { editor, bytes, parsed, base: editor.state.doc, read }
}

const find = (doc: PMNode, test: (node: PMNode) => boolean): { node: PMNode; pos: number } => {
  let hit: { node: PMNode; pos: number } | null = null
  doc.descendants((node, pos) => {
    if (!hit && test(node)) hit = { node, pos }
    return !hit
  })
  if (!hit) throw new Error('node not found')
  return hit
}
const textPos = (doc: PMNode, text: string) => find(doc, (n) => n.isText && n.text!.includes(text))

describe('a theme pick on the page', () => {
  it('leaves the untouched document exactly as save and reopen reads it', async () => {
    const { editor, parsed, base, read } = await open()
    const themed = await read(facet)
    expect(base.eq(themed)).toBe(false)
    rethemeBody(editor, base, themed)
    expect(editor.state.doc.eq(themed)).toBe(true)

    // against the file as read under the theme nothing is an edit: the save keeps every block as it is
    const themedBlocks = (await parseDocx(await themedDocxBytes(await themedDocx(), facet))).blocks
    const plan = pmDocToSavePlan(editor.getJSON() as never, themedBlocks)
    expect(plan.saveBlocks.map((b) => b.kind)).toEqual(plan.saveBlocks.map(() => 'original'))
    const saved = await parseDocx(
      await saveDocx(parsed, plan.saveBlocks, {
        themeFonts: facet.fonts,
        themeColors: facet.colors,
        themeFormat: facet.format,
      }),
    )
    expect(editor.state.doc.eq(editor.schema.nodeFromJSON(pmJson(saved)))).toBe(true)
  })

  it('recolours the table style, the cell, the text box and the chart', async () => {
    const { editor, base, read } = await open()
    const before = JSON.stringify(base.toJSON())
    for (const hex of ['4472C4', 'ED7D31', '70AD47']) expect(before).toContain(hex)
    rethemeBody(editor, base, await read(facet))
    const after = JSON.stringify(editor.state.doc.toJSON())
    for (const hex of ['90C226', '54A021', '918655']) expect(after).toContain(hex)
    // the literal red run is the only colour of the file's own left
    expect(after).toContain('FF0000')
    for (const hex of ['4472C4', 'ED7D31', '70AD47'])
      expect(after.replaceAll(/w:(val|fill|color)=\\+"[0-9A-F]{6}/g, '')).not.toContain(hex)
    expect(after).toContain('Trebuchet MS')
  })

  it('goes on to a second theme and back to the file as it was', async () => {
    const { editor, base, read } = await open()
    const [onFacet, onIon] = [await read(facet), await read(ion)]
    rethemeBody(editor, base, onFacet)
    rethemeBody(editor, onFacet, onIon)
    expect(editor.state.doc.eq(onIon)).toBe(true)
    rethemeBody(editor, onIon, base)
    expect(editor.state.doc.eq(base)).toBe(true)
  })

  it('is not an undo step of its own', async () => {
    const { editor, base, read } = await open()
    rethemeBody(editor, base, await read(facet))
    expect(editor.can().undo()).toBe(false)
  })

  it('keeps what the user changed and still reaches text typed next to a themed run', async () => {
    const { editor, base, read } = await open()
    const themed = await read(facet)
    const tinted = textPos(editor.state.doc, 'tinted')
    editor.commands.insertContentAt(tinted.pos + 2, 'XY')
    const plain = textPos(editor.state.doc, 'plain')
    editor.commands.insertContentAt(plain.pos, 'typed ')
    const cell = find(
      editor.state.doc,
      (n) => n.type.name === 'docTableCell' && n.textContent === 'North',
    )
    const mine = { ...cell.node.attrs, fill: '123456' }
    editor.view.dispatch(editor.state.tr.setNodeMarkup(cell.pos, undefined, mine))

    rethemeBody(editor, base, themed)
    const doc = editor.state.doc
    expect(
      find(doc, (n) => n.type.name === 'docTableCell' && n.textContent === 'North').node.attrs.fill,
    ).toBe('123456')
    const typedIn = textPos(doc, 'XY').node
    const wanted = textPos(themed, 'tinted').node
    expect(typedIn.marks.map((m) => m.attrs.color)).toEqual(wanted.marks.map((m) => m.attrs.color))
    expect(textPos(doc, 'typed plain')).toBeTruthy()
    // every other cell of the table took the theme
    const fills = (root: PMNode) => {
      const out: unknown[] = []
      root.descendants((n) => {
        if (n.type.name === 'docTableCell' && n.textContent !== 'North') out.push(n.attrs.fill)
      })
      return out
    }
    expect(fills(doc)).toEqual(fills(themed))
  })

  it('reaches the same page when the run colours were repainted first', async () => {
    const { editor, base, read } = await open()
    const themed = await read(facet)
    retintThemeRuns(editor, facet.colors)
    expect(editor.state.doc.eq(base)).toBe(false)
    rethemeBody(editor, base, themed)
    expect(editor.state.doc.eq(themed)).toBe(true)
  })

  it('finds nothing to do between two readings under one theme', async () => {
    const { base, editor } = await open()
    expect(themePatch(base, base).size).toBe(0)
    const tr = editor.state.tr
    expect(applyThemePatch(tr, themePatch(base, base))).toBe(false)
    expect(tr.docChanged).toBe(false)
  })
})
