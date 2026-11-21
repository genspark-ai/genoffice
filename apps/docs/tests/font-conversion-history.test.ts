/**
 * Convert Permanently in the undo history: the fonts the save has to write
 * follow Undo and Redo, before a save and after one.
 */
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx, writeFontConversions } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { convertFontsInEditor } from '../src/renderer/editor/convert-fonts'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  fontsBeingRestored,
  fontsToWrite,
  replayedFontConversions,
  withFontsConverted,
  withFontsReverted,
  type FontConversions,
} from '../src/renderer/editor/font-conversion-history'

const liveEditors: Editor[] = []
afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

const fonts = (name: string) => `<w:rPr><w:rFonts w:ascii="${name}" w:hAnsi="${name}"/></w:rPr>`
const run = (name: string, text: string) => `<w:r>${fonts(name)}<w:t>${text}</w:t></w:r>`
const CONVERT = { 'Old Sans': 'Arial' }

const source = () =>
  buildDocx({
    bodyXml:
      `<w:p>${run('Old Sans', 'body')}${run('Arial', 'was Arial')}</w:p>` +
      '<w:p><w:pPr><w:pStyle w:val="Fancy"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>',
    extraStylesXml:
      `<w:style w:type="paragraph" w:styleId="Fancy"><w:name w:val="Fancy"/>${fonts('Old Sans')}</w:style>` +
      `<w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/>${fonts('Arial')}</w:style>`,
    headerXml: `<w:p>${run('Old Sans', 'head')}${run('Arial', 'head too')}</w:p>`,
  })

const PARTS = ['word/document.xml', 'word/styles.xml', 'word/header1.xml']
const partsOf = async (bytes: Uint8Array): Promise<Record<string, string>> => {
  const zip = await JSZip.loadAsync(bytes)
  const out: Record<string, string> = {}
  for (const name of PARTS) out[name] = (await zip.file(name)?.async('string')) ?? ''
  return out
}
const fontSlots = (xml: string) => (xml.match(/w:ascii="[^"]*"/g) ?? []).join(' ')

/** an editor on the document with the app's share: state that follows the history, and the save */
async function session(bytes: Uint8Array) {
  let parsed = await parseDocx(bytes)
  let state: FontConversions | null = null
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  liveEditors.push(editor)
  editor.on('transaction', ({ transaction }) => {
    for (const step of replayedFontConversions(transaction))
      state = step.converting
        ? withFontsConverted(state, step.fonts)
        : withFontsReverted(state, step.fonts)
  })
  return {
    editor,
    state: () => state,
    convert(map: Record<string, string>) {
      convertFontsInEditor(editor, map)
      state = withFontsConverted(state, map)
    },
    async save(): Promise<Uint8Array> {
      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const built = await saveDocx(parsed, plan.saveBlocks)
      const out = await writeFontConversions(built, state?.pending ?? {}, state?.written ?? [])
      parsed = await parseDocx(out.bytes)
      // the app keeps its history when the saved document reads back as the editor holds it
      expect(editor.state.doc.eq(editor.schema.nodeFromJSON(blocksToPmDoc(parsed.blocks)))).toBe(
        true,
      )
      state = out.written.length > 0 ? { pending: {}, written: out.written } : null
      return out.bytes
    },
  }
}

describe('the fonts a save has to write', () => {
  it('are gone after an undo before any save, and back after a redo', () => {
    const converted = withFontsConverted(null, CONVERT)
    expect(converted).toEqual({ pending: CONVERT, written: [] })
    expect(fontsToWrite(converted)).toBe(true)
    expect(withFontsReverted(converted, CONVERT)).toBeNull()
    expect(fontsToWrite(null)).toBe(false)
  })

  it('take back only the fonts of the step that was undone', () => {
    const two = withFontsConverted(withFontsConverted(null, CONVERT), { 'Lost Serif': 'Georgia' })
    expect(withFontsReverted(two, { 'Lost Serif': 'Georgia' })).toEqual({
      pending: CONVERT,
      written: [],
    })
  })

  it('ask a saved conversion to be put back on undo, and drop the request on redo', () => {
    const written = { fonts: CONVERT, kept: CONVERT, parts: {} }
    const saved: FontConversions = { pending: {}, written: [written] }
    expect(fontsToWrite(saved)).toBe(false)
    const undone = withFontsReverted(saved, CONVERT)!
    expect(undone.written[0].kept).toEqual({})
    expect(fontsToWrite(undone)).toBe(true)
    expect(fontsBeingRestored(undone)).toEqual(['Old Sans'])
    const redone = withFontsConverted(undone, CONVERT)!
    expect(redone.pending).toEqual({})
    expect(redone.written[0].kept).toBe(written.fonts)
    expect(fontsToWrite(redone)).toBe(false)
  })
})

describe('Convert Permanently, Undo and Redo around a save', () => {
  it('writes nothing of the conversion when it was undone before the save', async () => {
    const original = await source()
    const s = await session(original)
    s.convert(CONVERT)
    expect(s.state()?.pending).toEqual(CONVERT)
    s.editor.commands.undo()
    expect(s.state()).toBeNull()
    const saved = await partsOf(await s.save())
    expect(saved).toEqual(await partsOf(original))
  })

  it('is one undo step even when no run of the body names the font', async () => {
    const s = await session(
      await buildDocx({
        bodyXml: '<w:p><w:pPr><w:pStyle w:val="Fancy"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>',
        extraStylesXml: `<w:style w:type="paragraph" w:styleId="Fancy"><w:name w:val="Fancy"/>${fonts('Old Sans')}</w:style>`,
      }),
    )
    s.convert(CONVERT)
    expect(s.editor.can().undo()).toBe(true)
    s.editor.commands.undo()
    expect(s.state()).toBeNull()
    expect(s.editor.can().undo()).toBe(false)
    s.editor.commands.redo()
    expect(s.state()?.pending).toEqual(CONVERT)
  })

  it('converts every part again after a redo', async () => {
    const s = await session(await source())
    s.convert(CONVERT)
    s.editor.commands.undo()
    s.editor.commands.redo()
    const saved = await partsOf(await s.save())
    for (const name of PARTS) expect(saved[name], name).not.toContain('Old Sans')
  })

  it('writes the original font back into every part when the conversion is undone after a save', async () => {
    const original = await source()
    const before = await partsOf(original)
    const s = await session(original)
    s.convert(CONVERT)
    const converted = await partsOf(await s.save())
    for (const name of PARTS) expect(converted[name], name).not.toContain('Old Sans')
    // only the slots that named the font: no East Asian slot is made up for the stand-in
    expect(converted['word/document.xml']).not.toContain('w:eastAsia')
    expect(fontsToWrite(s.state())).toBe(false)

    s.editor.commands.undo()
    expect(fontsToWrite(s.state())).toBe(true)
    const undone = await partsOf(await s.save())
    expect(undone['word/styles.xml']).toBe(before['word/styles.xml'])
    expect(undone['word/header1.xml']).toBe(before['word/header1.xml'])
    // the body is written from the editor: the same fonts in the same order
    expect(fontSlots(undone['word/document.xml'])).toBe(fontSlots(before['word/document.xml']))
    expect(s.state()).toBeNull()

    // and a redo after that save converts everything once more
    s.editor.commands.redo()
    expect(s.state()?.pending).toEqual(CONVERT)
    const redone = await partsOf(await s.save())
    for (const name of PARTS) expect(redone[name], name).not.toContain('Old Sans')
  })

  it('knows the font is back after an undo when only runs of the body named it', async () => {
    const original = await buildDocx({ bodyXml: `<w:p>${run('Old Sans', 'body')}</w:p>` })
    const s = await session(original)
    s.convert(CONVERT)
    await s.save()
    s.editor.commands.undo()
    expect(fontsBeingRestored(s.state())).toEqual(['Old Sans'])
    expect(fontSlots((await partsOf(await s.save()))['word/document.xml'])).toBe(
      fontSlots((await partsOf(original))['word/document.xml']),
    )
    expect(s.state()).toBeNull()
  })

  it('keeps the file converted when the undo is redone before the next save', async () => {
    const s = await session(await source())
    s.convert(CONVERT)
    const converted = await s.save()
    s.editor.commands.undo()
    s.editor.commands.redo()
    expect(fontsToWrite(s.state())).toBe(false)
    expect(await partsOf(await s.save())).toEqual(await partsOf(converted))
  })
})
