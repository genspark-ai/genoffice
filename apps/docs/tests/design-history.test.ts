import { Editor } from '@tiptap/core'
import type { Transaction } from '@tiptap/pm/state'
import {
  OFFICE_THEMES,
  listStylesXml,
  parseDocx,
  readDocxStylesXml,
  styleSetUpserts,
  upsertStyleXml,
} from '@genoffice/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { themedDocx } from '../../../packages/docx-engine/tests/helpers/themed-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import {
  designChange,
  designSides,
  pendingDesign,
  recordDesignStep,
  replayedDesign,
  replaysDesign,
  withDesignSide,
  withStyleUpsert,
  type DesignSlot,
  type DesignState,
} from '../src/renderer/editor/design-history'
import { editorExtensions } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const facet = OFFICE_THEMES.find((t) => t.id === 'facet')!
const OFFICE = { major: 'Calibri Light', minor: 'Calibri' }
const start: DesignState = {
  fonts: OFFICE,
  fontsDirty: false,
  colors: { accent1: '4472C4' },
  colorsDirty: false,
  format: null,
  styles: {},
}
const themed: DesignState = {
  ...start,
  fonts: facet.fonts,
  fontsDirty: true,
  colors: facet.colors,
  colorsDirty: true,
  format: { name: facet.name, xml: facet.format },
}

let editor: Editor | null = null
afterEach(() => {
  editor?.destroy()
  editor = null
})

describe('a Design change as an undo step', () => {
  it('says nothing when nothing changed, and only the half that did', () => {
    expect(designSides(start, { ...start })).toBeNull()
    const theme = designSides(start, themed)!
    expect(theme.undo.styles).toBeUndefined()
    expect(theme.undo.theme).toMatchObject({ fonts: OFFICE, fontsDirty: false, format: null })
    expect(theme.redo.theme?.colors).toBe(facet.colors)

    const spaced = { ...start, styles: { Normal: { styleId: 'Normal', pPr: { lineSpacing: 2 } } } }
    const styles = designSides(start, spaced)!
    expect(styles.undo).toEqual({ styles: { Normal: null } })
    expect(styles.redo).toEqual({ styles: { Normal: spaced.styles.Normal } })
  })

  it('puts a side back as it was while the file is the one the change was made on', () => {
    const { undo, redo } = designSides(start, themed)!
    expect(withDesignSide(themed, undo, null)).toEqual(start)
    expect(withDesignSide(start, redo, null)).toEqual(themed)
  })

  it('leaves alone a style the change did not touch', () => {
    const before = { ...start, styles: { Mine: { styleId: 'Mine', rPr: { bold: true } } } }
    const after = {
      ...before,
      styles: withStyleUpsert(before.styles, { styleId: 'Normal', pPr: { lineSpacing: 2 } }),
    }
    const { undo } = designSides(before, after)!
    const later = {
      ...after,
      styles: withStyleUpsert(after.styles, { styleId: 'Mine', rPr: { italic: true } }),
    }
    const undone = withDesignSide(later, undo, null)
    expect(undone.styles.Normal).toBeUndefined()
    expect(undone.styles.Mine).toEqual({ styleId: 'Mine', rPr: { bold: true, italic: true } })
  })

  it('says every value in full once a save has come in between', async () => {
    const bytes = await themedDocx()
    const parsed = await parseDocx(bytes)
    const xml = await readDocxStylesXml(bytes)
    const set = styleSetUpserts('modern', xml)
    const after = { ...themed, styles: set.reduce(withStyleUpsert, {}) }
    const change = designChange(parsed, bytes, start, after)!
    const stated = await change.stated

    // the saved file holds the theme and the set, and nothing is pending
    const saved: DesignState = { ...after, fontsDirty: false, colorsDirty: false, styles: {} }
    const undone = withDesignSide(saved, change.undo, stated.undo)
    expect(undone).toMatchObject({ fonts: OFFICE, fontsDirty: true, colorsDirty: true })
    expect(undone.format).toMatchObject({ name: 'Office Theme' })
    // every style the set wrote goes back to the file's own, and the ones it added go away
    const savedXml = set.reduce(upsertStyleXml, xml)
    const back = Object.values(undone.styles).reduce(upsertStyleXml, savedXml)
    const byId = (styles: string) => new Map(listStylesXml(styles).map((s) => [s.styleId, s.xml]))
    expect(byId(back)).toEqual(byId(xml))

    const redone = withDesignSide({ ...start, styles: {} }, change.redo, stated.redo)
    expect(redone.format).toEqual(after.format)
    expect(byId(Object.values(redone.styles).reduce(upsertStyleXml, xml))).toEqual(byId(savedXml))
  })

  it('is one step of the editor history, replayed on undo and redo and not when recorded', async () => {
    const parsed = await parseDocx(await themedDocx())
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const seen: string[] = []
    editor.on('transaction', ({ transaction }: { transaction: Transaction }) => {
      for (const { side } of replayedDesign(transaction)) seen.push(side)
    })
    const before = editor.state.doc
    const change = designChange(parsed, parsed.internal.originalBytes, start, themed)!
    recordDesignStep(editor, { change })
    // the content is the same, the document object is new: Repeat and a save in flight see an edit
    expect(editor.state.doc).not.toBe(before)
    expect(editor.state.doc.eq(before)).toBe(true)
    expect(seen).toEqual([])

    editor.commands.undo()
    expect(seen).toEqual(['undo'])
    expect(editor.can().undo()).toBe(false)
    editor.commands.redo()
    expect(seen).toEqual(['undo', 'redo'])

    // a step whose command has not read what it needs yet puts nothing back
    recordDesignStep(editor, { change: null })
    editor.commands.undo()
    expect(seen).toEqual(['undo', 'redo'])
    // the file read the change started ends inside the test, not after its teardown
    await change.stated
  })

  it('a change still reading the file waits for a redo when its step is undone, and is dropped by anything newer', async () => {
    const parsed = await parseDocx(await themedDocx())
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const clock = { current: 0 }
    const seen: string[] = []
    editor.on('transaction', ({ transaction }: { transaction: Transaction }) => {
      if (replaysDesign(transaction)) clock.current++
      for (const { side } of replayedDesign(transaction)) seen.push(side)
    })
    const change = designChange(parsed, parsed.internal.originalBytes, start, themed)!
    // the read of styles.xml the pick waits on, held open until the test lets it go
    let finishRead = () => {}
    let read = new Promise<void>((resolve) => (finishRead = resolve))
    const landed: string[] = []
    const pick = (name: string) => {
      const slot: DesignSlot = { change: null }
      recordDesignStep(editor!, slot)
      const wanted = pendingDesign(clock, slot)
      return read.then(() => {
        const fate = wanted()
        if (fate) slot.change = change
        landed.push(`${name}: ${fate ?? 'dropped'}`)
      })
    }
    const nextRead = () => {
      finishRead()
      read = new Promise<void>((resolve) => (finishRead = resolve))
    }

    // undone before the read ended: it stays off the page, and Redo puts it there
    const undone = pick('undone')
    editor.commands.undo()
    nextRead()
    await undone
    expect(landed).toEqual(['undone: keep'])
    expect(seen).toEqual([])
    editor.commands.redo()
    expect(seen).toEqual(['redo'])

    // undone and redone before the read ended: it applies as if left alone
    const redone = pick('redone')
    editor.commands.undo()
    editor.commands.redo()
    nextRead()
    await redone
    expect(landed.at(-1)).toBe('redone: apply')

    // two picks on one read: the later one is the look that stays
    const first = pick('first')
    const second = pick('second')
    nextRead()
    await Promise.all([first, second])
    expect(landed.slice(-2)).toEqual(['first: dropped', 'second: apply'])

    // an undo that goes on to a step holding its change counts as newer
    const late = pick('late')
    editor.commands.undo()
    editor.commands.undo()
    nextRead()
    await late
    expect(landed.at(-1)).toBe('late: dropped')
    await change.stated
  })
})
