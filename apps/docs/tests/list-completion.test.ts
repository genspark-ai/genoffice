/**
 * List completion: Set Numbering Value, Adjust List Indents (pending and parsed
 * definitions), picture bullets, CJK number formats, level-to-style links,
 * Change List Level and the line-spacing menu's space before/after toggle.
 * Every write is checked in the saved package, not just in the pending state.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import {
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  computeListMarkers,
  customLevelXml,
  formatNumber,
  parseDocx,
  saveDocx,
  type NumberingDef,
  type StyleUpsert,
} from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocToSavePlan } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { EMPTY_PENDING_NUMBERING, type PendingNumbering } from '../src/renderer/doc-state'
import {
  addListStyle,
  addPictureBullet,
  allocateListNumId,
  changeListLevel,
  createCustomListDef,
  currentListValue,
  deleteListStyle,
  editListLevel,
  listStyleNum,
  listStylesInUse,
  modifyListStyle,
  pruneUnreferencedNumbering,
  setNumberingValue,
  type NumberingContext,
} from '../src/renderer/numbering-actions'
import {
  MULTILEVEL_LIBRARY,
  bulletPresetLevels,
  documentListPresets,
  numberPresetLevels,
  presetKey,
  previewLevelText,
} from '../src/renderer/list-presets'
import { toggleParaSpace } from '../src/renderer/components/ribbon-tabs'
import {
  drawListStyleSide,
  modifyListStyleAsStep,
  replayedListStyles,
  withListStylePending,
  type ListStyleSide,
} from '../src/renderer/editor/list-style-history'
import {
  NO_LIST_STYLE_ID,
  listStyleEdits,
  listStyleEntries,
  listStylePreview,
  nineListLevels,
} from '../src/renderer/list-styles'

const PNG_1PX =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

interface Harness {
  editor: Editor
  parsed: Awaited<ReturnType<typeof parseDocx>>
  ctx: NumberingContext
  pending: () => PendingNumbering
  styleUpserts: StyleUpsert[]
  save: () => Promise<Awaited<ReturnType<typeof parseDocx>> & { zip: JSZip }>
}

async function harness(bytes?: Uint8Array): Promise<Harness> {
  const parsed = await parseDocx(bytes ?? (await buildBlankDocx()))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.push(editor)
  editor.storage.listNumbering.defs = new Map(parsed.numbering)
  let pending: PendingNumbering = EMPTY_PENDING_NUMBERING
  const styleUpserts: StyleUpsert[] = []
  const ctx: NumberingContext = {
    editor,
    doc: { parsed } as never,
    pendingNumbering: pending,
    setPendingNumbering: (u) => {
      pending = u(pending)
      ctx.pendingNumbering = pending
    },
    numIdFloorRef: { current: 0 },
    setStatus: () => {},
    linkStyle: (styleId, numId, ilvl) => {
      styleUpserts.push({ styleId, pPr: { numPr: { numId, ilvl } } })
    },
    renameStyle: (styleId, name) => {
      styleUpserts.push({ styleId, name })
    },
  }
  const save = async () => {
    const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
    const bytes = await saveDocx(parsed, plan.saveBlocks, {
      numbering: pruneUnreferencedNumbering(pending, editor.getJSON() as never),
      styleUpserts: styleUpserts.length ? styleUpserts : undefined,
    })
    const zip = await JSZip.loadAsync(bytes)
    return Object.assign(await parseDocx(bytes), { zip })
  }
  return { editor, parsed, ctx, pending: () => pending, styleUpserts, save }
}

/** replace the document outside the undo history, so undo only sees the action under test */
function setDoc(editor: Editor, content: unknown[]) {
  const nodes = content.map((json) => editor.schema.nodeFromJSON(json))
  const { tr } = editor.state
  editor.view.dispatch(
    tr.replaceWith(0, editor.state.doc.content.size, nodes).setMeta('addToHistory', false),
  )
  editor.commands.setTextSelection(editor.state.doc.content.size - 2)
}

/** three list items sharing numId, caret left in the last one */
function threeItems(editor: Editor, numId: string, kind: 'bullet' | 'ordered' = 'ordered') {
  setDoc(
    editor,
    ['one', 'two', 'three'].map((text) => ({
      type: 'docListItem',
      attrs: { kind, numId, ilvl: 0 },
      content: [{ type: 'text', text }],
    })),
  )
}

const markers = (editor: Editor) => {
  const refs: Array<{ numId: string | null; ilvl: number }> = []
  editor.state.doc.descendants((n) => {
    if (n.type.name === 'docListItem')
      refs.push({ numId: (n.attrs.numId as string) ?? null, ilvl: Number(n.attrs.ilvl) || 0 })
    return false
  })
  return computeListMarkers(refs, editor.storage.listNumbering.defs as Map<string, NumberingDef>)
}

const numberingXml = (zip: JSZip) => zip.file('word/numbering.xml')!.async('string')

describe('Set Numbering Value', () => {
  it('starts a new list at the value: startOverride in the save, markers update live, undo reverts', async () => {
    const h = await harness()
    const numId = allocateListNumId(h.ctx, 'ordered')!
    threeItems(h.editor, numId)
    expect(markers(h.editor)).toEqual(['1.', '2.', '3.'])
    expect(currentListValue(h.editor, numId, 0)).toBe(3)

    expect(setNumberingValue(h.ctx, { mode: 'restart', value: 5 })).toBe(true)
    const third = h.editor.state.doc.child(2)
    expect(third.attrs.numId).not.toBe(numId)
    expect(markers(h.editor)).toEqual(['1.', '2.', '5.'])
    const restart = h.pending().restartNums.find((r) => r.numId === third.attrs.numId)!
    expect(restart.startOverrides).toEqual({ 0: 5 })

    const saved = await h.save()
    expect(await numberingXml(saved.zip)).toContain('<w:startOverride w:val="5"/>')
    const reopened = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(saved.blocks) as never,
    })
    editors.push(reopened)
    reopened.storage.listNumbering.defs = saved.numbering
    expect(markers(reopened)).toEqual(['1.', '2.', '5.'])

    // undo only touches the document: the item goes back to the old list and the
    // unreferenced restart definition drops out of the save
    const savedRestart = () =>
      pruneUnreferencedNumbering(h.pending(), h.editor.getJSON() as never).restartNums.some(
        (r) => r.numId === third.attrs.numId,
      )
    h.editor.commands.undo()
    expect(h.editor.state.doc.child(2).attrs.numId).toBe(numId)
    expect(savedRestart()).toBe(false)
    h.editor.commands.redo()
    expect(h.editor.state.doc.child(2).attrs.numId).toBe(third.attrs.numId)
    expect(savedRestart()).toBe(true)
  })

  it('continues the previous list, advancing to the value when asked', async () => {
    const h = await harness()
    const a = allocateListNumId(h.ctx, 'ordered')!
    const b = createCustomListDef(h.ctx, numberPresetLevels('decimal', '%1.'))!
    setDoc(h.editor, [
      {
        type: 'docListItem',
        attrs: { kind: 'ordered', numId: a, ilvl: 0 },
        content: [{ type: 'text', text: 'a1' }],
      },
      { type: 'docParagraph', content: [{ type: 'text', text: 'gap' }] },
      {
        type: 'docListItem',
        attrs: { kind: 'ordered', numId: b, ilvl: 0 },
        content: [{ type: 'text', text: 'b1' }],
      },
    ])
    expect(markers(h.editor)).toEqual(['1.', '1.'])
    setNumberingValue(h.ctx, { mode: 'continue' })
    expect(h.editor.state.doc.child(2).attrs.numId).toBe(a)
    expect(markers(h.editor)).toEqual(['1.', '2.'])

    h.editor.commands.undo()
    expect(h.editor.state.doc.child(2).attrs.numId).toBe(b)

    // continue AND advance is one confirmation, so one undo step
    setNumberingValue(h.ctx, { mode: 'continue', value: 7 })
    expect(markers(h.editor)).toEqual(['1.', '7.'])
    const numId = h.editor.state.doc.child(2).attrs.numId as string
    const restart = h.pending().restartNums.find((r) => r.numId === numId)!
    const store = h.editor.storage.listNumbering.defs as Map<string, NumberingDef>
    expect(restart.abstractNumId).toBe(store.get(a)!.abstractNumId)
    h.editor.commands.undo()
    expect(h.editor.state.doc.child(2).attrs.numId).toBe(b)
    expect(markers(h.editor)).toEqual(['1.', '1.'])
  })
})

describe('Set Numbering Value on a definition created in this session', () => {
  it('points the restart num at the abstractNum the save assigns, never at "pending-"', async () => {
    const h = await harness()
    const numId = createCustomListDef(h.ctx, numberPresetLevels('decimal', '%1)'))!
    threeItems(h.editor, numId)
    setNumberingValue(h.ctx, { mode: 'restart', value: 9 })
    const restartId = h.editor.state.doc.child(2).attrs.numId as string
    expect(h.pending().restartNums[0].abstractNumId).toBe(`pending-${numId}`)
    const saved = await h.save()
    const xml = await numberingXml(saved.zip)
    expect(xml).not.toContain('pending-')
    const absId = new RegExp(`<w:num w:numId="${numId}"><w:abstractNumId w:val="(\\d+)"/>`).exec(
      xml,
    )![1]
    expect(xml).toContain(
      `<w:num w:numId="${restartId}"><w:abstractNumId w:val="${absId}"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="9"/></w:lvlOverride></w:num>`,
    )
    expect(saved.numbering.get(restartId)!.startOverrides).toEqual({ 0: 9 })
  })
})

describe('Adjust List Indents', () => {
  it('rewrites the level of a pending definition: ind, suff and tab stop reach numbering.xml', async () => {
    const h = await harness()
    const numId = createCustomListDef(h.ctx, numberPresetLevels('decimal', '%1.'))!
    threeItems(h.editor, numId)
    expect(
      editListLevel(h.ctx, { indentLeft: 1134, hanging: 567, suff: 'space', tabStop: 1134 }),
    ).toBe(true)
    const def = h.pending().newDefs.find((d) => d.numId === numId)!
    expect(def.levels![0]).toMatchObject({ indentLeft: 1134, hanging: 567, suff: 'space' })
    expect(def.levels![1].indentLeft).toBe(1440)
    const live = (h.editor.storage.listNumbering.defs as Map<string, NumberingDef>).get(numId)!
    expect(live.levels[0]).toMatchObject({ indentLeft: 1134, hanging: 567, suff: 'space' })

    const saved = await h.save()
    const xml = await numberingXml(saved.zip)
    expect(xml).toContain('<w:suff w:val="space"/>')
    expect(xml).toContain('<w:ind w:left="1134" w:hanging="567"/>')
    expect(xml).toContain('<w:tab w:val="num" w:pos="1134"/>')
    expect(saved.numbering.get(numId)!.levels[0]).toMatchObject({
      indentLeft: 1134,
      hanging: 567,
      suff: 'space',
    })
  })

  it("keeps the parsed level's marker weight, style link, restart rule and tab stop when only the indents change", async () => {
    const h = await harness()
    // give the blank template's decimal level 0 the properties the dialogs do not edit
    const rich =
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlRestart w:val="0"/>' +
      '<w:pStyle w:val="Heading1"/><w:lvlText w:val="%1."/><w:legacy w:legacy="1"/><w:lvlJc w:val="left"/>' +
      '<w:pPr><w:tabs><w:tab w:val="num" w:pos="1440"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
      '<w:rPr><w:rFonts w:ascii="Georgia"/><w:b/><w:bCs/><w:i/><w:color w:val="112233"/><w:sz w:val="28"/></w:rPr></w:lvl>'
    const seeded = await (async () => {
      const zip = await JSZip.loadAsync(await buildBlankDocx())
      const xml = await zip.file('word/numbering.xml')!.async('string')
      const abs = /<w:abstractNum w:abstractNumId="1">[\s\S]*?<\/w:abstractNum>/.exec(xml)![0]
      const patched = abs.replace(/<w:lvl w:ilvl="0">[\s\S]*?<\/w:lvl>/, rich)
      zip.file('word/numbering.xml', xml.replace(abs, patched))
      return parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    })()
    const parsed0 = seeded.numbering.get(BLANK_ORDERED_NUM_ID)!.levels[0]
    expect(parsed0).toMatchObject({
      bold: true,
      italic: true,
      pStyle: 'Heading1',
      lvlRestart: 0,
      tabStop: 1440,
    })
    h.ctx.doc = { parsed: seeded } as never
    h.editor.storage.listNumbering.defs = new Map(seeded.numbering)
    threeItems(h.editor, BLANK_ORDERED_NUM_ID)
    editListLevel(h.ctx, { indentLeft: 900, hanging: 300, suff: 'space' })
    const plan = pmDocToSavePlan(h.editor.getJSON() as never, seeded.blocks)
    const bytes = await saveDocx(seeded, plan.saveBlocks, {
      numbering: pruneUnreferencedNumbering(h.pending(), h.editor.getJSON() as never),
    })
    const zip = await JSZip.loadAsync(bytes)
    const xml = await numberingXml(zip)
    const decimalAbs = /<w:abstractNum w:abstractNumId="1">[\s\S]*?<\/w:abstractNum>/.exec(xml)![0]
    const lvl = /<w:lvl w:ilvl="0">[\s\S]*?<\/w:lvl>/.exec(decimalAbs)![0]
    expect(lvl).toBe(
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlRestart w:val="0"/>' +
        '<w:pStyle w:val="Heading1"/><w:suff w:val="space"/><w:lvlText w:val="%1."/><w:legacy w:legacy="1"/><w:lvlJc w:val="left"/>' +
        '<w:pPr><w:tabs><w:tab w:val="num" w:pos="1440"/></w:tabs><w:ind w:left="900" w:hanging="300"/></w:pPr>' +
        '<w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:eastAsia="Georgia" w:cs="Georgia" w:hint="default"/><w:b/><w:bCs/><w:i/><w:color w:val="112233"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:lvl>',
    )
    const reparsed = await parseDocx(bytes)
    expect(reparsed.numbering.get(BLANK_ORDERED_NUM_ID)!.levels[0]).toMatchObject({
      indentLeft: 900,
      hanging: 300,
      suff: 'space',
      bold: true,
      italic: true,
      pStyle: 'Heading1',
      lvlRestart: 0,
      tabStop: 1440,
      font: 'Georgia',
      color: '112233',
    })
  })

  it('rewrites one level of a parsed definition in place and leaves the other levels alone', async () => {
    const h = await harness()
    threeItems(h.editor, BLANK_ORDERED_NUM_ID)
    const before = h.parsed.numbering.get(BLANK_ORDERED_NUM_ID)!
    expect(before.levels[1].indentLeft).toBe(1440)
    editListLevel(h.ctx, { indentLeft: 900, hanging: 300 })
    expect(h.pending().levelEdits).toEqual([
      {
        abstractNumId: before.abstractNumId,
        ilvl: 0,
        level: expect.objectContaining({ indentLeft: 900, hanging: 300, numFmt: 'decimal' }),
      },
    ])
    const saved = await h.save()
    const after = saved.numbering.get(BLANK_ORDERED_NUM_ID)!
    expect(after.levels[0]).toMatchObject({ indentLeft: 900, hanging: 300, lvlText: '%1.' })
    expect(after.levels[1].indentLeft).toBe(1440)
    expect(after.levels[4].indentLeft).toBe(3600)
    // the level was swapped, not duplicated
    const xml = await numberingXml(saved.zip)
    expect(xml.match(/<w:lvl w:ilvl="0">/g)).toHaveLength(2)
  })
})

describe('Define New Bullet: picture bullet', () => {
  it('lands the image as a media part with a numbering relationship and a numPicBullet', async () => {
    const h = await harness()
    const id = addPictureBullet(h.ctx, { base64: PNG_1PX, mime: 'image/png' })
    expect(id).toBe(1)
    const numId = createCustomListDef(h.ctx, bulletPresetLevels('•', { picBulletId: id }))!
    threeItems(h.editor, numId, 'bullet')
    const live = (h.editor.storage.listNumbering.defs as Map<string, NumberingDef>).get(numId)!
    expect(live.levels[0].picBulletSrc).toBe(`data:image/png;base64,${PNG_1PX}`)

    const saved = await h.save()
    const xml = await numberingXml(saved.zip)
    expect(xml).toContain('<w:numPicBullet w:numPicBulletId="1">')
    expect(xml).toContain('<w:lvlPicBulletId w:val="1"/>')
    expect(xml.indexOf('<w:numPicBullet')).toBeLessThan(xml.indexOf('<w:abstractNum '))
    const rels = await saved.zip.file('word/_rels/numbering.xml.rels')!.async('string')
    expect(rels).toMatch(/Target="media\/aidocs\d+\.png"/)
    expect(saved.zip.file(/^word\/media\/aidocs\d+\.png$/)).toHaveLength(1)
    expect(saved.numbering.get(numId)!.levels[0].picBulletSrc).toMatch(/^data:image\/png;base64,/)
    // an unreferenced picture is not written
    const orphan = addPictureBullet(h.ctx, { base64: PNG_1PX, mime: 'image/png' })
    expect(orphan).toBe(2)
    const pruned = pruneUnreferencedNumbering(h.pending(), h.editor.getJSON() as never)
    expect(pruned.picBullets.map((p) => p.id)).toEqual([1])
  })
})

describe('Define New Bullet / Number Format on a nested item', () => {
  it("puts the chosen glyph, picture and marker formatting on the caret's level, not level 1", async () => {
    const h = await harness()
    const id = addPictureBullet(h.ctx, { base64: PNG_1PX, mime: 'image/png' })
    const bullets = bulletPresetLevels('\u2726', { picBulletId: id, color: 'FF0000' }, 2)
    expect(bullets[2]).toMatchObject({ lvlText: '\u2726', picBulletId: id, color: 'FF0000' })
    expect(bullets[0].picBulletId).toBeUndefined()
    expect(bullets[0].lvlText).not.toBe('\u2726')
    const numbers = numberPresetLevels('upperRoman', '%1)', { start: 4, bold: true }, 2)
    expect(numbers[2]).toMatchObject({ lvlText: '%3)', start: 4, bold: true, indentLeft: 2160 })
    expect(numbers[0].start).toBeUndefined()
    // every caret level 1-9 yields nine defined glyphs, the chosen one on the caret's level
    for (let at = 0; at < 9; at++) {
      const deep = bulletPresetLevels('\u2726', {}, at)
      expect(deep.map((l) => l.lvlText)).toHaveLength(9)
      expect(deep.every((l) => typeof l.lvlText === 'string' && l.lvlText.length > 0)).toBe(true)
      expect(deep[at].lvlText).toBe('\u2726')
    }

    const numId = createCustomListDef(h.ctx, bullets)!
    setDoc(h.editor, [
      {
        type: 'docListItem',
        attrs: { kind: 'bullet', numId, ilvl: 2 },
        content: [{ type: 'text', text: 'nested' }],
      },
    ])
    const live = (h.editor.storage.listNumbering.defs as Map<string, NumberingDef>).get(numId)!
    expect(live.levels[2].picBulletSrc).toBe(`data:image/png;base64,${PNG_1PX}`)
    const saved = await h.save()
    expect(saved.numbering.get(numId)!.levels[2]).toMatchObject({
      picBulletId: id,
      color: 'FF0000',
    })
  })
})

describe('Define New Number Format / Multilevel', () => {
  it('formats CJK and enclosed numbers and previews them', () => {
    expect(formatNumber(3, 'chineseCounting')).toBe('\u4e09')
    expect(formatNumber(12, 'chineseCountingThousand')).toBe('\u5341\u4e8c')
    expect(formatNumber(2, 'ideographTraditional')).toBe('\u4e59')
    expect(formatNumber(1, 'decimalEnclosedCircle')).toBe('①')
    expect(formatNumber(4, 'koreanDigital')).toBe('사')
    const cjk = MULTILEVEL_LIBRARY[MULTILEVEL_LIBRARY.length - 1]
    expect(previewLevelText(cjk, 0)).toBe('\u4e00\u3001')
    expect(previewLevelText(cjk, 1)).toBe('(\u4e00)')
    const legal = MULTILEVEL_LIBRARY[1].map((l) => ({ ...l, isLgl: true }))
    expect(previewLevelText(legal, 2)).toBe('1.1.1.')
  })

  it('writes marker font, size, color, weight, start and alignment on the level', async () => {
    const h = await harness()
    const numId = createCustomListDef(
      h.ctx,
      numberPresetLevels('upperRoman', '%1)', {
        start: 4,
        font: 'Georgia',
        szHalfPoints: 28,
        color: 'FF0000',
        bold: true,
        lvlJc: 'right',
      }),
    )!
    threeItems(h.editor, numId)
    expect(markers(h.editor)).toEqual(['IV)', 'V)', 'VI)'])
    const saved = await h.save()
    const xml = await numberingXml(saved.zip)
    expect(xml).toContain('<w:start w:val="4"/><w:numFmt w:val="upperRoman"/>')
    expect(xml).toContain('<w:lvlJc w:val="right"/>')
    expect(xml).toContain('<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:eastAsia="Georgia"')
    expect(xml).toContain('<w:b/><w:color w:val="FF0000"/><w:sz w:val="28"/>')
    expect(saved.numbering.get(numId)!.levels[0]).toMatchObject({
      font: 'Georgia',
      szHalfPoints: 28,
      color: 'FF0000',
      lvlJc: 'right',
      start: 4,
    })
  })

  it('links a level to a style: w:pStyle on the level and w:numPr on the style', async () => {
    const h = await harness()
    const numId = createCustomListDef(h.ctx, MULTILEVEL_LIBRARY[4])!
    expect(h.styleUpserts[0]).toEqual({
      styleId: 'Heading1',
      pPr: { numPr: { numId, ilvl: 0 } },
    })
    expect(h.styleUpserts).toHaveLength(9)
    threeItems(h.editor, numId)
    const saved = await h.save()
    const xml = await numberingXml(saved.zip)
    expect(xml).toContain('<w:pStyle w:val="Heading1"/>')
    const styles = await saved.zip.file('word/styles.xml')!.async('string')
    expect(styles).toContain(`<w:numPr><w:ilvl w:val="1"/><w:numId w:val="${numId}"/></w:numPr>`)
    expect(styles.indexOf('w:styleId="Heading2"')).toBeGreaterThan(-1)
  })

  it('legal numbering and restart-after-level write w:isLgl and w:lvlRestart in CT_Lvl order', () => {
    const xml = customLevelXml(
      {
        numFmt: 'lowerLetter',
        lvlText: '%1.%2',
        indentLeft: 720,
        hanging: 360,
        start: 2,
        lvlRestart: 0,
        pStyle: 'Heading2',
        isLgl: true,
        suff: 'nothing',
        lvlJc: 'center',
      },
      1,
    )
    expect(xml).toBe(
      '<w:lvl w:ilvl="1"><w:start w:val="2"/><w:numFmt w:val="lowerLetter"/><w:lvlRestart w:val="0"/>' +
        '<w:pStyle w:val="Heading2"/><w:isLgl/><w:suff w:val="nothing"/><w:lvlText w:val="%1.%2"/>' +
        '<w:lvlJc w:val="center"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>',
    )
  })

  it('document presets group the definitions the document already uses', async () => {
    const h = await harness()
    createCustomListDef(h.ctx, bulletPresetLevels('◆'))
    createCustomListDef(h.ctx, MULTILEVEL_LIBRARY[1])
    const store = h.editor.storage.listNumbering.defs as Map<string, NumberingDef>
    const presets = documentListPresets(store.values())
    expect(presets.bullets.map((p) => p[0].lvlText)).toContain('◆')
    expect(presets.multi.some((p) => p[1].lvlText === '%1.%2.')).toBe(true)
    // the blank template's decimal list shows up once under numbers
    expect(presets.numbers.filter((p) => p[0].lvlText === '%1.')).toHaveLength(1)
  })

  it('document presets keep a w14 custom enumeration and tell two enumerations apart', async () => {
    const h = await harness()
    const czech = {
      numFmt: 'custom',
      customFormat: 'Za prvé, Za druhé, ...',
      lvlText: '%1:',
      indentLeft: 720,
    }
    const greek = { ...czech, customFormat: 'α, β, γ, ...' }
    createCustomListDef(h.ctx, [czech])
    createCustomListDef(h.ctx, [greek])
    const store = h.editor.storage.listNumbering.defs as Map<string, NumberingDef>
    const custom = documentListPresets(store.values()).numbers.filter(
      (p) => p[0].numFmt === 'custom',
    )
    expect(custom.map((p) => p[0].customFormat).sort()).toEqual([
      'Za prvé, Za druhé, ...',
      'α, β, γ, ...',
    ])
    expect(presetKey([czech])).not.toBe(presetKey([greek]))
  })
})

describe('List Styles', () => {
  const entry = (styles: ReturnType<typeof listStyleEntries>, id: string) =>
    styles.find((s) => s.styleId === id)!

  it("offers Word's four built-ins in a document that defines none, No List first", async () => {
    const h = await harness()
    const styles = listStyleEntries(h.parsed.listStyles)
    expect(styles.map((s) => s.styleId)).toEqual([
      NO_LIST_STYLE_ID,
      '111111',
      '1ai',
      'ArticleSection',
    ])
    expect(styles.slice(1).every((s) => !s.inDocument && s.levels.length === 9)).toBe(true)
    expect(listStylePreview(entry(styles, '111111').levels).slice(0, 3)).toEqual([
      '1.',
      '1.1.',
      '1.1.1.',
    ])
    expect(listStylePreview(entry(styles, '1ai').levels)).toEqual([
      '1)',
      'a)',
      'i)',
      '(1)',
      '(a)',
      '(i)',
      '1.',
      'a.',
      'i.',
    ])
    expect(listStylePreview(entry(styles, 'ArticleSection').levels).slice(0, 4)).toEqual([
      'Article I.',
      'Section 1.01',
      '(a)',
      '(i)',
    ])
    expect(listStyleNum(h.ctx, entry(styles, NO_LIST_STYLE_ID))).toBeNull()
  })

  it('applies a built-in: live markers, the style and its linked list in the save, the same list on reopen', async () => {
    const h = await harness()
    const style = entry(listStyleEntries(h.parsed.listStyles), '1ai')
    const list = listStyleNum(h.ctx, style)!
    expect(list.kind).toBe('ordered')
    threeItems(h.editor, list.numId)
    changeListLevel(h.editor, 1)
    expect(markers(h.editor)).toEqual(['1)', '2)', 'a)'])
    expect(listStylesInUse(h.editor, h.parsed.numbering)).toEqual({
      used: new Set(['1ai']),
      current: '1ai',
    })
    // a second application in the session runs the same list on
    expect(listStyleNum(h.ctx, style)).toEqual(list)
    expect(h.pending().styleNums).toHaveLength(1)

    const saved = await h.save()
    const xml = await numberingXml(saved.zip)
    const linking =
      /<w:abstractNum w:abstractNumId="(\d+)"><w:multiLevelType w:val="multilevel"\/><w:numStyleLink w:val="1ai"\/><\/w:abstractNum>/.exec(
        xml,
      )!
    expect(xml).toContain(
      `<w:num w:numId="${list.numId}"><w:abstractNumId w:val="${linking[1]}"/></w:num>`,
    )
    expect(xml).toMatch(/<w:styleLink w:val="1ai"\/><w:lvl w:ilvl="0">/)
    const stylesXml = await saved.zip.file('word/styles.xml')!.async('string')
    expect(stylesXml).toContain(
      '<w:style w:type="numbering" w:styleId="1ai"><w:name w:val="Outline List 1"/>',
    )
    const documentXml = await saved.zip.file('word/document.xml')!.async('string')
    expect(documentXml.match(new RegExp(`<w:numId w:val="${list.numId}"/>`, 'g'))).toHaveLength(3)
    expect(saved.blocks.filter((b) => b.list).map((b) => b.list)).toMatchObject([
      { numId: list.numId, ilvl: 0 },
      { numId: list.numId, ilvl: 0 },
      { numId: list.numId, ilvl: 1 },
    ])
    expect(saved.numbering.get(list.numId)!.listStyleId).toBe('1ai')

    const reopened = await harness(await saved.zip.generateAsync({ type: 'uint8array' }))
    const again = entry(listStyleEntries(reopened.parsed.listStyles), '1ai')
    expect(again.inDocument).toBe(true)
    expect(again.levels.slice(0, 3).map((l) => l.lvlText)).toEqual(['%1)', '%2)', '%3)'])
    expect(markers(reopened.editor)).toEqual(['1)', '2)', 'a)'])
    expect(listStylesInUse(reopened.editor, reopened.parsed.numbering).current).toBe('1ai')
    // the document has both the style and a list following it: nothing new is written
    expect(listStyleNum(reopened.ctx, again)).toEqual(list)
    expect(reopened.pending().styleNums).toEqual([])
  })

  it('Adjust List Indents on a list that follows a style changes the style: a new one in its definition, an existing one in its abstractNum', async () => {
    const h = await harness()
    const list = listStyleNum(h.ctx, entry(listStyleEntries(h.parsed.listStyles), '1ai'))!
    threeItems(h.editor, list.numId)
    expect(editListLevel(h.ctx, { indentLeft: 1440, hanging: 720 })).toBe(true)
    const saved = await h.save()
    const style = saved.listStyles!.find((s) => s.styleId === '1ai')!
    expect(style.levels[0]).toMatchObject({ lvlText: '%1)', indentLeft: 1440, hanging: 720 })
    expect(saved.numbering.get(list.numId)!.levels[0]).toMatchObject({ indentLeft: 1440 })
    expect(await numberingXml(saved.zip)).toMatch(/<w:numStyleLink w:val="1ai"\/><\/w:abstractNum>/)

    // reopened and stripped back to plain paragraphs, the style is the document's own
    const zip = saved.zip
    zip.file(
      'word/document.xml',
      (await zip.file('word/document.xml')!.async('string')).replace(
        /<w:numPr>.*?<\/w:numPr>/g,
        '',
      ),
    )
    const again = await harness(await zip.generateAsync({ type: 'uint8array' }))
    const fresh = listStyleNum(again.ctx, entry(listStyleEntries(again.parsed.listStyles), '1ai'))!
    threeItems(again.editor, fresh.numId)
    expect(editListLevel(again.ctx, { indentLeft: 2160, hanging: 720 })).toBe(true)
    const resaved = await again.save()
    expect(resaved.listStyles!.find((s) => s.styleId === '1ai')!.levels[0]).toMatchObject({
      indentLeft: 2160,
    })
    expect((await numberingXml(resaved.zip)).match(/<w:styleLink w:val="1ai"\/>/g)).toHaveLength(1)
  })

  it('an applied style that is undone before the save leaves no trace in the package', async () => {
    const h = await harness()
    listStyleNum(h.ctx, entry(listStyleEntries(h.parsed.listStyles), '111111'))
    const saved = await h.save()
    expect(saved.listStyles?.some((s) => s.styleId === '111111') ?? false).toBe(false)
  })

  it("lists the document's own list styles after the built-ins and follows their levels", async () => {
    const blank = await JSZip.loadAsync(await buildBlankDocx())
    const stylesXml = await blank.file('word/styles.xml')!.async('string')
    blank.file(
      'word/styles.xml',
      stylesXml.replace(
        '</w:styles>',
        '<w:style w:type="numbering" w:styleId="Clauses" w:customStyle="1"><w:name w:val="Clauses"/>' +
          '<w:pPr><w:numPr><w:numId w:val="40"/></w:numPr></w:pPr></w:style></w:styles>',
      ),
    )
    const numbering = await blank.file('word/numbering.xml')!.async('string')
    blank.file(
      'word/numbering.xml',
      numbering
        .replace(
          /<w:num /,
          '<w:abstractNum w:abstractNumId="40"><w:multiLevelType w:val="multilevel"/><w:styleLink w:val="Clauses"/>' +
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1:"/></w:lvl>' +
            '</w:abstractNum><w:num ',
        )
        .replace(
          '</w:numbering>',
          '<w:num w:numId="40"><w:abstractNumId w:val="40"/></w:num></w:numbering>',
        ),
    )
    const h = await harness(await blank.generateAsync({ type: 'uint8array' }))
    const styles = listStyleEntries(h.parsed.listStyles)
    expect(styles.map((s) => s.styleId)).toEqual([
      NO_LIST_STYLE_ID,
      '111111',
      '1ai',
      'ArticleSection',
      'Clauses',
    ])
    const list = listStyleNum(h.ctx, entry(styles, 'Clauses'))!
    expect(list.numId).not.toBe('40')
    expect(h.pending().styleNums).toEqual([{ numId: list.numId, styleId: 'Clauses' }])
    threeItems(h.editor, list.numId)
    expect(markers(h.editor)).toEqual(['A:', 'B:', 'C:'])
    const saved = await h.save()
    const xml = await numberingXml(saved.zip)
    expect(xml.match(/<w:styleLink w:val="Clauses"\/>/g)).toHaveLength(1)
    expect(xml.match(/<w:numStyleLink w:val="Clauses"\/>/g)).toHaveLength(1)
    expect(saved.numbering.get(list.numId)!.levels[0]).toMatchObject({ numFmt: 'upperLetter' })
  })

  const shown = (h: Harness) =>
    listStyleEntries(h.parsed.listStyles, h.pending().styleDefs, h.pending().removedStyles)

  it('Add: a named style is saved unused, is listed on reopen, and its name cannot be taken twice', async () => {
    const h = await harness()
    const styleId = addListStyle(h.ctx, 'My Steps', MULTILEVEL_LIBRARY[0], shown(h))!
    expect(styleId).toBe('MySteps')
    expect(addListStyle(h.ctx, ' my steps ', MULTILEVEL_LIBRARY[0], shown(h))).toBeNull()
    expect(addListStyle(h.ctx, 'Normal', MULTILEVEL_LIBRARY[0], shown(h))).toBeNull()
    expect(shown(h).map((s) => s.styleId)).toContain('MySteps')
    expect(entry(shown(h), 'MySteps')).toMatchObject({ custom: true, inDocument: false })

    const unused = await h.save()
    expect(unused.listStyles!.find((s) => s.styleId === 'MySteps')).toMatchObject({
      name: 'My Steps',
      custom: true,
    })

    // applied, the list defines the style itself: it is still written once
    const list = listStyleNum(h.ctx, entry(shown(h), 'MySteps'))!
    threeItems(h.editor, list.numId)
    expect(markers(h.editor)).toEqual(['1)', '2)', '3)'])
    const saved = await h.save()
    const xml = await numberingXml(saved.zip)
    expect(xml.match(/<w:styleLink w:val="MySteps"\/>/g)).toHaveLength(1)
    expect(saved.numbering.get(list.numId)!.listStyleId).toBe('MySteps')

    const reopened = await harness(await saved.zip.generateAsync({ type: 'uint8array' }))
    expect(entry(shown(reopened), 'MySteps')).toMatchObject({
      label: 'My Steps',
      custom: true,
      inDocument: true,
    })
    expect(markers(reopened.editor)).toEqual(['1)', '2)', '3)'])
    expect(listStylesInUse(reopened.editor, reopened.parsed.numbering).current).toBe('MySteps')
  })

  it('Delete: the style leaves the document, its paragraphs turn plain, and an undo leaves an ordinary list', async () => {
    const first = await harness()
    addListStyle(first.ctx, 'My Steps', MULTILEVEL_LIBRARY[0], shown(first))
    const list = listStyleNum(first.ctx, entry(shown(first), 'MySteps'))!
    threeItems(first.editor, list.numId)
    const saved = await first.save()

    const h = await harness(await saved.zip.generateAsync({ type: 'uint8array' }))
    deleteListStyle(h.ctx, entry(shown(h), '1ai'))
    expect(h.pending().removedStyles).toEqual([])
    deleteListStyle(h.ctx, entry(shown(h), 'MySteps'))
    expect(shown(h).map((s) => s.styleId)).not.toContain('MySteps')
    expect(h.editor.state.doc.childCount).toBe(3)
    h.editor.state.doc.forEach((node) => expect(node.type.name).toBe('docParagraph'))
    expect(h.editor.state.doc.textContent).toBe('onetwothree')

    const plain = await h.save()
    expect(plain.listStyles?.some((s) => s.styleId === 'MySteps') ?? false).toBe(false)
    expect(await plain.zip.file('word/styles.xml')!.async('string')).not.toContain('MySteps')
    expect(await numberingXml(plain.zip)).not.toContain('tyleLink w:val="MySteps"')
    expect(plain.blocks.some((b) => b.list)).toBe(false)

    h.editor.commands.undo()
    expect(markers(h.editor)).toEqual(['1)', '2)', '3)'])
    expect(listStylesInUse(h.editor, h.parsed.numbering).used.size).toBe(0)
    const listed = await h.save()
    expect(listed.listStyles?.some((s) => s.styleId === 'MySteps') ?? false).toBe(false)
    expect(listed.numbering.get(list.numId)!.listStyleId).toBeUndefined()
    expect(listed.numbering.get(list.numId)!.levels[1]).toMatchObject({ numFmt: 'lowerLetter' })
    expect(listed.blocks.filter((b) => b.list)).toHaveLength(3)
  })

  it('Delete of a style added in this session writes nothing of it', async () => {
    const h = await harness()
    addListStyle(h.ctx, 'Draft', MULTILEVEL_LIBRARY[0], shown(h))
    const list = listStyleNum(h.ctx, entry(shown(h), 'Draft'))!
    threeItems(h.editor, list.numId)
    editListLevel(h.ctx, { indentLeft: 1234 })
    deleteListStyle(h.ctx, entry(shown(h), 'Draft'))
    expect(h.pending()).toMatchObject({ styleDefs: [], styleNums: [], removedStyles: [] })
    const saved = await h.save()
    expect(await saved.zip.file('word/styles.xml')!.async('string')).not.toContain('Draft')
    expect(saved.blocks.some((b) => b.list)).toBe(false)
    // undone, the list is an ordinary one that keeps the level as it was last edited
    h.editor.commands.undo()
    const listed = await h.save()
    expect(listed.numbering.get(list.numId)!.levels[0]).toMatchObject({ indentLeft: 1234 })
  })
})

describe('Modify Style on a list style', () => {
  /** the List Styles tab as the app shows it, edits of this session included */
  const shown = (h: Harness) =>
    listStyleEntries(
      h.parsed.listStyles,
      h.pending().styleDefs,
      h.pending().removedStyles,
      listStyleEdits(
        h.parsed,
        h.pending(),
        new Map(h.styleUpserts.flatMap((u) => (u.name ? [[u.styleId, u.name] as const] : []))),
      ),
    )
  const entry = (styles: ReturnType<typeof shown>, styleId: string) =>
    styles.find((s) => s.styleId === styleId)!
  const withLevel = (
    style: ReturnType<typeof entry>,
    ilvl: number,
    patch: Partial<(typeof style.levels)[number]>,
  ) => nineListLevels(style.levels).map((l, i) => (i === ilvl ? { ...l, ...patch } : l))

  async function documentWithClauses(): Promise<Uint8Array> {
    const blank = await JSZip.loadAsync(await buildBlankDocx())
    blank.file(
      'word/styles.xml',
      (await blank.file('word/styles.xml')!.async('string')).replace(
        '</w:styles>',
        '<w:style w:type="numbering" w:styleId="Clauses" w:customStyle="1"><w:name w:val="Clauses"/>' +
          '<w:pPr><w:numPr><w:numId w:val="40"/></w:numPr></w:pPr></w:style></w:styles>',
      ),
    )
    blank.file(
      'word/numbering.xml',
      (await blank.file('word/numbering.xml')!.async('string'))
        .replace(
          /<w:num /,
          '<w:abstractNum w:abstractNumId="40"><w:multiLevelType w:val="multilevel"/><w:styleLink w:val="Clauses"/>' +
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1:"/></w:lvl>' +
            '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2)"/></w:lvl>' +
            '</w:abstractNum>' +
            '<w:abstractNum w:abstractNumId="41"><w:multiLevelType w:val="multilevel"/><w:numStyleLink w:val="Clauses"/></w:abstractNum>' +
            '<w:num ',
        )
        .replace(
          '</w:numbering>',
          '<w:num w:numId="40"><w:abstractNumId w:val="40"/></w:num>' +
            '<w:num w:numId="41"><w:abstractNumId w:val="41"/></w:num></w:numbering>',
        ),
    )
    return blank.generateAsync({ type: 'uint8array' })
  }

  it("rewrites the style's own abstractNum: every list following it changes, on the page, in the file and on reopen", async () => {
    const h = await harness(await documentWithClauses())
    threeItems(h.editor, '41')
    expect(markers(h.editor)).toEqual(['A:', 'B:', 'C:'])
    const style = entry(shown(h), 'Clauses')
    expect(style.levels.map((l) => l.lvlText)).toEqual(['%1:', '%2)'])
    expect(
      modifyListStyle(
        h.ctx,
        style,
        'Terms',
        withLevel(style, 0, { numFmt: 'lowerRoman', lvlText: '(%1)', start: 3 }),
      ),
    ).toBe(true)
    expect(markers(h.editor)).toEqual(['(iii)', '(iv)', '(v)'])
    expect(h.pending().levelEdits).toEqual([
      { abstractNumId: '40', ilvl: 0, level: expect.objectContaining({ lvlText: '(%1)' }) },
    ])
    // the tab shows the edit before any save
    expect(entry(shown(h), 'Clauses')).toMatchObject({ name: 'Terms', label: 'Terms' })
    expect(entry(shown(h), 'Clauses').levels[0]).toMatchObject({ numFmt: 'lowerRoman', start: 3 })

    const saved = await h.save()
    const xml = await numberingXml(saved.zip)
    const abstract = /<w:abstractNum w:abstractNumId="40">[\s\S]*?<\/w:abstractNum>/.exec(xml)![0]
    expect(abstract).toContain('<w:styleLink w:val="Clauses"/>')
    expect(abstract).toMatch(
      /<w:lvl w:ilvl="0"[^>]*><w:start w:val="3"\/><w:numFmt w:val="lowerRoman"\/>/,
    )
    expect(abstract).toContain('<w:lvlText w:val="(%1)"/>')
    expect(abstract).toContain('<w:lvlText w:val="%2)"/>')
    expect(abstract.match(/<w:lvl /g)).toHaveLength(2)
    expect(xml.match(/<w:numStyleLink w:val="Clauses"\/>/g)).toHaveLength(1)
    expect(await saved.zip.file('word/styles.xml')!.async('string')).toMatch(
      /w:styleId="Clauses"[^>]*><w:name w:val="Terms"\/>/,
    )

    const reopened = await harness(await saved.zip.generateAsync({ type: 'uint8array' }))
    expect(entry(shown(reopened), 'Clauses')).toMatchObject({ name: 'Terms', custom: true })
    expect(markers(reopened.editor)).toEqual(['(iii)', '(iv)', '(v)'])
    expect(saved.numbering.get('41')!.listStyleId).toBe('Clauses')
  })

  it('legal numbering on a style of two levels rewrites those two and adds none', async () => {
    const h = await harness(await documentWithClauses())
    const style = entry(shown(h), 'Clauses')
    const legal = nineListLevels(style.levels).map((l) => ({ ...l, isLgl: true }))
    expect(modifyListStyle(h.ctx, style, style.name, legal)).toBe(true)
    expect(h.pending().levelEdits.map((e) => e.ilvl)).toEqual([0, 1])
    const saved = await h.save()
    const xml = await numberingXml(saved.zip)
    const abstract = /<w:abstractNum w:abstractNumId="40">[\s\S]*?<\/w:abstractNum>/.exec(xml)![0]
    expect(abstract.match(/<w:lvl /g)).toHaveLength(2)
    expect(abstract.match(/<w:isLgl\/>/g)).toHaveLength(2)
  })

  it('Modify Style is one undo step, in order with text edits, and an undone one leaves the file as it was', async () => {
    const h = await harness(await documentWithClauses())
    threeItems(h.editor, '41')
    h.editor.on('transaction', ({ transaction }) => {
      for (const { change, side } of replayedListStyles(transaction)) {
        h.ctx.setPendingNumbering((p) => withListStylePending(p, change, change[side].pending))
        queueMicrotask(() => drawListStyleSide(h.editor, change[side]))
      }
    })
    const settled = () => new Promise<void>((done) => queueMicrotask(done))
    const text = () => h.editor.state.doc.lastChild!.textContent
    const type = (text: string) =>
      h.editor.commands.insertContentAt(h.editor.state.doc.content.size - 1, text)
    const modify = (lvlText: string) =>
      modifyListStyleAsStep(
        h.ctx,
        entry(shown(h), 'Clauses'),
        'Clauses',
        withLevel(entry(shown(h), 'Clauses'), 0, { lvlText }),
        () => undefined,
      )

    type(' a')
    expect(modify('(%1)')).toBe(true)
    type(' b')
    expect(modify('[%1]')).toBe(true)
    expect(markers(h.editor)).toEqual(['[A]', '[B]', '[C]'])
    // OK without a change is not a step
    expect(modify('[%1]')).toBe(true)

    h.editor.commands.undo()
    await settled()
    expect(markers(h.editor)).toEqual(['(A)', '(B)', '(C)'])
    expect(h.pending().levelEdits).toEqual([
      { abstractNumId: '40', ilvl: 0, level: expect.objectContaining({ lvlText: '(%1)' }) },
    ])
    expect(text()).toBe('three a b')
    h.editor.commands.undo()
    expect(text()).toBe('three a')
    h.editor.commands.undo()
    await settled()
    expect(markers(h.editor)).toEqual(['A:', 'B:', 'C:'])
    expect(h.pending()).toEqual(EMPTY_PENDING_NUMBERING)
    expect(entry(shown(h), 'Clauses').levels[0].lvlText).toBe('%1:')
    const untouched = await numberingXml((await h.save()).zip)
    expect(untouched).toBe(await numberingXml(await JSZip.loadAsync(await documentWithClauses())))

    h.editor.commands.redo()
    h.editor.commands.redo()
    h.editor.commands.redo()
    await settled()
    expect(markers(h.editor)).toEqual(['[A]', '[B]', '[C]'])
    expect(text()).toBe('three a b')
    const saved = await h.save()
    expect(await numberingXml(saved.zip)).toContain('<w:lvlText w:val="[%1]"/>')

    // after a save the file holds the new levels: the undo side is what the style is modified back to
    const reopened = await harness(await saved.zip.generateAsync({ type: 'uint8array' }))
    threeItems(reopened.editor, '41')
    expect(markers(reopened.editor)).toEqual(['[A]', '[B]', '[C]'])
    let undo: ListStyleSide | undefined
    h.editor.on('transaction', ({ transaction }) => {
      undo = replayedListStyles(transaction)[0]?.change.undo ?? undo
    })
    h.editor.commands.undo()
    modifyListStyle(reopened.ctx, entry(shown(reopened), 'Clauses'), undo!.name, undo!.levels)
    expect(markers(reopened.editor)).toEqual(['(A)', '(B)', '(C)'])
  })

  it("Modify Style on one of Word's own styles the document lacks undoes to no definition at all", async () => {
    const h = await harness()
    h.editor.on('transaction', ({ transaction }) => {
      for (const { change, side } of replayedListStyles(transaction))
        h.ctx.setPendingNumbering((p) => withListStylePending(p, change, change[side].pending))
    })
    const style = entry(shown(h), '1ai')
    expect(
      modifyListStyleAsStep(
        h.ctx,
        style,
        style.name,
        withLevel(style, 0, { lvlText: '%1 -' }),
        () => undefined,
      ),
    ).toBe(true)
    expect(h.pending().styleDefs).toHaveLength(1)
    h.editor.commands.undo()
    expect(h.pending()).toEqual(EMPTY_PENDING_NUMBERING)
    h.editor.commands.redo()
    expect(h.pending().styleDefs[0]).toMatchObject({ styleId: '1ai', builtIn: true })
    expect(entry(shown(h), '1ai').levels[0].lvlText).toBe('%1 -')
  })

  it('OK without a change writes nothing', async () => {
    const h = await harness()
    const article = entry(shown(h), 'ArticleSection')
    expect(modifyListStyle(h.ctx, article, article.name, nineListLevels(article.levels))).toBe(true)
    expect(h.pending()).toEqual(EMPTY_PENDING_NUMBERING)
  })

  it("one of Word's own styles the document lacks is saved with the new levels, not as a custom style, and a list applied later follows them", async () => {
    const h = await harness()
    const style = entry(shown(h), '1ai')
    modifyListStyle(h.ctx, style, 'ignored', withLevel(style, 0, { lvlText: '%1 -' }))
    expect(entry(shown(h), '1ai')).toMatchObject({ name: 'Outline List 1', inDocument: false })
    expect(entry(shown(h), '1ai').levels[0].lvlText).toBe('%1 -')
    expect(shown(h).map((s) => s.styleId)).toEqual([
      NO_LIST_STYLE_ID,
      '111111',
      '1ai',
      'ArticleSection',
    ])

    const unused = await h.save()
    expect(await unused.zip.file('word/styles.xml')!.async('string')).toContain(
      '<w:style w:type="numbering" w:styleId="1ai"><w:name w:val="Outline List 1"/>',
    )
    expect(unused.listStyles!.find((s) => s.styleId === '1ai')!.levels[0].lvlText).toBe('%1 -')

    const list = listStyleNum(h.ctx, entry(shown(h), '1ai'))!
    threeItems(h.editor, list.numId)
    expect(markers(h.editor)).toEqual(['1 -', '2 -', '3 -'])
    const saved = await h.save()
    expect((await numberingXml(saved.zip)).match(/<w:styleLink w:val="1ai"\/>/g)).toHaveLength(1)
    const reopened = await harness(await saved.zip.generateAsync({ type: 'uint8array' }))
    expect(markers(reopened.editor)).toEqual(['1 -', '2 -', '3 -'])
  })

  it('a style applied or added in this session is modified where it is still pending', async () => {
    const h = await harness()
    const list = listStyleNum(h.ctx, entry(shown(h), '1ai'))!
    threeItems(h.editor, list.numId)
    const applied = entry(shown(h), '1ai')
    modifyListStyle(h.ctx, applied, applied.name, withLevel(applied, 0, { numFmt: 'upperLetter' }))
    expect(markers(h.editor)).toEqual(['A)', 'B)', 'C)'])
    expect(h.pending().styleDefs).toEqual([])
    expect(entry(shown(h), '1ai').levels[0].numFmt).toBe('upperLetter')

    const styleId = addListStyle(h.ctx, 'My Steps', MULTILEVEL_LIBRARY[0], shown(h))!
    const mine = entry(shown(h), styleId)
    modifyListStyle(h.ctx, mine, 'Our Steps', withLevel(mine, 1, { lvlText: '[%2]' }))
    expect(entry(shown(h), styleId)).toMatchObject({ name: 'Our Steps', label: 'Our Steps' })

    const saved = await h.save()
    expect(saved.numbering.get(list.numId)!.levels[0]).toMatchObject({ numFmt: 'upperLetter' })
    const steps = saved.listStyles!.find((s) => s.styleId === styleId)!
    expect(steps).toMatchObject({ name: 'Our Steps', custom: true })
    expect(steps.levels[1].lvlText).toBe('[%2]')
  })

  it('a style the document names without levels has nothing to modify', async () => {
    const blank = await JSZip.loadAsync(await buildBlankDocx())
    blank.file(
      'word/styles.xml',
      (await blank.file('word/styles.xml')!.async('string')).replace(
        '</w:styles>',
        '<w:style w:type="numbering" w:styleId="1ai"><w:name w:val="Outline List 1"/></w:style></w:styles>',
      ),
    )
    const h = await harness(await blank.generateAsync({ type: 'uint8array' }))
    const style = entry(shown(h), '1ai')
    expect(style.numberless).toBe(true)
    expect(modifyListStyle(h.ctx, style, style.name, withLevel(style, 0, { start: 5 }))).toBe(false)
    expect(h.pending()).toEqual(EMPTY_PENDING_NUMBERING)
  })
})

describe('Change List Level and paragraph spacing', () => {
  it('moves the caret item to the level and clamps to 0-8', async () => {
    const h = await harness()
    const numId = allocateListNumId(h.ctx, 'ordered')!
    threeItems(h.editor, numId)
    expect(changeListLevel(h.editor, 2)).toBe(true)
    expect(h.editor.state.doc.child(2).attrs.ilvl).toBe(2)
    expect(h.editor.state.doc.child(1).attrs.ilvl).toBe(0)
    changeListLevel(h.editor, 42)
    expect(h.editor.state.doc.child(2).attrs.ilvl).toBe(8)
    h.editor.commands.setTextSelection(1)
    h.editor.commands.setNode('docParagraph')
    expect(changeListLevel(h.editor, 1)).toBe(false)
  })

  it('adds 12 pt before/after when there is none and removes it when there is', async () => {
    const h = await harness()
    setDoc(h.editor, [{ type: 'docParagraph', content: [{ type: 'text', text: 'body' }] }])
    toggleParaSpace(h.editor, 'spaceBefore', 0)
    expect(h.editor.state.doc.child(0).attrs.spaceBefore).toBe(240)
    toggleParaSpace(h.editor, 'spaceAfter', 0)
    expect(h.editor.state.doc.child(0).attrs.spaceAfter).toBe(240)
    toggleParaSpace(h.editor, 'spaceBefore', 240)
    expect(h.editor.state.doc.child(0).attrs.spaceBefore).toBe(0)
    const saved = await h.save()
    const doc = await saved.zip.file('word/document.xml')!.async('string')
    expect(doc).toMatch(/<w:spacing [^>]*w:after="240"/)
  })
})
