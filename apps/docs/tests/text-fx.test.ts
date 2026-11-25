/** Text Effects: an outline, a shadow, a reflection and a glow go on the selected text, show, and reach the save as the text's own. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  REFLECTION_PRESETS,
  SHADOW_PRESETS,
  TEXT_FX_PRESETS,
  applyTextFx,
  applyTextFxPreset,
  glowOf,
  gradientFillOf,
  gradientSpreadOf,
  sameTextFx,
  textFillOf,
  textFillShownOf,
  textFxDecls,
  textFxOf,
} from '../src/renderer/editor/text-fx'

const NONE = { outline: null, shadow: null, glow: null, reflection: null }
const FX = {
  outline: { color: 'FF0000', widthPt: 1.5 },
  shadow: SHADOW_PRESETS[0].shadow,
  glow: glowOf('4472C4', 8),
  reflection: null,
}

describe('text effects', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'plain fancy text' }] }],
      },
    })
  })
  afterEach(() => editor.destroy())

  const savedRuns = () => {
    const first = pmDocToSavePlan(editor.getJSON() as PmNode, []).saveBlocks[0]
    return first.kind === 'generated' ? (first.block.runs ?? []) : []
  }

  it('puts the effects on the selection only, and reads them back from it', () => {
    editor.commands.setTextSelection({ from: 7, to: 12 })
    expect(textFxOf(editor)).toEqual(NONE)
    applyTextFx(editor, FX)
    expect(sameTextFx(textFxOf(editor), FX)).toBe(true)
    const runs = savedRuns()
    expect(runs.map((r) => r.text)).toEqual(['plain ', 'fancy', ' text'])
    expect(runs[1]).toMatchObject({
      textOutline: FX.outline,
      textShadow: FX.shadow,
      glow: FX.glow,
      fxOwn: FX,
    })
    expect(runs[0].fxOwn).toBeUndefined()
  })

  it('draws the shadow and the glow as one text-shadow, the outline as a stroke', () => {
    expect(textFxDecls(NONE)).toEqual([])
    const decls = textFxDecls(FX)
    expect(decls[0]).toBe('-webkit-text-stroke:1.5pt #FF0000')
    // the shadow thrown down and to the right, then the two layers of the halo
    expect(decls[1]).toBe(
      'text-shadow:2.12pt 2.12pt 4pt rgba(0,0,0,0.6),' +
        '0 0 4pt rgba(68,114,196,0.6),0 0 8pt rgba(68,114,196,0.6)',
    )
    editor.commands.setTextSelection({ from: 7, to: 12 })
    applyTextFx(editor, FX)
    const span = editor.view.dom.querySelector('span[style*="text-shadow"]') as HTMLElement
    expect(span.textContent).toBe('fancy')
    expect(span.style.textShadow).toContain('rgba(0,0,0,0.6)')
  })

  it('clears an effect as the text’s own, so the save takes it off', () => {
    editor.commands.setTextSelection({ from: 7, to: 12 })
    applyTextFx(editor, FX)
    applyTextFx(editor, { ...FX, glow: null })
    const run = savedRuns()[1]
    expect(run.glow).toBeUndefined()
    expect(run.fxOwn).toEqual({
      outline: FX.outline,
      shadow: FX.shadow,
      glow: null,
      reflection: null,
    })
  })

  it('mirrors the text under itself, fading out where the reflection ends', () => {
    const half = REFLECTION_PRESETS[1].reflection
    expect(textFxDecls({ ...NONE, reflection: half })).toEqual([
      '-webkit-box-reflect:below 0pt linear-gradient(transparent 45%,rgba(0,0,0,0.5))',
    ])
    editor.commands.setTextSelection({ from: 7, to: 12 })
    applyTextFx(editor, { ...NONE, reflection: half })
    expect(textFxOf(editor).reflection).toEqual(half)
    const run = savedRuns()[1]
    expect(run.textReflection).toEqual(half)
    expect(run.fxOwn?.reflection).toEqual(half)
    applyTextFx(editor, NONE)
    expect(savedRuns()[1].textReflection).toBeUndefined()
    expect(savedRuns()[1].fxOwn?.reflection).toBeNull()
  })

  it('applies a preset as its text colour with its effects, and nothing else of the run', () => {
    const preset = TEXT_FX_PRESETS.find((item) => item.id === 'white-outline-glow')!
    editor.commands.setTextSelection({ from: 7, to: 12 })
    editor.commands.setMark('docTextStyle', { sizeHalfPoints: 40 })
    applyTextFx(editor, preset.fx, preset.color)
    const run = savedRuns()[1]
    expect(run).toMatchObject({
      text: 'fancy',
      color: 'FFFFFF',
      sizeHalfPoints: 40,
      textOutline: preset.fx.outline,
      glow: preset.fx.glow,
    })
    // the colour is a text fill, as Word saves a preset
    expect(run.fxOwn?.fill).toBe('FFFFFF')
    expect(savedRuns()[0].color).toBeUndefined()
    expect(savedRuns()[0].fxOwn).toBeUndefined()
    // every preset is told from the others by what it looks like
    const looks = TEXT_FX_PRESETS.map((item) => `${item.color} ${textFxDecls(item.fx).join(';')}`)
    expect(new Set(looks).size).toBe(TEXT_FX_PRESETS.length)
  })

  it('shows the effects of a file and does not claim them as edits', async () => {
    const body =
      '<w:p><w:r><w:rPr><w14:shadow w14:blurRad="38100" w14:dist="25400" w14:dir="5400000">' +
      '<w14:srgbClr w14:val="112233"/></w14:shadow></w:rPr><w:t>shaded</w:t></w:r></w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const doc = blocksToPmDoc(parsed.blocks) as PmNode
    const mark = doc.content?.[0]?.content?.[0]?.marks?.find((m) => m.type === 'docTextStyle')
    expect(JSON.parse(String(mark?.attrs?.textShadow))).toMatchObject({
      color: '112233',
      dirDeg: 90,
    })
    expect(mark?.attrs?.fxOwn ?? null).toBeNull()
  })

  it('shows the reflection of a file and saves the run with it untouched', async () => {
    const body =
      '<w:p><w:r><w:rPr><w14:reflection w14:blurRad="6350" w14:stA="52000" w14:stPos="0" ' +
      'w14:endA="300" w14:endPos="35000" w14:dist="0" w14:dir="5400000" w14:fadeDir="5400000" ' +
      'w14:sx="100000" w14:sy="-100000" w14:kx="0" w14:ky="0" w14:algn="bl"/></w:rPr>' +
      '<w:t>mirrored</w:t></w:r></w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const doc = blocksToPmDoc(parsed.blocks) as PmNode
    const mark = doc.content?.[0]?.content?.[0]?.marks?.find((m) => m.type === 'docTextStyle')
    expect(JSON.parse(String(mark?.attrs?.textReflection))).toEqual({
      startAlpha: 0.52,
      endPos: 0.35,
      blurPt: 0.5,
      distPt: 0,
    })
    expect(mark?.attrs?.fxOwn ?? null).toBeNull()
    // unchanged, the paragraph is saved from the file's own bytes
    expect(pmDocToSavePlan(doc, parsed.blocks).saveBlocks[0].kind).toBe('original')
  })

  it('applies a look by its id, and keeps its fill through a later change of the effects', () => {
    editor.commands.setTextSelection({ from: 7, to: 12 })
    expect(textFillOf(editor)).toBeNull()
    expect(applyTextFxPreset(editor, 'orange-outline')).toBe(true)
    expect(textFillOf(editor)).toBe('ED7D31')
    expect(savedRuns()[1]).toMatchObject({
      color: 'ED7D31',
      textOutline: { color: 'ED7D31', widthPt: 0.75 },
      fxOwn: { fill: 'ED7D31' },
    })
    applyTextFx(editor, { ...NONE, glow: glowOf('4472C4', 5) })
    expect(savedRuns()[1]).toMatchObject({
      color: 'ED7D31',
      fxOwn: { fill: 'ED7D31', outline: null },
    })
    expect(applyTextFxPreset(editor, 'no-such-look')).toBe(false)
    expect(applyTextFxPreset(editor, undefined)).toBe(false)
  })

  it('leaves text without a fill without one when only effects are set', () => {
    editor.commands.setTextSelection({ from: 7, to: 12 })
    applyTextFx(editor, FX)
    expect(savedRuns()[1].fxOwn).not.toHaveProperty('fill')
  })

  it('puts a gradient over the colour, draws it clipped to the glyphs, and saves it as the text’s own', () => {
    const gradient = gradientFillOf('4472C4', 'down')
    expect(gradient).toEqual({
      kind: 'gradient',
      stops: [
        { pos: 0, color: 'B4C7E7' },
        { pos: 1, color: '4472C4' },
      ],
      angleDeg: 90,
    })
    expect(gradientSpreadOf(gradient)).toBe('down')
    expect(gradientSpreadOf(gradientFillOf('C00000', 'center'))).toBe('center')
    editor.commands.setTextSelection({ from: 7, to: 12 })
    applyTextFx(editor, NONE, '4472C4', gradient)
    expect(textFillShownOf(editor)).toEqual(gradient)
    const span = editor.view.dom.querySelector('span[style*="background-clip"]') as HTMLElement
    expect(span.textContent).toBe('fancy')
    expect(span.style.backgroundImage).toMatch(/^linear-gradient\(180deg, rgb\(180, 199, 231\) 0%/)
    expect(span.getAttribute('style')).toMatch(/-webkit-text-fill-color: ?transparent/)
    const run = savedRuns()[1]
    expect(run.color).toBe('4472C4')
    expect(run.textFill).toEqual(gradient)
    expect(run.fxOwn).toMatchObject({ fill: '4472C4', textFill: gradient })
  })

  it('gives a gradient up for a colour picked over it, on the page and in the save', () => {
    editor.commands.setTextSelection({ from: 7, to: 12 })
    applyTextFx(editor, NONE, '4472C4', gradientFillOf('4472C4', 'right'))
    editor.chain().setMark('docTextStyle', { color: 'C00000' }).run()
    expect(textFillShownOf(editor)).toBeNull()
    expect(editor.view.dom.querySelector('span[style*="background-clip"]')).toBeNull()
    const run = savedRuns()[1]
    expect(run.textFill).toBeUndefined()
    expect(run.fxOwn).toMatchObject({ fill: 'C00000', textFill: null })
  })

  it('makes a solid fill see-through over its colour, and takes that off again', () => {
    editor.commands.setTextSelection({ from: 7, to: 12 })
    const seeThrough = { kind: 'solid' as const, color: '4472C4', alpha: 0.6 }
    applyTextFx(editor, NONE, '4472C4', seeThrough)
    expect(textFillShownOf(editor)).toEqual(seeThrough)
    const span = editor.view.dom.querySelector('span[style*="text-fill-color"]') as HTMLElement
    expect(span.textContent).toBe('fancy')
    expect(span.getAttribute('style')).toMatch(/rgba\(68, ?114, ?196, ?0\.6\)/)
    const run = savedRuns()[1]
    expect(run.color).toBe('4472C4')
    expect(run.fxOwn).toMatchObject({ fill: '4472C4', textFill: seeThrough })

    applyTextFx(editor, NONE, undefined, null)
    expect(textFillShownOf(editor)).toBeNull()
    expect(savedRuns()[1].fxOwn).toMatchObject({ fill: '4472C4', textFill: null })
  })

  it('takes a gradient off and keeps it through a change of another effect', () => {
    const gradient = gradientFillOf('4472C4', 'diagonal')
    editor.commands.setTextSelection({ from: 7, to: 12 })
    applyTextFx(editor, NONE, '4472C4', gradient)
    applyTextFx(editor, { ...NONE, glow: glowOf('4472C4', 5) })
    expect(savedRuns()[1].fxOwn).toMatchObject({ textFill: gradient })
    expect(textFillShownOf(editor)).toEqual(gradient)
    applyTextFx(editor, NONE, undefined, null)
    expect(textFillShownOf(editor)).toBeNull()
    expect(savedRuns()[1].fxOwn).toMatchObject({ fill: '4472C4', textFill: null })
  })
})

describe('a text fill from a file', () => {
  const FILL =
    '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="C00000"/></w14:solidFill></w14:textFill>'
  const body = (rPr: string) =>
    '<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>' +
    `<w:r><w:rPr>${rPr}</w:rPr><w:t>filled</w:t></w:r>` +
    '<w:r><w:t xml:space="preserve"> word</w:t></w:r></w:p>'

  async function edited(rPr: string, edit: (editor: Editor) => void): Promise<string> {
    const parsed = await parseDocx(await buildDocx({ bodyXml: body(rPr) }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editor.commands.setTextSelection({ from: 3, to: 9 })
    expect(textFillOf(editor)).toBe('C00000')
    edit(editor)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    editor.destroy()
    const bytes = await saveDocx(parsed, plan.saveBlocks)
    return (await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) ?? ''
  }

  it('is the colour the text shows, and stays as written when an effect is added', async () => {
    const xml = await edited(FILL, (editor) =>
      applyTextFx(editor, { ...NONE, glow: glowOf('4472C4', 5) }),
    )
    expect(xml).toContain(`</w14:glow>${FILL}</w:rPr>`)
    expect(xml).not.toContain('<w:color')
  })

  it('takes the colour the text is given afterwards', async () => {
    const xml = await edited(`<w:color w:val="C00000"/>${FILL}`, (editor) =>
      editor.commands.setMark('docTextStyle', { color: '0070C0' }),
    )
    expect(xml).toContain(
      '<w:color w:val="0070C0"/><w14:textFill><w14:solidFill><w14:srgbClr w14:val="0070C0"/>',
    )
    expect(xml).not.toContain('C00000')
  })
})
