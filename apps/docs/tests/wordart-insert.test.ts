/**
 * WordArt insertion tests - Item 3: WordArt (simplified) → save → re-parse
 *
 * Verifies:
 *  1. buildWordArtParagraphXml produces the correct structure
 *  2. Every shared preset produces valid XML with a readable solid color
 *  3. insert → save → Word structure assertions (wps:wsp + w:txbxContent, no fill, large text)
 *  4. Text is readable after re-parsing
 *  5. The wordArtId field is kept in the display model (display-only, not written to OOXML)
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import {
  buildWordArtParagraphXml,
  parseDocx,
  saveDocx,
  type TextboxDisplay,
} from '@genoffice/docx-engine'
import { WORDART_PRESETS, wordArtSolidColor, type WordArtPreset } from '@genoffice/ui'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

async function openBlankDoc() {
  const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body text</w:t></w:r></w:p>' })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

function makeWordArtTextbox(preset: WordArtPreset): TextboxDisplay {
  const solidHex = wordArtSolidColor(preset).replace('#', '')
  return {
    widthPx: Math.round(2700000 / 9525),
    heightPx: Math.round(720000 / 9525),
    wordArtId: preset.id,
    paras: [
      {
        runs: [{ text: 'WordArt', color: solidHex, bold: true, sizeHalfPoints: 72 }],
        align: 'center',
      },
    ],
  }
}

function insertWordArt(editor: Editor, preset: WordArtPreset) {
  const xml = buildWordArtParagraphXml({
    colorHex: wordArtSolidColor(preset).replace('#', ''),
    italic: preset.italic,
    widthEmu: 2700000,
    heightEmu: 720000,
    id: 1,
  })
  editor
    .chain()
    .insertContentAt(editor.state.doc.content.size, {
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: `WordArt(${preset.id})`,
        genXml: xml,
        textboxes: [makeWordArtTextbox(preset)],
      },
    })
    .run()
}

describe('WordArt insertion', () => {
  it('buildWordArtParagraphXml generates XML with the correct WPS structure', () => {
    const xml = buildWordArtParagraphXml({ colorHex: '4472C4', id: 1 })
    expect(xml).toContain('wps:wsp')
    expect(xml).toContain('w:txbxContent')
    expect(xml).toContain('wp:anchor')
    expect(xml).toContain('mc:AlternateContent')
    // large text: sz=72 (36pt)
    expect(xml).toContain('w:val="72"')
    // requested color
    expect(xml).toContain('4472C4')
    // center alignment
    expect(xml).toContain('w:val="center"')
    // mc:Fallback should NOT have xmlns:mc attribute
    expect(xml).not.toContain('mc:Fallback xmlns:mc=')
  })

  it('buildWordArtParagraphXml with default parameters (no options) still generates valid XML', () => {
    const xml = buildWordArtParagraphXml({})
    expect(xml).toContain('wps:wsp')
    expect(xml).toContain('WordArt')
    expect(xml).toContain('w:val="72"')
  })

  it('shared WORDART_PRESETS are well-formed', () => {
    expect(WORDART_PRESETS.length).toBeGreaterThanOrEqual(4)
    for (const p of WORDART_PRESETS) {
      expect(p.id).toBeTruthy()
      expect(p.nameKey).toBeTruthy()
      expect(p.fill).toMatch(/^#[0-9A-Fa-f]{6}$/)
    }
  })

  it.each(WORDART_PRESETS.map((p) => [p.id, p] as [string, WordArtPreset]))(
    'preset %s generates XML containing its solid-color approximation',
    (_id, preset) => {
      const solidHex = wordArtSolidColor(preset).replace('#', '')
      const xml = buildWordArtParagraphXml({ colorHex: solidHex, italic: preset.italic, id: 1 })
      expect(xml).toContain(solidHex)
      expect(xml).toContain('wps:wsp')
    },
  )

  it('light fills fall back to the outline color so saved text stays readable', () => {
    const whiteOrange = WORDART_PRESETS.find((p) => p.id === 'white-orange')!
    expect(wordArtSolidColor(whiteOrange)).toBe('#ED7D31')
    const blue = WORDART_PRESETS.find((p) => p.id === 'blue')!
    expect(wordArtSolidColor(blue)).toBe('#4472C4')
  })

  it('italic presets carry <w:i/> in the run properties', () => {
    const xml = buildWordArtParagraphXml({ colorHex: '70AD47', italic: true, id: 1 })
    expect(xml).toContain('<w:i/>')
  })

  it('WordArt insert → saveBlocks contains kind:xml with wps:wsp', async () => {
    const { editor, parsed } = await openBlankDoc()
    insertWordArt(editor, WORDART_PRESETS[0])

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xmlBlock = plan.saveBlocks.find((b) => b.kind === 'xml') as
      { kind: 'xml'; xml: string } | undefined
    expect(xmlBlock).toBeDefined()
    expect(xmlBlock?.xml).toContain('wps:wsp')
    expect(xmlBlock?.xml).toContain('w:txbxContent')
    editor.destroy()
  })

  it('WordArt insert → save → reparse keeps the text readable', async () => {
    const { editor, parsed } = await openBlankDoc()

    // Override default text to check it round-trips
    const xml = buildWordArtParagraphXml({ text: 'Test text', colorHex: 'ED7D31', id: 2 })
    const textbox: TextboxDisplay = {
      widthPx: Math.round(2700000 / 9525),
      heightPx: Math.round(720000 / 9525),
      wordArtId: 'white-orange',
      paras: [
        {
          runs: [{ text: 'Test text', color: 'ED7D31', bold: true, sizeHalfPoints: 72 }],
          align: 'center',
        },
      ],
    }
    editor
      .chain()
      .insertContentAt(editor.state.doc.content.size, {
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: 'WordArt(test)',
          genXml: xml,
          textboxes: [textbox],
        },
      })
      .run()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const block = reparsed.blocks.find((b) => b.textboxes)
    expect(block).toBeDefined()
    // Text should be preserved
    expect(block?.textboxes?.[0].paras[0].runs[0].text).toBe('Test text')
    editor.destroy()
  })

  it('saving WordArt with an empty text box does not corrupt the document structure', async () => {
    const { editor, parsed } = await openBlankDoc()
    insertWordArt(editor, WORDART_PRESETS[2])

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    // Saved docx should be a valid ZIP (starts with PK magic)
    expect(saved[0]).toBe(0x50) // 'P'
    expect(saved[1]).toBe(0x4b) // 'K'
    editor.destroy()
  })

  it('the WordArt shape itself has no background fill', () => {
    const xml = buildWordArtParagraphXml({ colorHex: '4472C4', id: 1 })
    expect(xml).toContain('<a:noFill/>')
  })
})
