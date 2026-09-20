import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  activeBidi,
  alignAttrFor,
  effectiveBidi,
  firstStrongDir,
  setParagraphDirection,
  setSelectionAlign,
} from '../src/renderer/editor/direction'
import { en } from '../src/renderer/i18n/ribbon/en'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (t: string, attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content: t ? [text(t)] : [],
})

function createEditor(content: JsonNode[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

function select(editor: Editor, from: number, to = from) {
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)),
  )
}

const attrsOf = (editor: Editor, index: number) => editor.state.doc.child(index).attrs

describe('firstStrongDir', () => {
  it('detects Arabic and Hebrew as rtl', () => {
    expect(firstStrongDir('مرحبا بالعالم')).toBe('rtl')
    expect(firstStrongDir('שלום עולם')).toBe('rtl')
  })

  it('detects Latin, Cyrillic and CJK as ltr', () => {
    expect(firstStrongDir('Hello')).toBe('ltr')
    expect(firstStrongDir('Привет')).toBe('ltr')
    expect(firstStrongDir('你好')).toBe('ltr')
  })

  it('skips weak characters (digits, punctuation) before the first strong one', () => {
    expect(firstStrongDir('123 مرحبا')).toBe('rtl')
    expect(firstStrongDir('"(42) " Hello')).toBe('ltr')
    // Arabic-Indic digits are weak too: the Latin word decides
    expect(firstStrongDir('٤٢ Hello')).toBe('ltr')
  })

  it('treats script-Arabic punctuation as weak (bidi-neutral despite the script)', () => {
    // Arabic comma U+060C before a Latin word must not force rtl
    expect(firstStrongDir('\u060C Hello')).toBe('ltr')
    expect(firstStrongDir('\u060C مرحبا')).toBe('rtl')
  })

  it('returns null when there is no strong character', () => {
    expect(firstStrongDir('123 –—… !?')).toBe(null)
    expect(firstStrongDir('\u060C\u061B\u061F')).toBe(null)
    expect(firstStrongDir('')).toBe(null)
  })
})

describe('setParagraphDirection', () => {
  it('sets bidi on every paragraph-like block in the selection', () => {
    const editor = createEditor([
      para('first'),
      { type: 'docHeading', attrs: { docxIndex: null, level: 2 }, content: [text('heading')] },
      {
        type: 'docListItem',
        attrs: { docxIndex: null, kind: 'bullet', ilvl: 0 },
        content: [text('item')],
      },
    ])
    select(editor, 2, editor.state.doc.content.size - 2)
    expect(setParagraphDirection(editor, 'rtl')).toBe(true)
    expect(attrsOf(editor, 0).bidi).toBe(true)
    expect(attrsOf(editor, 1).bidi).toBe(true)
    expect(attrsOf(editor, 2).bidi).toBe(true)
    editor.destroy()
  })

  it('swaps explicit left/right alignment on a direction flip (visual align, logical preserved)', () => {
    const editor = createEditor([para('left aligned', { align: 'left' })])
    select(editor, 2)
    setParagraphDirection(editor, 'rtl')
    expect(attrsOf(editor, 0)).toMatchObject({ bidi: true, align: 'right' })
    setParagraphDirection(editor, 'ltr')
    expect(attrsOf(editor, 0)).toMatchObject({ bidi: false, align: 'left' })
    editor.destroy()
  })

  it('leaves center/justify and unset alignment alone', () => {
    const editor = createEditor([para('centered', { align: 'center' }), para('plain')])
    select(editor, 2, editor.state.doc.content.size - 2)
    setParagraphDirection(editor, 'rtl')
    expect(attrsOf(editor, 0).align).toBe('center')
    expect(attrsOf(editor, 1).align).toBe(null)
    editor.destroy()
  })

  it('is a no-op when the selection already has the direction', () => {
    const editor = createEditor([para('عربي', { bidi: true })])
    select(editor, 2)
    expect(setParagraphDirection(editor, 'rtl')).toBe(false)
    editor.destroy()
  })
})

describe('setSelectionAlign', () => {
  it('resolves the stored attr per paragraph in a mixed LTR/RTL selection', () => {
    const editor = createEditor([para('latin'), para('عربي', { bidi: true })])
    select(editor, 2, editor.state.doc.content.size - 2)
    setSelectionAlign(editor, 'right')
    // visual right is explicit for the LTR paragraph, the start default for the RTL one
    expect(attrsOf(editor, 0).align).toBe('right')
    expect(attrsOf(editor, 1).align).toBe(null)
    setSelectionAlign(editor, 'left')
    expect(attrsOf(editor, 0).align).toBe(null)
    expect(attrsOf(editor, 1).align).toBe('left')
    setSelectionAlign(editor, 'center')
    expect(attrsOf(editor, 0).align).toBe('center')
    expect(attrsOf(editor, 1).align).toBe('center')
    editor.destroy()
  })
})

describe('alignAttrFor', () => {
  it('maps start-side alignment to null (the direction default)', () => {
    expect(alignAttrFor('left', false)).toBe(null)
    expect(alignAttrFor('right', false)).toBe('right')
    expect(alignAttrFor('left', true)).toBe('left')
    expect(alignAttrFor('right', true)).toBe(null)
    expect(alignAttrFor('center', true)).toBe('center')
    expect(alignAttrFor('justify', false)).toBe('justify')
  })
})

describe('activeBidi', () => {
  it('reads the bidi attr of the paragraph at the cursor', () => {
    const editor = createEditor([para('عربي', { bidi: true }), para('latin')])
    select(editor, 2)
    expect(activeBidi(editor)).toBe(true)
    select(editor, editor.state.doc.content.size - 2)
    expect(activeBidi(editor)).toBe(false)
    editor.destroy()
  })
})

describe('AutoDirectionExtension', () => {
  it('flips a paragraph to rtl on its first strong Arabic character', () => {
    const editor = createEditor([para('')])
    select(editor, 1)
    editor.commands.insertContent('مرحبا')
    expect(attrsOf(editor, 0).bidi).toBe(true)
    editor.destroy()
  })

  it('flips an inherited rtl paragraph back to ltr when Latin text is typed first', () => {
    const editor = createEditor([para('', { bidi: true })])
    select(editor, 1)
    editor.commands.insertContent('Hello')
    expect(attrsOf(editor, 0).bidi).toBe(false)
    editor.destroy()
  })

  it('does not touch paragraphs that already contain strong text', () => {
    const editor = createEditor([para('Hello ')])
    select(editor, editor.state.doc.child(0).nodeSize - 1)
    editor.commands.insertContent('مرحبا')
    expect(attrsOf(editor, 0).bidi).toBe(false)
    editor.destroy()
  })

  it('swaps explicit alignment like the manual toggle, so the two round-trip', () => {
    // end-aligned LTR paragraph (visual right) auto-flips to rtl on Arabic input
    const editor = createEditor([para('', { align: 'right' })])
    select(editor, 1)
    editor.commands.insertContent('مرحبا')
    expect(attrsOf(editor, 0)).toMatchObject({ bidi: true, align: 'left' })
    // manual toggle back restores the original alignment
    select(editor, 2)
    setParagraphDirection(editor, 'ltr')
    expect(attrsOf(editor, 0)).toMatchObject({ bidi: false, align: 'right' })
    editor.destroy()
  })

  it('stays put while weak characters are typed, then follows the first strong one', () => {
    const editor = createEditor([para('')])
    select(editor, 1)
    editor.commands.insertContent('123 ')
    expect(attrsOf(editor, 0).bidi).toBe(false)
    editor.commands.insertContent('שלום')
    expect(attrsOf(editor, 0).bidi).toBe(true)
    editor.destroy()
  })
})

describe('inferred bidi (render-only bidiInferred attr)', () => {
  it('renders direction:rtl without the explicit bidi attr', () => {
    const editor = createEditor([para('مرحبا', { bidiInferred: true, align: 'right' })])
    const style = (editor.view.dom.querySelector('p') as HTMLElement).style
    expect(style.direction).toBe('rtl')
    expect(style.textAlign).toBe('right')
    editor.destroy()
  })

  it('activeBidi and the rtl toggle treat inferred paragraphs as rtl', () => {
    const editor = createEditor([para('مرحبا', { bidiInferred: true, align: 'right' })])
    select(editor, 2)
    expect(activeBidi(editor)).toBe(true)
    expect(setParagraphDirection(editor, 'rtl')).toBe(false)
    editor.destroy()
  })

  it('explicit ltr clears the inference and flips the visual alignment', () => {
    const editor = createEditor([para('مرحبا', { bidiInferred: true, align: 'right' })])
    select(editor, 2)
    setParagraphDirection(editor, 'ltr')
    expect(attrsOf(editor, 0)).toMatchObject({ bidi: false, bidiInferred: false, align: 'left' })
    editor.destroy()
  })

  it('align buttons resolve start/end against the inferred direction', () => {
    const editor = createEditor([para('مرحبا', { bidiInferred: true })])
    select(editor, 2)
    setSelectionAlign(editor, 'left')
    expect(attrsOf(editor, 0).align).toBe('left')
    setSelectionAlign(editor, 'right')
    expect(attrsOf(editor, 0).align).toBe(null)
    editor.destroy()
  })

  it('auto-detect leaves a weak-only inferred paragraph alone on a strong RTL character', () => {
    const editor = createEditor([para('42 ', { bidiInferred: true, align: 'right' })])
    select(editor, editor.state.doc.child(0).nodeSize - 1)
    editor.commands.insertContent('مرحبا')
    expect(attrsOf(editor, 0)).toMatchObject({ bidi: false, bidiInferred: true, align: 'right' })
    editor.destroy()
  })

  it('auto-detect flips a weak-only inferred paragraph to ltr when Latin is typed first', () => {
    const editor = createEditor([para('42 ', { bidiInferred: true, align: 'right' })])
    select(editor, editor.state.doc.child(0).nodeSize - 1)
    editor.commands.insertContent('Hello')
    expect(attrsOf(editor, 0)).toMatchObject({ bidi: false, bidiInferred: false, align: 'left' })
    editor.destroy()
  })
})

// bidiVisual tables mirror column order via dir="rtl" on the <table>, but Word
// keeps each cell paragraph's base direction governed by its own w:bidi only:
// weak-only text like "50,0 %" must not reorder to "% 50,0".
describe('bidiVisual tables do not reorder cell text', () => {
  it('editor paragraphs carry an explicit direction, so they never inherit the table dir', () => {
    const editor = createEditor([para('50,0 %'), para('مرحبا', { bidiInferred: true })])
    const html = editor.getHTML()
    expect(html).toContain('direction: ltr')
    expect(html).toContain('direction: rtl')
    editor.destroy()
  })
})

describe('effectiveBidi', () => {
  it('is true for explicit bidi or the render-only inference', () => {
    expect(effectiveBidi({ bidi: true })).toBe(true)
    expect(effectiveBidi({ bidiInferred: true })).toBe(true)
    expect(effectiveBidi({ bidi: true, bidiInferred: true })).toBe(true)
  })

  it('is false when neither flag is set', () => {
    expect(effectiveBidi({})).toBe(false)
    expect(effectiveBidi({ bidi: false })).toBe(false)
    expect(effectiveBidi({ bidi: false, bidiInferred: false })).toBe(false)
  })
})

describe('paragraph direction for mixed content (dir=auto semantics)', () => {
  it('the first strong character wins, regardless of what follows', () => {
    expect(firstStrongDir('Hello مرحبا')).toBe('ltr')
    expect(firstStrongDir('مرحبا Hello')).toBe('rtl')
    expect(firstStrongDir('שלום Hello مرحبا')).toBe('rtl')
    expect(firstStrongDir('Hello שלום')).toBe('ltr')
  })

  it('weak prefixes never decide the direction', () => {
    expect(firstStrongDir('123, Hello')).toBe('ltr')
    expect(firstStrongDir('123, مرحبا')).toBe('rtl')
    expect(firstStrongDir('... (...) ...')).toBe(null)
  })

  it('each paragraph resolves from its own attrs in a mixed selection', () => {
    expect(effectiveBidi({ bidi: false, bidiInferred: false })).toBe(false)
    expect(effectiveBidi({ bidi: false, bidiInferred: true })).toBe(true)
    expect(effectiveBidi({ bidi: true, bidiInferred: false })).toBe(true)
  })
})

describe('direction flip preserves logical alignment', () => {
  it('start-side alignment stays the start default (null) in both directions', () => {
    // alignAttrFor is what the align buttons write; the direction flip swaps
    // the stored visual value so the same logical side survives the toggle
    expect(alignAttrFor('left', false)).toBe(null)
    expect(alignAttrFor('right', true)).toBe(null)
  })

  it('end-side alignment stays explicit in both directions', () => {
    expect(alignAttrFor('right', false)).toBe('right')
    expect(alignAttrFor('left', true)).toBe('left')
  })

  it('center and justify are direction-independent', () => {
    for (const bidi of [false, true]) {
      expect(alignAttrFor('center', bidi)).toBe('center')
      expect(alignAttrFor('justify', bidi)).toBe('justify')
    }
  })
})

describe('ribbon direction buttons carry translated aria-labels', () => {
  it('reuses the existing ribbonDirLtrTip / ribbonDirRtlTip keys (English, distinct)', () => {
    expect(en.ribbonDirLtrTip).toBeTruthy()
    expect(en.ribbonDirRtlTip).toBeTruthy()
    expect(en.ribbonDirLtrTip).not.toBe(en.ribbonDirRtlTip)
    expect(en.ribbonDirLtrTip).toMatch(/left.*right/i)
    expect(en.ribbonDirRtlTip).toMatch(/right.*left/i)
  })

  it('wires both buttons with data-tip and aria-label (no new keys)', () => {
    const ribbon = readFileSync(join(__dirname, '../src/renderer/components/Ribbon.tsx'), 'utf8')
    for (const key of ['ribbonDirLtrTip', 'ribbonDirRtlTip']) {
      expect(ribbon).toContain(`aria-label={t('${key}')}`)
      expect(ribbon).toContain(`data-tip={t('${key}')}`)
    }
  })
})

describe('ai panel messages follow their own content direction', () => {
  it('renders historic, live assistant, and user text with dir=auto', () => {
    // Panel chrome follows the UI language; message text must follow its own
    // content, so every message body carries dir="auto" (markup contract in
    // source, no browser needed)
    const panel = readFileSync(join(__dirname, '../src/renderer/ai/AiPanel.tsx'), 'utf8')
    const hits = panel.match(/dir="auto"/g) ?? []
    expect(hits.length).toBeGreaterThanOrEqual(3)
  })

  it('resolves dir=auto message direction from the first strong character', () => {
    expect(firstStrongDir('שלום, how are you?')).toBe('rtl')
    expect(firstStrongDir('How are you, שלום?')).toBe('ltr')
    expect(firstStrongDir('مرحبا! 123 Hello')).toBe('rtl')
    expect(firstStrongDir('123 Hello مرحبا')).toBe('ltr')
  })
})
