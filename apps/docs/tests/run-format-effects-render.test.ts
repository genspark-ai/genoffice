/** Rendering of legacy text effects, glow, w:w compression, distribute and the Hyperlink style. */
import { Editor } from '@tiptap/core'
import { DOMSerializer } from '@tiptap/pm/model'
import { parseDocx } from '@genoffice/docx-engine'
import type { ParsedDocFull, StyleDisplay, StyleInfo } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { renderFieldSpec } from '../src/renderer/editor/protected-render'
import {
  charScaleXDecls,
  glowDecl,
  textAlignDecl,
  textEffectDecls,
} from '../src/renderer/editor/text-effects'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function parsedWith(
  styleId: string,
  type: StyleInfo['type'],
  display: StyleDisplay,
): ParsedDocFull {
  const styles = new Map<string, StyleInfo>()
  styles.set(styleId, { styleId, name: styleId, type, display } as StyleInfo)
  return { styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull
}

async function renderSpans(bodyXml: string): Promise<HTMLElement[]> {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const spans = [...editor.view.dom.querySelectorAll('p span')] as HTMLElement[]
  editor.destroy()
  return spans
}

describe('text effect declarations', () => {
  it('hollows outline glyphs and paints engraved glyphs in the paper colour', () => {
    expect(textEffectDecls('outline')).toEqual([
      '-webkit-text-stroke:0.5px currentColor',
      '-webkit-text-fill-color:transparent',
    ])
    expect(textEffectDecls('imprint')[0]).toBe('color:var(--docs-paper)')
    expect(textEffectDecls('imprint')[1]).toContain('text-shadow:-1px -1px')
    expect(textEffectDecls('emboss')[1]).toContain('text-shadow:1px 1px')
  })

  it('maps glow to a blurred shadow with alpha', () => {
    expect(glowDecl({ color: 'FF0000', radiusPt: 8, alpha: 0.6 })).toBe(
      'text-shadow:0 0 4pt rgba(255,0,0,0.6),0 0 8pt rgba(255,0,0,0.6)',
    )
  })

  it('compresses a whitespace-free run and gives the saved width back', () => {
    expect(charScaleXDecls('{"s":0.33,"gapEm":-0.348}')).toEqual([
      'display:inline-block',
      'transform:scaleX(0.33)',
      'transform-origin:0 50%',
      'margin-right:-0.348em',
    ])
    expect(charScaleXDecls('nope')).toEqual([])
  })

  it('spreads every line of a distributed paragraph', () => {
    expect(textAlignDecl('distribute')).toBe(
      'text-align:justify;text-align-last:justify;text-justify:distribute',
    )
    expect(textAlignDecl('right')).toBe('text-align:right')
  })
})

describe('editor run rendering', () => {
  it('engraved white text stays visible through its shadow', async () => {
    const spans = await renderSpans(
      '<w:p><w:r><w:rPr><w:imprint/><w:color w:val="FFFFFF"/></w:rPr><w:t>engraved</w:t></w:r></w:p>',
    )
    const span = spans.find((s) => s.textContent === 'engraved')!
    expect(span.style.color).toBe('var(--docs-paper)')
    expect(span.style.textShadow).toBe('-1px -1px 0 var(--docs-paper-ink-mid)')
    expect(span.getAttribute('style')).not.toContain('--dk-c')
  })

  it('scales a single glyph with w:w instead of letter-spacing it', async () => {
    const spans = await renderSpans(
      '<w:p><w:r><w:rPr><w:w w:val="33"/></w:rPr><w:t>x</w:t></w:r>' +
        '<w:r><w:rPr><w:w w:val="50"/></w:rPr><w:t xml:space="preserve">two words</w:t></w:r></w:p>',
    )
    const single = spans.find((s) => s.textContent === 'x')!
    expect(single.style.transform).toBe('scaleX(0.33)')
    expect(single.style.marginRight).toBe('-0.348em')
    expect(single.style.letterSpacing).toBe('')
    const spaced = spans.find((s) => s.textContent === 'two words')!
    expect(spaced.style.letterSpacing).not.toBe('')
    expect(spaced.style.transform).toBe('')
  })

  it('lowers a run by its w:position', async () => {
    const spans = await renderSpans(
      '<w:p><w:r><w:rPr><w:position w:val="-8"/></w:rPr><w:t>low</w:t></w:r></w:p>',
    )
    expect(spans.find((s) => s.textContent === 'low')!.style.verticalAlign).toBe('-4pt')
  })
})

describe('Hyperlink style rendering', () => {
  it('lets a defined Hyperlink style beat the default link look', () => {
    const css = docStyleCss(
      parsedWith('Hyperlink', 'character', { color: '008000', underline: false }),
    )
    expect(css).toContain('[data-style="Hyperlink"] { color:#008000 }')
    expect(css).toContain(
      '.doc-link:has([data-style="Hyperlink"]) { color:inherit;text-decoration:none }',
    )
    expect(docStyleCss(parsedWith('Body', 'paragraph', { align: 'distribute' }))).toContain(
      'text-align-last:justify',
    )
    expect(docStyleCss(parsedWith('Body', 'paragraph', {}))).not.toContain('.doc-link:has')
  })

  it('carries the entry run style onto the TOC title', () => {
    const dom = DOMSerializer.renderSpec(
      document,
      renderFieldSpec({
        kind: 'tocLine',
        left: 'H1',
        right: '1',
        runStyleId: 'Hyperlink',
      }) as never,
    ).dom as HTMLElement
    expect(dom.querySelector('.doc-toc-title')?.getAttribute('data-style')).toBe('Hyperlink')
  })
})
