/**
 * Word-style dark page: authored document colors get a remapped `--dk-*` twin
 * next to the untouched authored declaration, and the generated document
 * stylesheet gets screen-only `.page-dark` twins of its color rules. The
 * authored declaration must stay the real one: exports and the pagination
 * preview render outside `.page-dark` and rely on it.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import type {
  ParsedDocFull,
  StyleDisplay,
  StyleInfo,
  TableStyleDisplay,
} from '@genoffice/docx-engine'
import {
  DARK_PAPER_HEX,
  darkPageBorderCss,
  darkPageColor,
  dkBackground,
  dkBorder,
  dkColor,
  dkStyleProps,
  dkTableBorders,
} from '../src/renderer/editor/dark-page'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { HIGHLIGHT_CSS, TextStyleMark } from '../src/renderer/editor/marks'
import { DocTable, tableBordersCss } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const STYLES_CSS = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
  'utf8',
)

/** `--name: value;` declarations of the first rule whose selector matches */
function ruleDecls(css: string, selectorRe: RegExp): Map<string, string> {
  const m = new RegExp(`${selectorRe.source}\\s*\\{([^}]*)\\}`, 'm').exec(css)
  if (!m) throw new Error(`rule not found: ${selectorRe}`)
  const out = new Map<string, string>()
  for (const d of m[1].matchAll(/(--[\w-]+):\s*([^;]+);/g)) out.set(d[1], d[2].trim())
  return out
}

function luminance(hex: string): number {
  const n = parseInt(hex.replace('#', ''), 16)
  return 0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255)
}

describe('darkPageColor', () => {
  it('inverts luminance: authored black becomes white, authored white becomes the paper tone', () => {
    expect(darkPageColor('000000')).toBe('#ffffff')
    expect(darkPageColor('#000')).toBe('#ffffff')
    expect(darkPageColor('#ffffff')).toBe(DARK_PAPER_HEX)
    expect(darkPageColor('FFFFFF')).toBe(DARK_PAPER_HEX)
  })

  it('keeps hue while flipping lightness (red text stays red, dark blue heading turns light blue)', () => {
    const red = darkPageColor('ff0000')
    expect(red).toMatch(/^#ff/)
    expect(luminance(red)).toBeGreaterThan(luminance('#ff0000'))
    const heading = darkPageColor('2f5496')
    expect(luminance(heading)).toBeGreaterThan(150)
    // blue channel stays dominant
    const n = parseInt(heading.slice(1), 16)
    expect(n & 255).toBeGreaterThan((n >> 16) & 255)
  })

  it('leaves mid-gray roughly where it is', () => {
    expect(Math.abs(luminance(darkPageColor('808080')) - 128)).toBeLessThan(20)
  })

  it('preserves alpha for rgb()/rgba() and #rrggbbaa inputs', () => {
    expect(darkPageColor('rgb(0 0 0 / 7%)')).toBe('rgb(255 255 255 / 7%)')
    expect(darkPageColor('rgba(255, 255, 255, 0.5)')).toBe('rgb(30 30 30 / 0.5)')
    expect(darkPageColor('#00000080')).toBe('#ffffff80')
  })

  it('passes anything it cannot parse through untouched', () => {
    for (const v of ['auto', 'transparent', 'currentColor', '', 'var(--x)']) {
      expect(darkPageColor(v)).toBe(v)
    }
  })

  it('remaps every color token inside a border value', () => {
    expect(darkPageBorderCss('1px solid #000')).toBe('1px solid #ffffff')
    expect(darkPageBorderCss('2px dashed #ff0000')).toMatch(/^2px dashed #ff/)
    expect(darkPageBorderCss('none')).toBe('none')
  })
})

describe('inline twins', () => {
  it('emit --dk-* declarations next to the authored value', () => {
    expect(dkColor('000000')).toBe('--dk-c:#ffffff')
    expect(dkBackground('#ffff00')).toMatch(/^--dk-bg:#/)
    expect(dkBorder('t', '1px solid #000')).toBe('--dk-b-t:1px solid #ffffff')
    expect(dkStyleProps({ color: '000000', borders: { b: '1px solid #000' } })).toEqual({
      '--dk-c': '#ffffff',
      '--dk-b-b': '1px solid #ffffff',
    })
  })

  it('table-level border variables get --dk-tb-* twins that keep "none"', () => {
    const authored = tableBordersCss({
      top: { style: 'single', szEighths: 4, color: '000000' },
      insideH: { style: 'none' },
    })
    expect(authored).toContain('--doc-b-t:1px solid #000000')
    expect(authored).toContain('--dk-tb-t:1px solid #ffffff')
    expect(authored).toContain('--doc-b-h:none')
    expect(authored).toContain('--dk-tb-h:none')
    expect(dkTableBorders(['--doc-cell-pad:0'])).toEqual([])
  })

  it('TextStyle mark keeps the authored color/highlight and adds their twins', () => {
    const spec = TextStyleMark.config.renderHTML!.call(
      TextStyleMark as never,
      {
        mark: { attrs: { color: 'ff0000', highlight: 'yellow' } },
        HTMLAttributes: {},
      } as never,
    ) as unknown as [string, Record<string, string>]
    const style = spec[1].style
    expect(style).toContain('color:#ff0000')
    expect(style).toContain(`background-color:${HIGHLIGHT_CSS.yellow}`)
    expect(style).toContain(`--dk-c:${darkPageColor('ff0000')}`)
    // the twin follows the winning background (highlight over shading)
    expect(style).toContain(`--dk-bg:${darkPageColor(HIGHLIGHT_CSS.yellow)}`)
  })
})

describe('docStyleCss dark twins', () => {
  function parsed(styles: Map<string, StyleInfo>, docDefaults = {}): ParsedDocFull {
    return { styles, docDefaults, blocks: [] } as unknown as ParsedDocFull
  }

  it('pairs every color rule with a .page-dark rule inside one @media screen block', () => {
    const styles = new Map<string, StyleInfo>()
    styles.set('Heading1', {
      styleId: 'Heading1',
      name: 'heading 1',
      type: 'paragraph',
      display: { color: '2f5496', shadingFill: 'f2f2f2' } as StyleDisplay,
    } as StyleInfo)
    styles.set('Grid', {
      styleId: 'Grid',
      name: 'Grid',
      type: 'table',
      tableDisplay: {
        fill: 'ffffff',
        band1Fill: 'd9d9d9',
        firstRow: { fill: '4472c4', color: 'ffffff', bold: true },
      } as TableStyleDisplay,
    } as StyleInfo)
    const css = docStyleCss(parsed(styles, { color: '333333' }))
    // authored rules untouched
    expect(css).toContain(
      '.doc-page [data-style="Heading1"] { color:#2f5496;background-color:#f2f2f2 }',
    )
    expect(css).toContain('color:#333333')
    // twins; every subject skips filled text boxes (light islands) with a
    // zero-specificity :where() so the twin still beats exactly its authored rule
    const island = ':where(:not(.doc-textbox-filled *))'
    const dark = css.slice(css.indexOf('@media screen {'))
    expect(dark).toContain(`.page-dark .doc-page${island} { color:${darkPageColor('333333')} }`)
    expect(dark).toContain(
      `.page-dark .doc-page [data-style="Heading1"]${island} { color:${darkPageColor('2f5496')};background-color:${darkPageColor('f2f2f2')} }`,
    )
    expect(dark).toContain(
      `.page-dark .doc-page table[data-tbl-style="Grid"] td${island}, .page-dark .doc-page table[data-tbl-style="Grid"] th${island} { background:${DARK_PAPER_HEX} }`,
    )
    expect(dark).toContain(
      `tr:nth-child(even) td${island} { background:${darkPageColor('d9d9d9')} }`,
    )
    expect(dark).toContain(
      `tr:first-child th${island} { background:${darkPageColor('4472c4')};color:${DARK_PAPER_HEX} }`,
    )
    // a filled box would inherit the remapped document default from .doc-page:
    // the island re-sets the authored one
    expect(dark).toContain('.page-dark .doc-textbox-filled { color:#333333 }')
    // non-color declarations never leak into the twins
    expect(dark).not.toContain('font-weight')
    expect(css.split('@media screen {').length).toBe(2)
  })

  it('table-level shading (w:tblPr w:shd) gets a --dk-bg twin like cell fills', () => {
    const defaults = Object.fromEntries(
      Object.entries(DocTable.config.addAttributes!.call(DocTable as never) as object).map(
        ([k, v]) => [k, (v as { default: unknown }).default],
      ),
    )
    const spec = DocTable.config.renderHTML!.call(
      DocTable as never,
      {
        node: { attrs: { ...defaults, tblFill: 'deeaf6' } },
        HTMLAttributes: {},
      } as never,
    ) as unknown as [string, Record<string, string>]
    expect(spec[1].style).toContain('background-color:#deeaf6')
    expect(spec[1].style).toContain(`--dk-bg:${darkPageColor('deeaf6')}`)
  })

  it('emits no dark block when the document declares no colors', () => {
    const css = docStyleCss(parsed(new Map(), { sizeHalfPoints: 22 }))
    expect(css).not.toContain('@media screen')
    expect(css).not.toContain('.page-dark')
  })
})

describe('styles.css dark page block', () => {
  const light = ruleDecls(STYLES_CSS, /:root,\s*\.page-dark \.doc-textbox-filled/)
  const dark = ruleDecls(STYLES_CSS, /@media screen \{\s*\.page-dark/)

  it('keeps every dark paper constant equal to darkPageColor() of its light value', () => {
    expect(dark.size).toBeGreaterThan(30)
    for (const [name, value] of dark) {
      const lightValue = light.get(name)
      expect(lightValue, name).toBeDefined()
      expect(value, name).toBe(darkPageColor(lightValue!))
    }
    expect(dark.get('--docs-paper')).toBe(DARK_PAPER_HEX)
  })

  it('filled text boxes re-enter the light constants and skip every twin rule', () => {
    // the light rule doubles as the island: same declarations, no second copy
    expect(light.get('--docs-paper-ink')).toBe('#000')
    const block = STYLES_CSS.slice(STYLES_CSS.indexOf('@media screen {\n  .page-dark {'))
    expect(block).toMatch(/\.page-dark \.doc-textbox-filled \{\s*color: var\(--docs-paper-ink\);/)
    for (const twin of ['--dk-c', '--dk-bg', '--dk-b-t', '--dk-b-r', '--dk-b-b', '--dk-b-l']) {
      expect(block).toContain(
        `[style*='${twin}:']:where(:not(.doc-textbox-filled, .doc-textbox-filled *))`,
      )
    }
    // table-level border twins: every subject carries the exclusion
    const tb = block.match(/\.page-dark \.doc-table [^{,]+/g) ?? []
    expect(tb.length).toBeGreaterThanOrEqual(10)
    for (const sel of tb) expect(sel.trim()).toMatch(/:where\(:not\(\.doc-textbox-filled \*\)\)$/)
  })
})
