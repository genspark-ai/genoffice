import { describe, expect, it } from 'vitest'

import {
  auditHtmlLayout,
  layoutAuditNote,
  sectionOfBlock,
  textColumnWidthTwips,
} from '../src/renderer/ai/layout-audit'
import type { AiSectionState } from '../src/renderer/ai/page-setup'
import { parseHtmlFragment } from '../src/renderer/ai/protocol'

const NUM_IDS = { bullet: null, ordered: null }
const A4_TEXT = 9026

function section(over: Partial<AiSectionState> = {}): AiSectionState {
  return {
    index: 0,
    firstBlock: 0,
    lastBlock: 9,
    paper: 'A4',
    width: 11906,
    height: 16838,
    orientation: 'portrait',
    margins: {
      top: 1440,
      right: 1440,
      bottom: 1440,
      left: 1440,
      header: 720,
      footer: 720,
      gutter: 0,
    },
    columns: { count: 1, spacing: 720 },
    titlePg: false,
    startType: 'nextPage',
    ...over,
  }
}

describe('parseHtmlFragment with wrapper elements', () => {
  it('keeps the blocks inside nested <div> wrappers instead of one merged paragraph', () => {
    const nodes = parseHtmlFragment(
      '<div style="display:flex"><div><h2>Skills</h2><ul><li>Rust</li></ul></div>' +
        '<div><h2>Experience</h2><p>Acme, 2020</p></div></div>',
      NUM_IDS,
    )
    expect(nodes.map((n) => n.type)).toEqual([
      'docHeading',
      'docListItem',
      'docHeading',
      'docParagraph',
    ])
  })

  it('keeps loose inline content beside a block as one paragraph with its marks', () => {
    const nodes = parseHtmlFragment(
      '<div><strong>Name:</strong> John <a href="https://example.com">site</a><p>Next</p> tail</div>',
      NUM_IDS,
    )
    expect(nodes.map((n) => n.type)).toEqual(['docParagraph', 'docParagraph', 'docParagraph'])
    const first = nodes[0].content ?? []
    expect(first.map((n) => n.text).join('')).toBe('Name: John site')
    expect(first[0].marks?.map((m) => m.type)).toEqual(['bold'])
    expect(first.at(-1)?.marks?.[0].type).toBe('link')
  })

  it('still reads a <div> of inline content as one paragraph', () => {
    const nodes = parseHtmlFragment('<div>plain <strong>bold</strong> text</div>', NUM_IDS)
    expect(nodes.map((n) => n.type)).toEqual(['docParagraph'])
  })
})

describe('auditHtmlLayout', () => {
  it('is silent for content the editor holds as written', () => {
    const html =
      '<h1>Report</h1><p>Body</p><table><tr><th>Name</th><th>Value</th></tr><tr><td>a</td><td>1</td></tr></table>'
    expect(auditHtmlLayout(html, A4_TEXT)).toEqual([])
    expect(layoutAuditNote(html, A4_TEXT)).toBe('')
    expect(auditHtmlLayout('plain text, no tags', A4_TEXT)).toEqual([])
  })

  it('reports a side-by-side layout asked for in CSS', () => {
    for (const css of ['display:flex', 'display: grid', 'column-count: 2', 'float:left']) {
      const findings = auditHtmlLayout(`<div style="${css}"><p>a</p><p>b</p></div>`)
      expect(findings).toHaveLength(1)
      expect(findings[0]).toContain('ONE column')
      expect(findings[0]).toContain('set_page_setup')
    }
    expect(auditHtmlLayout('<style>.cv{display:grid}</style><p>a</p>')).toHaveLength(1)
  })

  it('reports merged cells, nested tables and block markup inside cells', () => {
    const findings = auditHtmlLayout(
      '<table><tr><th colspan="2">Head</th></tr>' +
        '<tr><td><ul><li>x</li></ul></td><td><table><tr><td>in</td></tr></table></td></tr></table>',
    )
    expect(findings.some((f) => f.includes('colspan/rowspan on 1 cell'))).toBe(true)
    expect(findings.some((f) => f.includes('table inside a cell'))).toBe(true)
    expect(findings.some((f) => f.includes('inside 1 cell(s)'))).toBe(true)
  })

  it('reports for a PDF only what the printed page loses', () => {
    const html =
      '<div style="display:flex"><table><tr><td colspan="2"><ul><li>a</li></ul><table><tr><td>x</td></tr></table></td></tr>' +
      `<tr>${'<td>c</td>'.repeat(20)}</tr></table></div>`
    const findings = auditHtmlLayout(html, 9026, 'pdf')
    expect(findings).toHaveLength(2)
    expect(findings[0]).toContain('side by side')
    expect(findings[1]).toContain('colspan/rowspan')
    expect(findings.join(' ')).not.toContain('set_page_setup')
  })

  it('reports columns too narrow for the page', () => {
    const cells = Array.from({ length: 14 }, (_, i) => `<td>c${i}</td>`).join('')
    const findings = auditHtmlLayout(`<table><tr>${cells}</tr></table>`, A4_TEXT)
    expect(findings).toHaveLength(1)
    expect(findings[0]).toContain('14 columns')
    // the same table fits a landscape page
    expect(auditHtmlLayout(`<table><tr>${cells}</tr></table>`, 13958)).toEqual([])
  })

  it('reports an unbreakable string wider than its column, and not text that can wrap', () => {
    const url = 'https://example.com/a/very/long/path/that/cannot/wrap/anywhere/at/all/index.html'
    const row = (text: string) =>
      `<table><tr><td>${text}</td><td>b</td><td>c</td><td>d</td></tr></table>`
    const findings = auditHtmlLayout(row(url), A4_TEXT)
    expect(findings).toHaveLength(1)
    expect(findings[0]).toContain('column 1')
    expect(findings[0]).toContain(`${url.length} characters`)
    expect(auditHtmlLayout(row('many short words that wrap without any trouble'), A4_TEXT)).toEqual(
      [],
    )
  })

  it('does not count CJK text as unbreakable', () => {
    const cjk = String.fromCharCode(0x4e2d).repeat(60)
    expect(
      auditHtmlLayout(`<table><tr><td>${cjk}</td><td>b</td><td>c</td></tr></table>`, A4_TEXT),
    ).toEqual([])
  })

  it('skips the page-fit checks without a text width', () => {
    const cells = Array.from({ length: 14 }, () => '<td>c</td>').join('')
    expect(auditHtmlLayout(`<table><tr>${cells}</tr></table>`)).toEqual([])
  })
})

describe('text column width', () => {
  it('subtracts margins, gutter and the gaps between columns', () => {
    expect(textColumnWidthTwips(section())).toBe(9026)
    expect(textColumnWidthTwips(section({ columns: { count: 2, spacing: 720 } }))).toBe(4153)
  })

  it('finds the section of a block and falls back to the first', () => {
    const sections = [section(), section({ index: 1, firstBlock: 10, lastBlock: 20 })]
    expect(sectionOfBlock(sections, 12)?.index).toBe(1)
    expect(sectionOfBlock(sections, 99)?.index).toBe(0)
    expect(sectionOfBlock([], 0)).toBeUndefined()
  })
})
