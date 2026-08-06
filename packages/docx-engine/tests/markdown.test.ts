import { describe, expect, it } from 'vitest'
import type { Block, Run, TableModel } from '../src/types'
import { serializeBlocksToMarkdown, serializeBlocksToPlainText } from '../src/markdown'

function run(text: string, patch: Partial<Run> = {}): Run {
  return { text, ...patch }
}

function para(text: string, extra: Partial<Block> = {}): Block {
  return {
    id: 'p',
    type: 'paragraph',
    docxIndex: null,
    originalXml: null,
    runs: [run(text)],
    ...extra,
  }
}

function heading(text: string, level: number): Block {
  return { id: 'h', type: 'heading', docxIndex: null, originalXml: null, level, runs: [run(text)] }
}

function item(text: string, list: Block['list']): Block {
  return {
    id: 'li',
    type: 'listItem',
    docxIndex: null,
    originalXml: null,
    list,
    runs: [run(text)],
  }
}

const table: TableModel = {
  rows: [
    [
      { paras: ['Name'], bold: true },
      { paras: ['Score'], bold: true },
    ],
    [
      { paras: ['Ada'], richParas: [{ runs: [run('A|da')] }] },
      { paras: ['9'], richParas: [{ runs: [run('9')] }] },
    ],
  ],
}

describe('serializeBlocksToMarkdown', () => {
  it('renders headings, paragraphs and lists', () => {
    const out = serializeBlocksToMarkdown([
      heading('Intro', 1),
      para('First paragraph.'),
      heading('Details', 2),
      item('bullet one', { kind: 'bullet', numId: '1', ilvl: 0 }),
      item('nested bullet', { kind: 'bullet', numId: '1', ilvl: 1 }),
      item('ordered one', { kind: 'ordered', numId: '2', ilvl: 0 }),
    ])
    expect(out).toBe(
      [
        '# Intro',
        '',
        'First paragraph.',
        '',
        '## Details',
        '',
        '- bullet one',
        '  - nested bullet',
        '1. ordered one',
        '',
      ].join('\n'),
    )
  })

  it('applies inline formatting and links', () => {
    const out = serializeBlocksToMarkdown([
      para('', {
        runs: [
          run('plain '),
          run('bold', { bold: true }),
          run(' and '),
          run('ital', { italic: true }),
          run(' and '),
          run('strike', { strike: true }),
          run(' and '),
          run('under', { underline: true }),
          run(' and '),
          run('www', { link: { href: 'https://example.com' } }),
          run(' and '),
          run('H2O', { vertAlign: 'subscript' }),
          run('x', { vertAlign: 'superscript' }),
        ],
      }),
    ])
    expect(out).toContain('**bold**')
    expect(out).toContain('*ital*')
    expect(out).toContain('~~strike~~')
    expect(out).toContain('<u>under</u>')
    expect(out).toContain('[www](https://example.com)')
    expect(out).toContain('~H2O~')
    expect(out).toContain('^x^')
  })

  it('drops tracked deletions and keeps insertions', () => {
    const out = serializeBlocksToMarkdown([
      para('', {
        runs: [
          run('keep'),
          run(' gone', { del: { author: 'x' } }),
          run(' added', { ins: { author: 'x' } }),
        ],
      }),
    ])
    expect(out).toBe('keep added\n')
  })

  it('renders pipe tables with a header row and escapes pipes', () => {
    const out = serializeBlocksToMarkdown([{ id: 't', type: 'table', docxIndex: null, originalXml: null, table }])
    expect(out).toContain('| Name | Score |')
    expect(out).toContain('| --- | --- |')
    expect(out).toContain('| A\\|da | 9 |')
  })

  it('embeds image data URLs and falls back to alt-only', () => {
    const withUrl = serializeBlocksToMarkdown([
      { id: 'i1', type: 'image', docxIndex: null, originalXml: null, label: 'Logo', imageDataUrl: 'data:image/png;base64,AA==' },
    ])
    expect(withUrl).toBe('![Logo](data:image/png;base64,AA==)\n')
    const noUrl = serializeBlocksToMarkdown([
      { id: 'i2', type: 'image', docxIndex: null, originalXml: null, previewText: 'Chart title' },
    ])
    expect(noUrl).toBe('![Chart title]()\n')
  })

  it('skips hidden blocks and deleted revisions', () => {
    const out = serializeBlocksToMarkdown([
      para('visible'),
      { ...para('hidden'), hidden: true },
      { ...para('removed'), blockRevision: { kind: 'del', author: 'x' } },
    ])
    expect(out).toBe('visible\n')
  })

  it('renders math and formula blocks', () => {
    const out = serializeBlocksToMarkdown([
      para('', { runs: [run('E=mc', { math: { omml: '<m:oMath/>' } })] }),
      para('', { formulaDisplay: { tokens: ['x', '+', '1'] } }),
    ])
    expect(out).toContain('$E=mc$')
    expect(out).toContain('$$x + 1$$')
  })
})

describe('serializeBlocksToPlainText', () => {
  it('strips formatting but keeps structure', () => {
    const out = serializeBlocksToPlainText([
      heading('Title', 1),
      para('', { runs: [run('**bold** raw ', { bold: true }), run('and plain')] }),
    ])
    expect(out).toBe('Title\n\n**bold** raw and plain\n')
  })

  it('renders tables with tab-separated cells', () => {
    const out = serializeBlocksToPlainText([{ id: 't', type: 'table', docxIndex: null, originalXml: null, table }])
    expect(out).toContain('Name\tScore')
    expect(out).toContain('A|da\t9')
  })
})
