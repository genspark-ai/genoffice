/**
 * Font dialog Advanced tab and effects: the docTextStyle attrs it writes round-trip
 * to the run model (both directions) and untouched fields stay out of the patch.
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { CellSelection } from '@tiptap/pm/tables'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  charSpacingAttrs,
  charSpacingFromAttrs,
  fontTabAttrs,
  selectionRanges,
  setTextStyleWithScale,
} from '../src/renderer/components/FontDialog'
import { blocksToPmDoc, inlineToRuns, runsToInline } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const styleAttrs = (inline: ReturnType<typeof runsToInline>) =>
  (inline[0] as { marks?: Array<{ type: string; attrs: Record<string, unknown> }> }).marks?.find(
    (m) => m.type === 'docTextStyle',
  )?.attrs

describe('character spacing state', () => {
  it('reads Word dialog fields from run attrs', () => {
    const s = charSpacingFromAttrs(
      { charScalePct: 150, charSpacingTwips: -30, positionHalfPoints: 6, kernHalfPoints: 28 },
      11,
    )
    expect(s).toEqual({
      scalePct: 150,
      spacing: 'condensed',
      spacingPt: 1.5,
      position: 'raised',
      positionPt: 3,
      kern: true,
      kernPt: 14,
    })
    expect(charSpacingFromAttrs({}, 12)).toMatchObject({
      scalePct: 100,
      spacing: 'normal',
      position: 'normal',
      kern: false,
      kernPt: 12,
    })
  })

  it('writes only the touched groups', () => {
    const state = charSpacingFromAttrs({}, 11)
    expect(charSpacingAttrs(state, new Set(), {}, 22)).toEqual({})
    const out = charSpacingAttrs(
      { ...state, spacing: 'expanded', spacingPt: 2, scalePct: 150, kern: true, kernPt: 12 },
      new Set(['spacing', 'scale', 'kern']),
      {},
      22,
    )
    expect(out).toEqual({
      charSpacingTwips: 40,
      charScalePct: 150,
      charScaleEm: null,
      charScaleX: null,
      kernHalfPoints: 24,
      kern: false,
    })
  })

  it('normal / off remove the value, or write an explicit off over a prior one', () => {
    const state = charSpacingFromAttrs({ charSpacingTwips: 40, kernHalfPoints: 28 }, 11)
    const cleared = charSpacingAttrs(
      { ...state, spacing: 'normal', kern: false, scalePct: 100 },
      new Set(['spacing', 'kern', 'scale']),
      { charSpacingTwips: 40, kernHalfPoints: 28 },
      22,
    )
    expect(cleared).toMatchObject({
      charSpacingTwips: null,
      charScalePct: null,
      kernHalfPoints: 0,
      kern: false,
    })
    expect(charSpacingAttrs({ ...state, kern: false }, new Set(['kern']), {}, 22)).toMatchObject({
      kernHalfPoints: null,
      kern: null,
    })
  })
})

describe('run <-> mark round trip of the dialog fields', () => {
  it('carries scale, kern threshold, own vanish, caps and explicit dstrike off', () => {
    const inline = runsToInline([
      {
        text: 'Spaced',
        charScalePct: 150,
        kernHalfPoints: 28,
        sizeHalfPoints: 24,
        caps: 'small',
        dstrike: false,
        vanish: true,
        vanishOwn: true,
        positionHalfPoints: -4,
        charSpacingTwips: 40,
      },
    ])
    expect(styleAttrs(inline)).toMatchObject({
      charScalePct: 150,
      kernHalfPoints: 28,
      kern: false,
      caps: 'small',
      dstrike: false,
      vanish: true,
      vanishOwn: true,
    })
    expect(inlineToRuns(inline)[0]).toMatchObject({
      charScalePct: 150,
      kernHalfPoints: 28,
      caps: 'small',
      dstrike: false,
      vanish: true,
      vanishOwn: true,
      positionHalfPoints: -4,
      charSpacingTwips: 40,
    })
  })

  it('an inherited vanish stays display-only', () => {
    const inline = runsToInline([{ text: 'hidden by style', vanish: true }])
    const run = inlineToRuns(inline)[0]
    expect(run.vanish).toBe(true)
    expect(run.vanishOwn).toBeUndefined()
  })

  it('renders the dialog fields on the editable span', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'Spaced',
                marks: [
                  {
                    type: 'docTextStyle',
                    attrs: {
                      charSpacingTwips: 40,
                      charScaleEm: 0.26,
                      caps: 'small',
                      dstrike: true,
                      positionHalfPoints: 6,
                      kern: false,
                    },
                  },
                ],
              },
            ],
          },
        ],
      } as never,
    })
    const span = editor.view.dom.querySelector('span[style]') as HTMLElement
    // jsdom normalizes pt to px inside calc()
    expect(span.style.letterSpacing).toContain('0.26em')
    expect(span.style.letterSpacing).toMatch(/2pt|2\.66+7px/)
    expect(span.style.fontVariantCaps).toBe('small-caps')
    expect(span.style.textDecoration).toContain('line-through')
    expect(span.style.verticalAlign).toBe('3pt')
    expect(span.style.fontKerning).toBe('none')
    editor.destroy()
  })
})

describe('dialog apply is one undo step', () => {
  it('scale is written with the other attrs and a single undo restores the original marks', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'Spaced',
                marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 24, color: '1F4E78' } }],
              },
            ],
          },
        ],
      } as never,
    })
    const styleOf = () =>
      editor.state.doc.firstChild!.firstChild!.marks.find((m) => m.type.name === 'docTextStyle')!
        .attrs
    const before = styleOf()
    editor
      .chain()
      .selectAll()
      .command(({ tr }) =>
        setTextStyleWithScale(tr, selectionRanges(tr.selection), {
          charScalePct: 150,
          charScaleX: null,
          charScaleEm: null,
          charSpacingTwips: 40,
          caps: 'small',
        }),
      )
      .setMark('bold')
      .run()
    expect(styleOf()).toMatchObject({
      sizeHalfPoints: 24,
      color: '1F4E78',
      charScalePct: 150,
      charScaleEm: 0.26,
      charSpacingTwips: 40,
      caps: 'small',
    })
    expect(editor.isActive('bold')).toBe(true)
    editor.commands.undo()
    expect(styleOf()).toEqual(before)
    expect(editor.isActive('bold')).toBe(false)
    editor.destroy()
  })
})

describe('table cell selections', () => {
  const TABLE =
    '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

  it('a column selection formats only its own cells', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const cells: number[] = []
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docTableCell') cells.push(pos)
    })
    // first column: A (row 1) and C (row 2); the CellSelection spans B in document order
    editor.view.dispatch(
      editor.state.tr.setSelection(CellSelection.create(editor.state.doc, cells[0], cells[2])),
    )
    expect(selectionRanges(editor.state.selection)).toHaveLength(2)
    editor
      .chain()
      .command(({ tr }) =>
        setTextStyleWithScale(tr, selectionRanges(tr.selection), {
          caps: 'all',
          charScalePct: 150,
        }),
      )
      .run()
    const capsOf: Record<string, unknown> = {}
    editor.state.doc.descendants((node) => {
      if (node.isText)
        capsOf[node.text!] =
          node.marks.find((m) => m.type.name === 'docTextStyle')?.attrs.caps ?? null
    })
    expect(capsOf).toEqual({ A: 'all', B: null, C: 'all', D: null })
    editor.destroy()
  })
})

describe('caret apply', () => {
  it('keeps the run attrs at the caret and adds the dialog attrs to the stored mark', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'Spaced',
                marks: [
                  {
                    type: 'docTextStyle',
                    attrs: {
                      color: '1F4E78',
                      styleId: 'Strong',
                      rawRPr: '<w:rPr><w:color w:val="1F4E78"/></w:rPr>',
                      charSpacingTwips: 20,
                    },
                  },
                ],
              },
            ],
          },
        ],
      } as never,
    })
    editor.commands.setTextSelection(4)
    editor
      .chain()
      .command(({ tr }) =>
        setTextStyleWithScale(tr, selectionRanges(tr.selection), { caps: 'all' }),
      )
      .insertContent('Z')
      .run()
    const inserted = editor.state.doc.firstChild!.child(1)
    expect(inserted.text).toBe('Z')
    expect(inserted.marks.find((m) => m.type.name === 'docTextStyle')!.attrs).toMatchObject({
      color: '1F4E78',
      styleId: 'Strong',
      rawRPr: '<w:rPr><w:color w:val="1F4E78"/></w:rPr>',
      charSpacingTwips: 20,
      caps: 'all',
    })
    editor.destroy()
  })
})

describe('mixed selections keep untouched fields', () => {
  const FONT_TAB = {
    fontLatin: 'Arial',
    fontEastAsia: '',
    sizePt: 11,
    colorHex: '#000000',
    vertAlign: '',
  }

  it('the Font tab writes only touched attrs', () => {
    expect(fontTabAttrs(FONT_TAB, new Set())).toEqual({})
    expect(fontTabAttrs(FONT_TAB, new Set(['size', 'fontEastAsia']))).toEqual({
      sizeHalfPoints: 22,
      font: null,
      eastAsiaFont: null,
      eaSlotEmpty: null,
    })
  })

  it('toggling small caps over two sizes leaves both sizes in place', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'big ',
                marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 56, color: 'FF0000' } }],
              },
              {
                type: 'text',
                text: 'small',
                marks: [
                  { type: 'docTextStyle', attrs: { sizeHalfPoints: 20, fontAscii: 'Arial' } },
                ],
              },
            ],
          },
        ],
      } as never,
    })
    editor
      .chain()
      .selectAll()
      .command(({ tr }) =>
        setTextStyleWithScale(tr, selectionRanges(tr.selection), {
          ...fontTabAttrs(FONT_TAB, new Set()),
          caps: 'small',
        }),
      )
      .run()
    const runs: Array<Record<string, unknown>> = []
    editor.state.doc.descendants((node) => {
      if (node.isText) runs.push(node.marks.find((m) => m.type.name === 'docTextStyle')!.attrs)
    })
    expect(runs[0]).toMatchObject({ sizeHalfPoints: 56, color: 'FF0000', caps: 'small' })
    expect(runs[1]).toMatchObject({ sizeHalfPoints: 20, fontAscii: 'Arial', caps: 'small' })
    editor.destroy()
  })

  it('a caret scale change carries a display twin for the typed text', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'ab' }] }],
      } as never,
    })
    editor.commands.setTextSelection(2)
    editor
      .chain()
      .command(({ tr }) =>
        setTextStyleWithScale(tr, selectionRanges(tr.selection), {
          charScalePct: 150,
          charScaleEm: null,
          charScaleX: null,
        }),
      )
      .insertContent('Z')
      .run()
    const z = editor.state.doc.firstChild!.child(1)
    expect(z.text).toBe('Z')
    expect(z.marks[0].attrs).toMatchObject({ charScalePct: 150, charScaleEm: 0.26 })
    editor.destroy()
  })
})
