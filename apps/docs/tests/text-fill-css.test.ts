import { describe, expect, it } from 'vitest'
import type { TextFill } from '@genoffice/docx-engine'
import {
  declStyle,
  liveTextFill,
  styleTextFillDecls,
  textFillAttr,
  textFillDecls,
} from '../src/renderer/editor/text-effects'

const GRADIENT: TextFill = {
  kind: 'gradient',
  stops: [
    { pos: 0, color: 'FF0000', alpha: 0.5 },
    { pos: 1, color: '0000FF' },
  ],
  angleDeg: 0,
}

describe('w14:textFill on the page', () => {
  it('fills the glyphs with a colour and its transparency, with a twin for the dark page', () => {
    expect(textFillDecls({ kind: 'solid', color: 'C00000' })).toEqual([
      '-webkit-text-fill-color:#C00000',
      expect.stringMatching(/^--dk-tf:#[0-9a-f]{6}$/i),
    ])
    expect(textFillDecls({ kind: 'solid', color: '000000', alpha: 0.4 })[0]).toBe(
      '-webkit-text-fill-color:rgba(0,0,0,0.4)',
    )
    expect(textFillDecls({ kind: 'none' })).toEqual(['-webkit-text-fill-color:transparent'])
  })

  it('clips a gradient to the glyphs; OOXML angles turn from pointing right', () => {
    const decls = textFillDecls(GRADIENT)
    expect(decls.slice(0, 4)).toEqual([
      'background-image:linear-gradient(90deg,rgba(255,0,0,0.5) 0%,#0000FF 100%)',
      '-webkit-background-clip:text',
      'background-clip:text',
      '-webkit-text-fill-color:transparent',
    ])
    expect(decls[4]).toMatch(/^--dk-bgi:linear-gradient\(90deg,/)
    expect(textFillDecls({ ...GRADIENT, path: 'circle' })[0]).toContain('radial-gradient(circle,')
    // a span that paints a highlight or shading keeps its plain colour
    expect(textFillDecls(GRADIENT, true)).toEqual([])
  })

  it('counts only while the text colour is the one the fill came with', () => {
    const attr = textFillAttr(GRADIENT, 'ff0000')
    expect(liveTextFill(attr, 'FF0000')).toEqual(GRADIENT)
    expect(liveTextFill(attr, '00B050')).toBeUndefined()
    expect(liveTextFill(textFillAttr({ kind: 'none' }, null), null)).toEqual({ kind: 'none' })
    expect(liveTextFill('{"kind":"gradient","stops":[]}', null)).toBeUndefined()
    expect(liveTextFill('not json', null)).toBeUndefined()
  })

  it('keeps custom properties by name in a style object', () => {
    expect(declStyle(['-webkit-text-fill-color:#C00000', '--dk-tf:#ff8080'])).toEqual({
      WebkitTextFillColor: '#C00000',
      '--dk-tf': '#ff8080',
    })
  })

  it('a style empties its glyphs only when it outlines them; a paragraph style takes no gradient', () => {
    const none: TextFill = { kind: 'none' }
    expect(styleTextFillDecls(none, { outlined: false, character: false }).decls).toEqual([])
    expect(styleTextFillDecls(none, { outlined: true, character: false }).decls).toEqual([
      '-webkit-text-fill-color:transparent',
    ])
    expect(styleTextFillDecls(GRADIENT, { outlined: false, character: false }).decls).toEqual([])
    const char = styleTextFillDecls(GRADIENT, { outlined: false, character: true })
    expect(char.decls[0]).toMatch(/^background-image:linear-gradient/)
    expect(char.darkDecls[0]).toMatch(/^background-image:linear-gradient/)
  })
})
