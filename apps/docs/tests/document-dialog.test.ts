/**
 * Word's Format > Document dialog: OK writes only what was changed, so an
 * untouched dialog leaves the section, and its fixed margins, exactly as read.
 */
import { describe, expect, it } from 'vitest'
import type { SectionSettings } from '@genoffice/docx-engine'
import {
  documentDialogChanges,
  gutterMoved,
  marginsWithGutter,
  marginsWithoutGutter,
  type DocumentDialogValue,
} from '../src/renderer/components/DocumentDialog'

const opened: DocumentDialogValue = {
  margins: { top: 1440, right: 1800, bottom: 1440, left: 1800 },
  mirror: false,
  titlePg: false,
  evenOdd: false,
  headerDist: 720,
  footerDist: 720,
  gutter: 0,
  gutterAtTop: false,
  startType: 'nextPage',
  vAlign: 'top',
  columns: 1,
  gridType: 'default',
  linePitch: 360,
  charSpace: 0,
  textDirection: 'lrTb',
  suppressEndnotes: false,
}

describe('document dialog changes', () => {
  it('writes nothing when nothing was changed', () => {
    expect(documentDialogChanges(opened, { ...opened, margins: { ...opened.margins } })).toEqual({})
  })

  it('writes the margins with the mirror choice when either changed', () => {
    const wider = { ...opened, margins: { ...opened.margins, left: 2160 } }
    expect(documentDialogChanges(opened, wider)).toEqual({
      margins: wider.margins,
      mirror: false,
      gutter: 0,
      gutterAtTop: false,
    })
    expect(documentDialogChanges(opened, { ...opened, mirror: true })).toEqual({
      margins: opened.margins,
      mirror: true,
      gutter: 0,
      gutterAtTop: false,
    })
  })

  it('writes suppress endnotes only when it was switched', () => {
    expect(documentDialogChanges(opened, { ...opened, suppressEndnotes: true })).toEqual({
      suppressEndnotes: true,
    })
    const on = { ...opened, suppressEndnotes: true }
    expect(documentDialogChanges(on, { ...on })).toEqual({})
    expect(documentDialogChanges(on, opened)).toEqual({ suppressEndnotes: false })
  })

  it('writes a changed gutter with the margins it is part of', () => {
    expect(documentDialogChanges(opened, { ...opened, gutter: 360 })).toEqual({
      margins: opened.margins,
      mirror: false,
      gutter: 360,
      gutterAtTop: false,
    })
  })

  it('writes a changed gutter position with the margins that hold the gutter', () => {
    expect(documentDialogChanges(opened, { ...opened, gutterAtTop: true })).toEqual({
      margins: opened.margins,
      mirror: false,
      gutter: 0,
      gutterAtTop: true,
    })
  })

  it('writes the section start on its own', () => {
    expect(documentDialogChanges(opened, { ...opened, startType: 'continuous' })).toEqual({
      startType: 'continuous',
    })
  })

  it('writes the vertical alignment and the columns on their own', () => {
    expect(documentDialogChanges(opened, { ...opened, vAlign: 'center' })).toEqual({
      vAlign: 'center',
    })
    expect(documentDialogChanges(opened, { ...opened, columns: 2 })).toEqual({ columns: 2 })
    expect(documentDialogChanges(opened, { ...opened, textDirection: 'tbRl' })).toEqual({
      textDirection: 'tbRl',
    })
  })

  it('writes a line pitch only for a grid that snaps lines to it', () => {
    expect(documentDialogChanges(opened, { ...opened, linePitch: 480 })).toEqual({})
    expect(documentDialogChanges(opened, { ...opened, gridType: 'lines', linePitch: 480 })).toEqual(
      { gridType: 'lines', linePitch: 480 },
    )
    const lined = { ...opened, gridType: 'linesAndChars' as const }
    expect(documentDialogChanges(lined, { ...lined, linePitch: 312 })).toEqual({ linePitch: 312 })
  })

  it('writes a character pitch only for a grid that has character cells', () => {
    expect(documentDialogChanges(opened, { ...opened, charSpace: 532 })).toEqual({})
    expect(documentDialogChanges(opened, { ...opened, gridType: 'lines', charSpace: 532 })).toEqual(
      { gridType: 'lines' },
    )
    expect(
      documentDialogChanges(opened, { ...opened, gridType: 'snapToChars', charSpace: 532 }),
    ).toEqual({ gridType: 'snapToChars', charSpace: 532 })
    const celled = { ...opened, gridType: 'linesAndChars' as const, charSpace: 532 }
    expect(documentDialogChanges(celled, { ...celled, charSpace: 0 })).toEqual({ charSpace: 0 })
  })

  it('hands on a font only when Set Font changed something', () => {
    expect(documentDialogChanges(opened, { ...opened, gridFont: {} })).toEqual({})
    expect(documentDialogChanges(opened, { ...opened, gridFont: { sizeHalfPoints: 24 } })).toEqual({
      gridFont: { sizeHalfPoints: 24 },
    })
  })

  it('writes each header and footer option on its own', () => {
    expect(documentDialogChanges(opened, { ...opened, titlePg: true })).toEqual({ titlePg: true })
    expect(documentDialogChanges(opened, { ...opened, evenOdd: true, footerDist: 360 })).toEqual({
      evenOdd: true,
      footerDist: 360,
    })
    expect(documentDialogChanges(opened, { ...opened, headerDist: 0 })).toEqual({ headerDist: 0 })
  })
})

describe('margins beside a gutter', () => {
  const section = (extra: Partial<SectionSettings>) =>
    ({
      marginTop: 1440,
      marginRight: 1800,
      marginBottom: 1440,
      marginLeft: 2520,
      ...extra,
    }) as SectionSettings

  it('shows the left margin without the gutter it was parsed with, and puts it back', () => {
    const withGutter = section({ gutter: 720 })
    const shown = marginsWithoutGutter(withGutter)
    expect(shown).toEqual({ top: 1440, right: 1800, bottom: 1440, left: 1800 })
    expect(marginsWithGutter(shown, withGutter).left).toBe(2520)
  })

  it('takes a gutter at the top out of the top margin instead', () => {
    const atTop = section({ gutter: 720, gutterAtTop: true, marginTop: 2160, marginLeft: 1800 })
    const shown = marginsWithoutGutter(atTop)
    expect(shown).toEqual({ top: 1440, right: 1800, bottom: 1440, left: 1800 })
    expect(marginsWithGutter(shown, atTop)).toMatchObject({ top: 2160, left: 1800 })
  })

  it('puts the gutter on the side the dialog names, whatever the section was read with', () => {
    const withGutter = section({ gutter: 720 })
    const shown = marginsWithoutGutter(withGutter)
    expect(marginsWithGutter(shown, withGutter, 720, true)).toMatchObject({ top: 2160, left: 1800 })
  })

  it('moves a gutter between the left and the top margin with the document setting', () => {
    const left = section({ gutter: 720 })
    const top = gutterMoved(left, true)
    expect(top).toMatchObject({ marginTop: 2160, marginLeft: 1800, gutter: 720, gutterAtTop: true })
    const back = gutterMoved(top, false)
    expect(back).toEqual(left)
    expect(gutterMoved(left, false)).toBe(left)
    const plain = section({})
    expect(gutterMoved(plain, true)).toBe(plain)
  })

  it('leaves a section without a gutter as it is', () => {
    const plain = section({})
    expect(marginsWithGutter(marginsWithoutGutter(plain), plain)).toEqual({
      top: 1440,
      right: 1800,
      bottom: 1440,
      left: 2520,
    })
  })
})
