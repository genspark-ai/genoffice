import { OFFICE_THEMES, parseDocx, themedDocxBytes, type NoteInfo } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { themedPartsDocx } from '../../../packages/docx-engine/tests/helpers/themed-parts-docx'
import {
  partsReading,
  rethemedNotes,
  rethemedValue,
} from '../src/renderer/editor/theme-retheme-parts'

const theme = (id: string) => {
  const t = OFFICE_THEMES.find((x) => x.id === id)!
  return { fonts: t.fonts, colors: t.colors, format: { name: t.name, xml: t.format } }
}

async function readings() {
  const bytes = await themedPartsDocx()
  const own = partsReading(await parseDocx(bytes))
  const facet = partsReading(await parseDocx(await themedDocxBytes(bytes, theme('facet'))))
  return { own, facet }
}

describe('rethemedValue', () => {
  it('takes a value to the picked reading only where the old reading still stands', () => {
    const from = { a: { color: '#111111', size: 10 }, runs: [{ color: '#222222' }] }
    const to = { a: { color: '#AAAAAA', size: 10 }, runs: [{ color: '#BBBBBB' }] }
    const live = { a: { color: '#111111', size: 12 }, runs: [{ color: '#00FF00' }], own: 1 }
    expect(rethemedValue(live, from, to)).toEqual({
      a: { color: '#AAAAAA', size: 12 },
      runs: [{ color: '#00FF00' }],
      own: 1,
    })
  })

  it('returns the same object when the theme decides nothing in it', () => {
    const live = { a: [1, { b: 'x' }] }
    expect(rethemedValue(live, { a: [1, { b: 'x' }] }, { a: [1, { b: 'x' }] })).toBe(live)
  })

  it('leaves a list alone whose length the user changed', () => {
    const live = [{ c: 'old' }, { c: 'new paragraph' }]
    expect(rethemedValue(live, [{ c: 'old' }], [{ c: 'themed' }])).toBe(live)
  })

  it('adds and drops a key with the reading', () => {
    expect(rethemedValue({ k: 1 }, { k: 1 }, { k: 1, tint: 'x' })).toEqual({ k: 1, tint: 'x' })
    expect(rethemedValue({ k: 1, tint: 'x' }, { k: 1, tint: 'x' }, { k: 1 })).toEqual({ k: 1 })
  })
})

describe('a theme pick on headers, footers and notes', () => {
  it('reads every themed part differently under another theme', async () => {
    const { own, facet } = await readings()
    const text = (v: unknown) => JSON.stringify(v)
    expect(text(own.strips.header)).toContain('"color":"ED7D31"')
    expect(text(facet.strips.header)).toContain('"color":"54A021"')
    expect(text(facet.strips.header)).toContain('"font":"Trebuchet MS"')
    expect(text(facet.strips.header)).toContain('"color":"FF0000"')
    expect(text(facet.strips.header)).toContain('"color":"C0E474"')
    expect(text(facet.strips.footer)).not.toBe(text(own.strips.footer))
    expect(text(facet.strips.variants.headerFirst)).not.toBe(text(own.strips.variants.headerFirst))
    expect(text(facet.parsed.hfParts)).not.toBe(text(own.parsed.hfParts))
    expect(text(facet.footnotes)).not.toBe(text(own.footnotes))
    expect(text(facet.endnotes)).not.toBe(text(own.endnotes))
  })

  it('puts untouched parts exactly where save and reopen reads them, and back', async () => {
    const { own, facet } = await readings()
    const live = rethemedValue(own.strips, own.strips, facet.strips)
    expect(live).toEqual(facet.strips)
    expect(rethemedValue(live, facet.strips, own.strips)).toEqual(own.strips)
    const notes = rethemedNotes(own.footnotes, own.footnotes, facet.footnotes)
    expect(notes).toEqual(facet.footnotes)
    expect(rethemedNotes(notes, facet.footnotes, own.footnotes)).toEqual(own.footnotes)
  })

  it('keeps what the user changed in a header and a note', async () => {
    const { own, facet } = await readings()
    const edited = JSON.parse(
      JSON.stringify(own.strips.header).replaceAll('"color":"ED7D31"', '"color":"123456"'),
    ) as typeof own.strips.header
    const live = rethemedValue(edited, own.strips.header, facet.strips.header)
    expect(JSON.stringify(live)).toContain('"color":"123456"')
    expect(JSON.stringify(live)).toContain('"font":"Trebuchet MS"')
    const added: NoteInfo = { id: '9', text: 'typed', created: true }
    const retyped = { ...own.footnotes[0], text: 'retyped', richParas: undefined }
    const notes = rethemedNotes([retyped, added], own.footnotes, facet.footnotes)
    expect(notes[0].text).toBe('retyped')
    expect(notes[1]).toBe(added)
  })

  it('leaves notes alone that the theme decides nothing in', async () => {
    const { own } = await readings()
    expect(rethemedNotes(own.footnotes, own.footnotes, own.footnotes)).toBe(own.footnotes)
  })
})
