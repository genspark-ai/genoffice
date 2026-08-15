import { existsSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { fontCoversText } from '../src/main/font-cmap'
import { findSystemFont } from '../src/main/font-locate'
import { hasPingFang } from './pingfang'

describe('findSystemFont', () => {
  it('resolves a ttc face by PostScript name into a standalone covering sfnt', () => {
    if (!hasPingFang) return
    const bytes = findSystemFont('PingFangSC-Regular', '')
    expect(bytes).not.toBeNull()
    // A standalone face, not the ttc container: sfnt tag, either glyf- or CFF-flavored
    expect([0x00010000, 0x74727565, 0x4f54544f]).toContain(bytes!.readUInt32BE(0))
    expect(fontCoversText(bytes!, '放心，未来十年')).toBe(true)
  })

  it('falls back to a family match when the PostScript name misses', () => {
    if (!hasPingFang) return
    const bytes = findSystemFont('PingFangSC-NoSuchStyle', 'PingFang SC')
    expect(bytes).not.toBeNull()
    expect(fontCoversText(bytes!, '放心')).toBe(true)
    // No style tokens in the PS name → the Regular face must win over other weights
    const u16 = (s: string) => Buffer.from(s, 'utf16le').swap16()
    expect(
      bytes!.includes(u16('PingFangSC-Regular')) ||
        bytes!.includes(Buffer.from('PingFangSC-Regular', 'latin1')),
    ).toBe(true)
  })

  it('family fallback ranks by exact style token, not substring', () => {
    if (!existsSync('/System/Library/Fonts/Helvetica.ttc')) return
    // Wanted "bold" must pick the Bold face over Bold Oblique (whose style text also
    // contains "bold" as a substring); the returned face's name table proves the pick
    const bytes = findSystemFont('Helvetica-BoldZZZ', 'Helvetica')
    expect(bytes).not.toBeNull()
    // PS names in Helvetica.ttc are Mac-platform (single-byte) name records
    expect(bytes!.includes(Buffer.from('Helvetica-Bold', 'latin1'))).toBe(true)
    expect(bytes!.includes(Buffer.from('Oblique', 'latin1'))).toBe(false)
  })

  it('counts Oblique faces as hits for an italic want', () => {
    if (!existsSync('/System/Library/Fonts/Helvetica.ttc')) return
    // Keep-original italic toggle appends "italic", but Helvetica names its slanted
    // face "Oblique" — the tokenizer must fold oblique into italic or the toggle
    // silently hands back the upright base face
    const bytes = findSystemFont('Helvetica-italic', 'Helvetica')
    expect(bytes).not.toBeNull()
    expect(bytes!.includes(Buffer.from('Helvetica-Oblique', 'latin1'))).toBe(true)
  })

  it('merges separate bold + italic wants into the BoldItalic face', () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Bold Italic.ttf')) return
    // Keep-original style toggle: bolding an italic run appends "bold" to the PS name,
    // so the wants are ['italic', 'bold'] — the Bold Italic face must outrank the
    // single-style Italic and Bold faces (its subfamily folds to "bolditalic")
    const bytes = findSystemFont('Arial-ItalicMT-bold', 'Arial')
    expect(bytes).not.toBeNull()
    const u16 = (s: string) => Buffer.from(s, 'utf16le').swap16()
    expect(
      bytes!.includes(u16('Arial-BoldItalicMT')) ||
        bytes!.includes(Buffer.from('Arial-BoldItalicMT', 'latin1')),
    ).toBe(true)
  })

  it('returns null when nothing matches', () => {
    expect(findSystemFont('NoSuchFont-Zzz', 'No Such Family At All')).toBeNull()
  })
})
