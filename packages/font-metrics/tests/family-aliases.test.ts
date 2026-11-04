/**
 * The family aliases of an installed face: a font file spells its family in
 * several languages, and the canonical (en-US) spelling is the one the browser
 * resolves (genoffice#2449).
 *
 * The faces here are synthetic: a name table with an en-US and a zh-CN family
 * record for the same face, built from bytes (no binary fixture).
 */
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { familyAliases, indexFontDirectory } from '../src/sfnt'
import type { FontIndex } from '../src/sfnt'
import { buildNameTable, buildSfnt, type NameRecordSpec } from './helpers/synthetic-font'

const EN_US = 0x0409
const ZH_CN = 0x0804

/**
 * A font directory holding one font file whose name table spells the family in
 * the given languages.
 */
function fontDir(families: readonly { text: string; language: number }[]): string {
  const records: NameRecordSpec[] = []
  for (const family of families) {
    // the legacy family name and the typographic family, as most Windows faces do
    for (const nameId of [1, 16]) {
      records.push({
        platform: 3,
        encoding: 1,
        language: family.language,
        nameId,
        text: family.text,
      })
    }
  }
  const dir = mkdtempSync(join(tmpdir(), 'font-metrics-aliases-'))
  writeFileSync(join(dir, 'crafted.ttf'), buildSfnt([['name', buildNameTable(records)]]))
  return dir
}

/** the font index of a directory holding the given font files */
function indexOf(dir: string): FontIndex {
  const index: FontIndex = { byPs: new Map(), byFamily: new Map() }
  indexFontDirectory(dir, index)
  return index
}

describe('familyAliases', () => {
  it('a localized family name yields the canonical English spelling', () => {
    const index = indexOf(
      fontDir([
        { text: 'FZFangSong-Z02', language: EN_US },
        { text: '方正仿宋_GBK', language: ZH_CN },
      ]),
    )
    expect(familyAliases('方正仿宋_GBK', index)).toEqual(['FZFangSong-Z02'])
  })

  it('the canonical spelling yields the localized one', () => {
    const index = indexOf(
      fontDir([
        { text: 'FZFangSong-Z02', language: EN_US },
        { text: '方正仿宋_GBK', language: ZH_CN },
      ]),
    )
    expect(familyAliases('FZFangSong-Z02', index)).toEqual(['方正仿宋_GBK'])
  })

  it('a face with a single spelling yields no alias', () => {
    const index = indexOf(fontDir([{ text: 'SimSun', language: ZH_CN }]))
    expect(familyAliases('SimSun', index)).toEqual([])
  })

  it('a name no installed face carries yields no alias', () => {
    const index = indexOf(
      fontDir([
        { text: 'FZFangSong-Z02', language: EN_US },
        { text: '方正仿宋_GBK', language: ZH_CN },
      ]),
    )
    expect(familyAliases('No Such Font Family ZZZ', index)).toEqual([])
    // the match is exact after normalization, never fuzzy
    expect(familyAliases('FZFangSong', index)).toEqual([])
  })

  it('normalizes the requested name the way the index keys it', () => {
    const index = indexOf(
      fontDir([
        { text: 'FZFangSong-Z02', language: EN_US },
        { text: '方正仿宋_GBK', language: ZH_CN },
      ]),
    )
    expect(familyAliases('fzfang-song z02', index)).toEqual(['方正仿宋_GBK'])
  })

  it('the canonical spelling comes first among several aliases', () => {
    // the zh-CN record comes first in the name table: the canonical spelling
    // is still the English one, and it leads the alias list
    const index = indexOf(
      fontDir([
        { text: '仿宋_GB2312', language: ZH_CN },
        { text: 'FangSong_GB2312', language: EN_US },
        { text: 'FangSong_JP', language: 0x0411 },
      ]),
    )
    expect(familyAliases('FangSong_JP', index)).toEqual(['FangSong_GB2312', '仿宋_GB2312'])
    expect(familyAliases('仿宋_GB2312', index)).toEqual(['FangSong_GB2312', 'FangSong_JP'])
  })

  it('the canonical spelling is the one the English record spells', () => {
    // the zh-CN record comes first in the name table: the canonical spelling
    // is still the English one
    const index = indexOf(
      fontDir([
        { text: '仿宋_GB2312', language: ZH_CN },
        { text: 'FangSong_GB2312', language: EN_US },
      ]),
    )
    expect(familyAliases('仿宋_GB2312', index)).toEqual(['FangSong_GB2312'])
    expect(familyAliases('FangSong_GB2312', index)).toEqual(['仿宋_GB2312'])
  })
})
