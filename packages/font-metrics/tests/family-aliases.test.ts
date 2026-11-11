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

/** the legacy family name (1) and the typographic family (16) */
const LEGACY_FAMILY = 1
const TYPOGRAPHIC_FAMILY = 16

/**
 * A font directory holding one font file whose name table spells the family in
 * the given languages.
 */
function fontDir(families: readonly { text: string; language: number }[]): string {
  const records: NameRecordSpec[] = []
  for (const family of families) {
    // the legacy family name and the typographic family, as most Windows faces do
    for (const nameId of [LEGACY_FAMILY, TYPOGRAPHIC_FAMILY]) {
      records.push({
        platform: 3,
        encoding: 1,
        language: family.language,
        nameId,
        text: family.text,
      })
    }
  }
  return fontDirOf([{ fileName: 'crafted.ttf', records }])
}

/** A font directory holding the given font files. */
function fontDirOf(
  files: readonly { fileName: string; records: readonly NameRecordSpec[] }[],
): string {
  const dir = mkdtempSync(join(tmpdir(), 'font-metrics-aliases-'))
  for (const file of files) {
    writeFileSync(join(dir, file.fileName), buildSfnt([['name', buildNameTable(file.records)]]))
  }
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
        { text: '\u65b9\u6b63\u4eff\u5b8b_GBK', language: ZH_CN },
      ]),
    )
    expect(familyAliases('\u65b9\u6b63\u4eff\u5b8b_GBK', index)).toEqual(['FZFangSong-Z02'])
  })

  it('the canonical spelling yields the localized one', () => {
    const index = indexOf(
      fontDir([
        { text: 'FZFangSong-Z02', language: EN_US },
        { text: '\u65b9\u6b63\u4eff\u5b8b_GBK', language: ZH_CN },
      ]),
    )
    expect(familyAliases('FZFangSong-Z02', index)).toEqual(['\u65b9\u6b63\u4eff\u5b8b_GBK'])
  })

  it('a face with a single spelling yields no alias', () => {
    const index = indexOf(fontDir([{ text: 'SimSun', language: ZH_CN }]))
    expect(familyAliases('SimSun', index)).toEqual([])
  })

  it('a name no installed face carries yields no alias', () => {
    const index = indexOf(
      fontDir([
        { text: 'FZFangSong-Z02', language: EN_US },
        { text: '\u65b9\u6b63\u4eff\u5b8b_GBK', language: ZH_CN },
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
        { text: '\u65b9\u6b63\u4eff\u5b8b_GBK', language: ZH_CN },
      ]),
    )
    expect(familyAliases('fzfang-song z02', index)).toEqual(['\u65b9\u6b63\u4eff\u5b8b_GBK'])
  })

  it('the canonical spelling comes first among several aliases', () => {
    // the zh-CN record comes first in the name table: the canonical spelling
    // is still the English one, and it leads the alias list
    const index = indexOf(
      fontDir([
        { text: '\u4eff\u5b8b_GB2312', language: ZH_CN },
        { text: 'FangSong_GB2312', language: EN_US },
        { text: 'FangSong_JP', language: 0x0411 },
      ]),
    )
    expect(familyAliases('FangSong_JP', index)).toEqual(['FangSong_GB2312', '\u4eff\u5b8b_GB2312'])
    expect(familyAliases('\u4eff\u5b8b_GB2312', index)).toEqual(['FangSong_GB2312', 'FangSong_JP'])
  })

  it('the canonical spelling is the one the English record spells', () => {
    // the zh-CN record comes first in the name table: the canonical spelling
    // is still the English one
    const index = indexOf(
      fontDir([
        { text: '\u4eff\u5b8b_GB2312', language: ZH_CN },
        { text: 'FangSong_GB2312', language: EN_US },
      ]),
    )
    expect(familyAliases('\u4eff\u5b8b_GB2312', index)).toEqual(['FangSong_GB2312'])
    expect(familyAliases('FangSong_GB2312', index)).toEqual(['\u4eff\u5b8b_GB2312'])
  })

  it('a sibling weight face contributes no alias of its own family', () => {
    // The shape a CJK family uses: every weight face spells the typographic
    // family the same way ("Hiragino Sans GB") while the legacy family name
    // carries the weight ("Hiragino Sans GB W3"). The two weight faces below
    // share the typographic family and differ in the legacy one.
    const weightFace = (weight: string, family: string, localized: string): NameRecordSpec[] => [
      // the Macintosh record Apple writes next to the Windows ones: family-wide
      // (no weight), platform 1 — it must not pull the Windows weight names in
      {
        platform: 1,
        encoding: 0,
        language: 0,
        nameId: LEGACY_FAMILY,
        text: family,
      },
      {
        platform: 3,
        encoding: 1,
        language: EN_US,
        nameId: LEGACY_FAMILY,
        text: `${family} ${weight}`,
      },
      {
        platform: 3,
        encoding: 1,
        language: EN_US,
        nameId: TYPOGRAPHIC_FAMILY,
        text: family,
      },
      {
        platform: 3,
        encoding: 1,
        language: ZH_CN,
        nameId: LEGACY_FAMILY,
        text: `${localized} ${weight}`,
      },
      {
        platform: 3,
        encoding: 1,
        language: ZH_CN,
        nameId: TYPOGRAPHIC_FAMILY,
        text: localized,
      },
    ]
    const index = indexOf(
      fontDirOf([
        {
          fileName: 'w3.ttf',
          records: weightFace(
            'W3',
            'Hiragino Sans GB',
            '\u51ac\u9752\u9ed1\u4f53\u7b80\u4f53\u4e2d\u6587',
          ),
        },
        {
          fileName: 'w6.ttf',
          records: weightFace(
            'W6',
            'Hiragino Sans GB',
            '\u51ac\u9752\u9ed1\u4f53\u7b80\u4f53\u4e2d\u6587',
          ),
        },
      ]),
    )
    // the family name resolves through its own spelling only: the sibling
    // weight's legacy name is a different family, not another spelling
    expect(familyAliases('\u51ac\u9752\u9ed1\u4f53\u7b80\u4f53\u4e2d\u6587', index)).toEqual([
      'Hiragino Sans GB',
    ])
    expect(familyAliases('Hiragino Sans GB', index)).toEqual([
      '\u51ac\u9752\u9ed1\u4f53\u7b80\u4f53\u4e2d\u6587',
    ])
    // the round-4 leak: with the platform-1 record in the fixture, a name-id
    // match alone would return the sibling faces' "Hiragino Sans GB W3"/"W6"
    const canonicalAliases = familyAliases('Hiragino Sans GB', index)
    expect(canonicalAliases.some((alias) => /W\d$/.test(alias))).toBe(false)
    // the weight-qualified legacy name is its own slot
    expect(familyAliases('\u51ac\u9752\u9ed1\u4f53\u7b80\u4f53\u4e2d\u6587 W3', index)).toEqual([
      'Hiragino Sans GB W3',
    ])
    expect(familyAliases('Hiragino Sans GB W6', index)).toEqual([
      '\u51ac\u9752\u9ed1\u4f53\u7b80\u4f53\u4e2d\u6587 W6',
    ])
  })
})
