# Metric-compatible fallback fonts

Source: fonts bundled with LibreOffice (`/Applications/LibreOffice.app/Contents/Resources/fonts/truetype`),
freely redistributable with the app. Licenses: Carlito and Liberation are
**SIL Open Font License 1.1** (see `LICENSE-OFL.txt`); Caladea is
**Apache License 2.0** (copyright Huerta Tipografica).

| Font             | License    | Metric-compatible Word counterpart |
| ---------------- | ---------- | ---------------------------------- |
| Carlito GO       | OFL 1.1    | Calibri (Word's default body font) |
| Caladea          | Apache-2.0 | Cambria                            |
| Liberation Serif | OFL 1.1    | Times New Roman                    |
| Liberation Sans  | OFL 1.1    | Arial                              |
| Liberation Mono  | OFL 1.1    | Courier New                        |

"Carlito GO" (`Carlito-*.ttf`) is a derivative of Carlito 1.103: a build-time patch
(`tools/patch-carlito-vi.py`) rebuilds Vietnamese precomposed glyphs whose above mark
(circumflex/breve) was dropped (Ậ/Ệ/Ộ in Regular/Bold); advance widths are unchanged.
Renamed per OFL 1.1 §2 — "Carlito" is a Reserved Font Name.

Purpose: when a Word font declared by the document is missing on this machine, the
browser's silent fallback (Helvetica etc.) changes glyph widths, so line-break points
and pagination diverge from Word. Falling back to a metric-compatible font keeps
canvas line breaking aligned with Word, and stays consistent with the offline
pagination model (`tests/helpers/lo-fonts.ts` measures the same set of files).

Registration lives in `fonts.css`; family-name mapping in `cssFontFamily()` of `line-metrics.ts`.

## CJK fallback

| Font                                    | Role                                       |
| --------------------------------------- | ------------------------------------------ |
| Noto Sans CJK SC (GB2312-subset woff2)  | fallback for heiti-style (sans) families   |
| Noto Serif CJK SC (GB2312-subset woff2) | fallback for songti-style (serif) families |

Source: [notofonts/noto-cjk](https://github.com/notofonts/noto-cjk) (SIL OFL 1.1),
subset with fonttools to all 7,445 GB2312 Han characters + CJK punctuation/fullwidth
forms + basic Latin
(`pyftsubset --text-file=gb2312 --unicodes="U+0020-024F,U+2000-206F,U+3000-303F,U+FF00-FFEF" --flavor=woff2`).
Rare characters outside the subset still fall through to system fonts (shown as
missing glyphs in minimal environments); bold is synthesized by the browser.

The serif subset also backs the `GenOffice Fullwidth TC` face (`fonts.css`), a
unicode-range shim (U+FF0D/FF0F/FF3C/FF3F/FF5E) slotted before Songti TC in the
Traditional Chinese serif chain: Songti TC draws those fullwidth glyphs at
~0.2-0.5em of ink inside the 1em advance, so a PMingLiU document's U+FF0F
rendered as a spaced half-width slash. Real PMingLiU (Windows) still wins by
chain order; advances are 1.0em everywhere, so line breaking is unchanged.

## Korean fallback

| Font                              | Role                                             |
| --------------------------------- | ------------------------------------------------ |
| GenOffice Serif KR (subset woff2) | Batang-metric stand-in for Korean serif families |
| GenOffice Sans KR (subset woff2)  | fallback for Korean sans families (Malgun etc.)  |

Source: Noto Serif/Sans CJK KR Regular from [notofonts/noto-cjk](https://github.com/notofonts/noto-cjk)
(SIL OFL 1.1), subset with fonttools to the 2,350 KS X 1001 syllables + jamo
(U+1100-11FF, U+3130-318F) + basic Latin/CJK punctuation/fullwidth forms
(`U+0020-024F,U+2000-206F,U+3000-303F,U+FF00-FFEF`), then hmtx-normalized to the
metrics of the Windows faces Word substitutes for missing Korean fonts: hangul
syllables/compatibility jamo → 1.0em (Noto CJK KR ships 0.92/0.966em, which
would shift line breaks ~8% vs Word), serif digits → 0.596em and space →
0.333em (measured Batang values), sans Basic Latin (U+0020-007E, U+00A0) →
measured Malgun Gothic advances (space 0.352em, digits 0.551em; Noto's 0.224em
space alone drifted Korean sans line breaks ~3%/line —
`tools/normalize-kr-sans-hmtx.py`, asserted by `tests/kr-font-metrics.test.ts`).
Renamed because OFL reserves the "Noto" name
for unmodified builds. Conjoining jamo keep native advances (shaping). Word
counterpart line factors live in `lineHeightFactor()` of `line-metrics.ts`.

## Arabic fallback

| Font                             | Role                                                     |
| -------------------------------- | -------------------------------------------------------- |
| Noto Naskh Arabic (subset woff2) | fallback for naskh/serif-class Arabic families (default) |
| Noto Sans Arabic (subset woff2)  | fallback for kufi/sans-class Arabic families             |

Source: Noto Naskh/Sans Arabic Regular from [notofonts/arabic](https://github.com/notofonts/arabic)
(SIL OFL 1.1), subset with fonttools to the Arabic blocks + presentation forms +
digits/punctuation, keeping all shaping features
(`pyftsubset --unicodes="U+0020-024F,U+0600-06FF,U+0750-077F,U+08A0-08FF,U+FB50-FDFF,U+FE70-FEFF,U+2000-206F" --layout-features='*' --flavor=woff2`).
Names are kept ("Noto ..."): glyphs and advances are unmodified, so the OFL
Reserved Font Name clause does not apply. The upstream fonts carry no Latin
letters (only digits/punctuation); Latin text in a cs-font run falls through to
the rest of the chain. Word substitutes a missing Arabic font with a naskh-style
serif, so unknown Arabic families default to the Naskh chain.
