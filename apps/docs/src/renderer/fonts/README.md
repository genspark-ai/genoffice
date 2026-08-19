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

### GenOffice Gothic KR

| Font                               | Role                                            |
| ---------------------------------- | ----------------------------------------------- |
| GenOffice Gothic KR (subset woff2) | real-metric face for NanumGothic-declaring docs |

Source: NanumGothic Regular from [google/fonts](https://github.com/google/fonts/tree/main/ofl/nanumgothic)
(SIL OFL 1.1). Word for Mac renders NanumGothic documents with the OS
_downloadable_ Nanum asset (FontServices subset Chromium cannot see): hangul
0.94em, space 0.28em, digits 0.606em (M3 probe 2026-08-14), while the
Batang-normalized subset above ships 1.0/0.333/0.596 — +6.4% per hangul line.
Subset to the same ranges as the KR fallbacks (KS X 1001 syllables + jamo +
Basic Latin/punctuation/fullwidth forms), advances **unmodified**
(`tools/build-gothic-kr-font.py`) and checked in as
`GenOfficeGothicKR-Regular-subset.woff2`. Renamed per OFL (the upstream
Reserved Font Names include "Nanum" and "NanumGothic"; subsetting is a
modification). The exact NHN copyright/Reserved Font Name notice and the full
OFL 1.1 text are in `LICENSE-OFL.txt`.

## Tamil fallback

| Font                    | Role                                     |
| ----------------------- | ---------------------------------------- |
| GenOffice Tamil (woff2) | Latha-metric stand-in for Tamil families |

Source: Noto Sans Tamil Regular from [notofonts](https://github.com/notofonts/notofonts.github.io)
(SIL OFL 1.1). Word substitutes missing Tamil families with Latha; Chromium's
macOS fallback (Tamil Sangam MN) shapes ~27% narrower (M3 probe: sentence R
0.728, space 0.39×), far past what size-adjust can fix without inflating glyph
ink (137%). Advances are rewritten to Latha's: the 109 cmap-shared codepoints
exactly, remaining glyphs (GSUB conjunct/matra outputs) by the median
Tamil-letter ratio (`tools/build-tamil-font.py`; shaped sentence R vs Latha
0.994, every probe sentence within ±2.3%). The face ships no Latin letters
(upstream Noto Sans Tamil has none); Latin falls through the chain. Renamed
per OFL ("Noto" is a Reserved Font Name; advances are modified).

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
