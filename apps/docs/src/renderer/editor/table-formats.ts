import type { StyleParaProps, StyleRunProps, StyleUpsert } from '@genoffice/docx-engine'
import type { IndexTablesKind } from './index-tables'

/** a built-in paragraph style the entries of a generated table are written in */
export interface EntryStyle {
  styleId: string
  /** the name Word gives the built-in style in styles.xml */
  name: string
}

const numbered = (id: string, name: string): EntryStyle[] =>
  Array.from({ length: 9 }, (_, i) => ({ styleId: `${id}${i + 1}`, name: `${name} ${i + 1}` }))

/** the styles Modify lists on each tab of Index and Tables */
export const ENTRY_STYLES: Record<IndexTablesKind, EntryStyle[]> = {
  toc: numbered('TOC', 'toc'),
  index: [{ styleId: 'IndexHeading', name: 'index heading' }, ...numbered('Index', 'index')],
  figures: [{ styleId: 'TableofFigures', name: 'table of figures' }],
  authorities: [
    { styleId: 'TOAHeading', name: 'toa heading' },
    { styleId: 'TableofAuthorities', name: 'table of authorities' },
  ],
}

export type TableFormat =
  | 'template'
  | 'classic'
  | 'distinctive'
  | 'fancy'
  | 'modern'
  | 'bulleted'
  | 'centered'
  | 'formal'
  | 'simple'

/** the formats each tab lists, in Word's order */
export const TABLE_FORMATS: Record<IndexTablesKind, TableFormat[]> = {
  index: ['template', 'classic', 'fancy', 'modern', 'bulleted', 'formal', 'simple'],
  toc: ['template', 'classic', 'distinctive', 'fancy', 'modern', 'formal', 'simple'],
  figures: ['template', 'classic', 'distinctive', 'centered', 'formal', 'simple'],
  authorities: ['template', 'classic', 'distinctive', 'formal', 'simple'],
}

interface Look {
  rPr?: StyleRunProps
  pPr?: StyleParaProps
}

// every look names the same properties, so one format fully replaces another
const look = (r: StyleRunProps = {}, p: StyleParaProps = {}): Look => ({
  rPr: {
    bold: false,
    italic: false,
    caps: false,
    smallCaps: false,
    underline: false,
    sizeHalfPoints: null,
    ...r,
  },
  pPr: { align: null, spaceBeforeTwips: 0, spaceAfterTwips: 0, ...p },
})

const PLAIN = look()

interface FormatDef {
  /** by style id; a style the format does not name keeps its definition */
  looks: Record<string, Look>
  /** index only: the line above each letter's entries (INDEX \h); null for none */
  heading?: string | null
}

/**
 * What a format other than "From template" makes of the entry styles. Word
 * redefines the styles too; the looks themselves are not measured from Word.
 */
const FORMATS: Record<IndexTablesKind, Partial<Record<TableFormat, FormatDef>>> = {
  toc: {
    classic: {
      looks: {
        TOC1: look({ bold: true, caps: true }, { spaceBeforeTwips: 120, spaceAfterTwips: 120 }),
        TOC2: look({ smallCaps: true }),
        TOC3: look({ italic: true }),
      },
    },
    distinctive: {
      looks: {
        TOC1: look({ bold: true, sizeHalfPoints: 24 }, { spaceBeforeTwips: 240 }),
        TOC2: look({ bold: true }, { spaceBeforeTwips: 60 }),
        TOC3: PLAIN,
      },
    },
    fancy: {
      looks: {
        TOC1: look(
          { bold: true, caps: true },
          { align: 'center', spaceBeforeTwips: 240, spaceAfterTwips: 120 },
        ),
        TOC2: look({ smallCaps: true }),
        TOC3: look({ italic: true }),
      },
    },
    modern: {
      looks: {
        TOC1: look({ bold: true, italic: true, sizeHalfPoints: 24 }, { spaceBeforeTwips: 240 }),
        TOC2: look({ smallCaps: true }),
        TOC3: PLAIN,
      },
    },
    formal: {
      looks: {
        TOC1: look({ caps: true }, { spaceBeforeTwips: 120 }),
        TOC2: look({ smallCaps: true }),
        TOC3: look({ italic: true }),
      },
    },
    simple: { looks: { TOC1: PLAIN, TOC2: PLAIN, TOC3: PLAIN } },
  },
  index: {
    classic: {
      heading: 'A',
      looks: {
        IndexHeading: look(
          { bold: true },
          { align: 'center', spaceBeforeTwips: 240, spaceAfterTwips: 120 },
        ),
      },
    },
    fancy: {
      heading: 'A',
      looks: {
        IndexHeading: look(
          { bold: true, italic: true, sizeHalfPoints: 28 },
          { spaceBeforeTwips: 240, spaceAfterTwips: 120 },
        ),
      },
    },
    modern: {
      heading: 'A',
      looks: {
        IndexHeading: look(
          { bold: true, sizeHalfPoints: 26 },
          { spaceBeforeTwips: 240, spaceAfterTwips: 60 },
        ),
      },
    },
    bulleted: {
      heading: 'A',
      looks: { IndexHeading: look({ bold: true }, { spaceBeforeTwips: 240 }) },
    },
    formal: {
      heading: 'A',
      looks: {
        IndexHeading: look({ bold: true, caps: true }, { spaceBeforeTwips: 240 }),
      },
    },
    simple: { heading: null, looks: { IndexHeading: PLAIN } },
  },
  figures: {
    classic: { looks: { TableofFigures: look({ smallCaps: true }) } },
    distinctive: { looks: { TableofFigures: look({ italic: true }) } },
    centered: { looks: { TableofFigures: look({}, { align: 'center' }) } },
    formal: { looks: { TableofFigures: look({ caps: true }) } },
    simple: { looks: { TableofFigures: PLAIN } },
  },
  authorities: {
    classic: {
      looks: {
        TOAHeading: look(
          { bold: true, caps: true },
          { spaceBeforeTwips: 240, spaceAfterTwips: 120 },
        ),
        TableofAuthorities: PLAIN,
      },
    },
    distinctive: {
      looks: {
        TOAHeading: look(
          { bold: true, underline: true },
          { spaceBeforeTwips: 240, spaceAfterTwips: 120 },
        ),
        TableofAuthorities: PLAIN,
      },
    },
    formal: {
      looks: {
        TOAHeading: look(
          { bold: true, caps: true },
          { align: 'center', spaceBeforeTwips: 240, spaceAfterTwips: 120 },
        ),
        TableofAuthorities: PLAIN,
      },
    },
    simple: {
      looks: {
        TOAHeading: look({ bold: true }, { spaceBeforeTwips: 120 }),
        TableofAuthorities: PLAIN,
      },
    },
  },
}

export type StyleNames = ReadonlyMap<string, { name: string }>

/** the id the document knows a built-in entry style by: ours, else the one of the same name */
export function entryStyleId(style: EntryStyle, styles: StyleNames | undefined): string {
  if (!styles || styles.has(style.styleId)) return style.styleId
  for (const [id, info] of styles) if (info.name.toLowerCase() === style.name) return id
  return style.styleId
}

const BUILT_IN = new Map(
  Object.values(ENTRY_STYLES)
    .flat()
    .map((style) => [style.styleId, style]),
)

/** the id a generated line's built-in style has in this document (`TOC1` may be stored under another id) */
export function documentStyleId(styleId: string, styles: StyleNames | undefined): string {
  const style = BUILT_IN.get(styleId)
  return style ? entryStyleId(style, styles) : styleId
}

/** a definition for a built-in entry style: one the document lacks is created as Word names it */
export function entryStyleUpsert(
  style: EntryStyle,
  patch: Pick<StyleUpsert, 'rPr' | 'pPr'>,
  styles: StyleNames | undefined,
  defaultParagraphStyleId: string | undefined,
): StyleUpsert {
  const styleId = entryStyleId(style, styles)
  if (styles?.has(styleId)) return { styleId, ...patch }
  return {
    styleId,
    type: 'paragraph',
    name: style.name,
    builtIn: true,
    quickFormat: false,
    ...(defaultParagraphStyleId ? { basedOn: defaultParagraphStyleId } : {}),
    ...patch,
  }
}

/** the style definitions a format writes; none for "From template" */
export function formatStyleUpserts(
  kind: IndexTablesKind,
  format: TableFormat,
  styles: StyleNames | undefined,
  defaultParagraphStyleId: string | undefined,
): StyleUpsert[] {
  const looks = FORMATS[kind][format]?.looks ?? {}
  return ENTRY_STYLES[kind]
    .filter((style) => looks[style.styleId])
    .map((style) => entryStyleUpsert(style, looks[style.styleId], styles, defaultParagraphStyleId))
}

/** index formats: the letter heading the format brings (null: none); undefined for "From template" */
export const formatIndexHeading = (format: TableFormat): string | null | undefined =>
  format === 'template' ? undefined : (FORMATS.index[format]?.heading ?? null)
