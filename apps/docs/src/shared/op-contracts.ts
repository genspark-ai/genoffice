export const FONT_KEYS = [
  'bold',
  'italic',
  'underline',
  'strike',
  'color',
  'highlight',
  'fontSize',
  'fontFamily',
  'baseline',
  'link',
] as const
export const PARA_KEYS = [
  'align',
  'lineSpacing',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'spaceBefore',
  'spaceAfter',
  'pageBreakBefore',
  'shadingFill',
  'borders',
  'tabs',
] as const
export const IMAGE_KEYS = ['widthPx', 'heightPx', 'align'] as const

const FONT_SIGNATURE_FIELDS =
  'bold?, italic?, underline?, strike?, color?: "#RRGGBB"|null, highlight?, fontSize?: pt|null, fontFamily?: string|null, baseline?: "superscript"|"subscript"|"none"|null, link?: {url}|null'

const rowSig = (name: string, extra: string) =>
  `{ op: "${name}", target, row, ${extra} }  // target must match one native table (nodeType "table" and/or blockIndexes); row is 0-based`
const colSig = (name: string, extra: string) =>
  `{ op: "${name}", target, col, ${extra} }  // target must match one native table; col is a 0-based grid column`

export const OP_CONTRACTS = {
  setFont: {
    name: 'setFont',
    signature: `{ op: "setFont", target, ${FONT_SIGNATURE_FIELDS} }  // run-level style of the targeted blocks (only the selected characters when the target is a partial selection); a present field is set, null clears it, absent fields stay; fontFamily applies to its own script slot (an East Asian font keeps the run's Latin font and vice versa)`,
    keys: FONT_KEYS,
    target: 'required',
  },
  setMatchedFont: {
    name: 'setMatchedFont',
    signature: `{ op: "setMatchedFont", text?, matchCase?, matchColor?: "#RRGGBB"|"auto", matchFontSize?: pt|"default", target?, ${FONT_SIGNATURE_FIELDS} }  // character-level: styles only the runs that match (not the whole block) — occurrences of text, and/or runs currently in matchColor / at matchFontSize (as shown by read_blocks spans); give at least one of text / matchColor / matchFontSize; target narrows the scanned blocks, omitted = whole document (table cells included)`,
    keys: ['text', 'matchCase', 'matchColor', 'matchFontSize', ...FONT_KEYS],
    target: 'optional',
  },
  setParagraphFormat: {
    name: 'setParagraphFormat',
    signature:
      '{ op: "setParagraphFormat", target, align?: "left"|"center"|"right"|"justify"|null, lineSpacing?, indentLeft?, indentRight?, indentFirstLine?, spaceBefore?, spaceAfter?, pageBreakBefore?: boolean, shadingFill?: "#RRGGBB"|null, borders?: subset of "tblr"|null, tabs?: [{ pos: twips|"6.5in", align?: "left"|"center"|"right"|"decimal"|"bar", leader?: "none"|"dot"|"hyphen"|"underscore"|"middleDot", clear?: true }]|null }  // paragraph format of the targeted blocks (whole blocks, even for a partial selection); lengths in twips, lineSpacing as a multiple; indentFirstLine positive = first-line indent, negative = hanging; a two-character CJK indent ≈ font size in pt × 40 twips; tabs replaces the paragraph\'s own stops ([] / null clears them, clear: true cancels a style stop); a tab character in HTML is <span data-tab></span>',
    keys: PARA_KEYS,
    target: 'required',
  },
  setHeadingLevel: {
    name: 'setHeadingLevel',
    signature: '{ op: "setHeadingLevel", target, level: 0-6 }  // 0 = demote to body paragraph',
    keys: ['level'],
    target: 'required',
    required: ['level'],
  },
  findReplace: {
    name: 'findReplace',
    signature:
      '{ op: "findReplace", find, replace, matchCase?, match?: "one"|"all", target? }  // literal logical-text replacement across formatting runs inside one paragraph; inherits match-start formatting, preserves outside text; default one rejects zero/multiple matches; all requires an explicit target scope',
    keys: ['find', 'replace', 'matchCase', 'match'],
    target: 'optional',
    required: ['find', 'replace'],
  },
  deleteBlocks: {
    name: 'deleteBlocks',
    signature: '{ op: "deleteBlocks", target }',
    keys: [],
    target: 'required',
  },
  moveBlocks: {
    name: 'moveBlocks',
    signature:
      '{ op: "moveBlocks", blockIndexes: number[], afterBlockIndex }  // move the blocks (relative order kept) after afterBlockIndex; -1 = start of document',
    keys: ['blockIndexes', 'afterBlockIndex'],
    target: 'none',
    required: ['blockIndexes', 'afterBlockIndex'],
  },
  setList: {
    name: 'setList',
    signature:
      '{ op: "setList", target, kind?: "bullet"|"number" }  // paragraphs to list items (default bullet); heading blocks are not converted',
    keys: ['kind'],
    target: 'required',
  },
  clearList: {
    name: 'clearList',
    signature: '{ op: "clearList", target }  // list items back to body paragraphs',
    keys: [],
    target: 'required',
  },
  setImageProperties: {
    name: 'setImageProperties',
    signature:
      '{ op: "setImageProperties", target, widthPx?, heightPx?, align?: "left"|"center"|"right"|null }  // image blocks only (target nodeType "image"); giving one dimension scales the other proportionally',
    keys: IMAGE_KEYS,
    target: 'required',
  },
  insertToc: {
    name: 'insertToc',
    signature:
      '{ op: "insertToc", afterBlockIndex }  // insert a TOC field after that block (-1 = start of document); entries come from the current headings, Word computes page numbers on open; fails if the document has no headings',
    keys: ['afterBlockIndex'],
    target: 'none',
    required: ['afterBlockIndex'],
  },
  setParagraphAttrs: {
    name: 'setParagraphAttrs',
    signature: '{ op: "setParagraphAttrs", target, attrs: Record<string, unknown> }',
    keys: ['attrs'],
    target: 'required',
    hidden: true,
  },
  setOutlineLevel: {
    name: 'setOutlineLevel',
    signature: '{ op: "setOutlineLevel", target, level: 0-9 }',
    keys: ['level'],
    target: 'required',
    hidden: true,
  },
  stepIndent: {
    name: 'stepIndent',
    signature: '{ op: "stepIndent", target, delta: 1|-1 }',
    keys: ['delta'],
    target: 'required',
    hidden: true,
  },
  stepHangingIndent: {
    name: 'stepHangingIndent',
    signature: '{ op: "stepHangingIndent", target, delta: 1|-1 }',
    keys: ['delta'],
    target: 'required',
    hidden: true,
  },
  insertField: {
    name: 'insertField',
    signature:
      '{ op: "insertField", target, type: "SEQ"|"REF"|"PAGEREF"|"DATE"|"TIME"|"MERGEFIELD"|"IF"|"PAGE"|"NUMPAGES"|"AUTHOR"|..., args?: string, result?: string, position?: "start"|"end", afterText?: string }  // insert a Word field into exactly one paragraph, at its end unless position or afterText says otherwise; args is the instruction after the keyword ({ SEQ Figure \\* ARABIC } = type "SEQ", args "Figure \\* ARABIC"). Computed here: SEQ (visible fields with that identifier so far + 1), REF/PAGEREF (the bookmarked paragraph; insertBookmark creates one), DATE/TIME (now, \\@ picture), MERGEFIELD («Name»), IF with literal operands; PAGE/NUMPAGES show "1"; anything else needs result, the text shown until Word updates it. Fields only Word can compute are marked dirty and refresh when the file opens in Word',
    keys: ['type', 'args', 'result', 'position', 'afterText'],
    target: 'required',
    required: ['type'],
  },
  insertBookmark: {
    name: 'insertBookmark',
    signature:
      '{ op: "insertBookmark", target, name }  // bookmark exactly one paragraph (name: letters, digits, underscore, up to 40 characters, unique) so REF / PAGEREF fields can point at it',
    keys: ['name'],
    target: 'required',
    required: ['name'],
  },
  updateFields: {
    name: 'updateFields',
    signature:
      '{ op: "updateFields", scope?: "all"|"SEQ"|"REF"|"DATE" }  // recompute the inline fields the editor can: SEQ numbering in document order, REF results from their bookmarks, DATE/TIME to now; with scope "all" every other field (PAGE, NUMPAGES, AUTHOR...) is marked dirty so Word recomputes it on open. Table-of-contents blocks are untouched (insertToc rebuilds one)',
    keys: ['scope'],
    target: 'none',
  },
  insertTableRow: {
    name: 'insertTableRow',
    signature:
      rowSig('insertTableRow', 'position?: "before"|"after", count?') +
      "; new rows copy the neighbouring row's cell shading/borders (default: after)",
    keys: ['row', 'position', 'count'],
    target: 'required',
    required: ['row'],
  },
  deleteTableRow: {
    name: 'deleteTableRow',
    signature: rowSig('deleteTableRow', 'count?') + '; cells merged across the row shrink',
    keys: ['row', 'count'],
    target: 'required',
    required: ['row'],
  },
  insertTableColumn: {
    name: 'insertTableColumn',
    signature:
      colSig('insertTableColumn', 'position?: "before"|"after", count?') +
      "; the new column takes the neighbouring column's width and cell formatting (default: after)",
    keys: ['col', 'position', 'count'],
    target: 'required',
    required: ['col'],
  },
  deleteTableColumn: {
    name: 'deleteTableColumn',
    signature: colSig('deleteTableColumn', 'count?') + '; cells merged across the column shrink',
    keys: ['col', 'count'],
    target: 'required',
    required: ['col'],
  },
  mergeTableCells: {
    name: 'mergeTableCells',
    signature:
      '{ op: "mergeTableCells", target, range: { row, col, rowEnd?, colEnd? } }  // merge the rectangle (0-based grid coordinates, ends inclusive) into one cell; refused when an existing merged cell crosses the rectangle edge; the texts are kept as separate paragraphs',
    keys: ['range'],
    target: 'required',
    required: ['range'],
  },
  splitTableCell: {
    name: 'splitTableCell',
    signature:
      '{ op: "splitTableCell", target, cell: { row, col } }  // undo a merge: the cell becomes one cell per grid slot it covered; the text stays in the top-left one',
    keys: ['cell'],
    target: 'required',
    required: ['cell'],
  },
  setTableCellText: {
    name: 'setTableCellText',
    signature:
      '{ op: "setTableCellText", target, range: { row, col, rowEnd?, colEnd? }, text, match?: "one"|"all" }  // replace ordinary cell text in one native table; zero-based grid coordinates, ends inclusive; covered merged coordinates name their physical owner once; default one physical cell, explicit all applies to the bounded range; empty text clears; objects/fields/revisions/nested tables are refused',
    keys: ['range', 'text', 'match'],
    target: 'required',
    required: ['range', 'text'],
  },
  setTableCellFormat: {
    name: 'setTableCellFormat',
    signature:
      '{ op: "setTableCellFormat", target, range?: { row, col, rowEnd?, colEnd? }, fill?: "#RRGGBB"|null, vAlign?: "top"|"center"|"bottom"|null, width?: twips|"2.5cm"|"1in"|"72pt"|"96px", borders?: { all?, top?, left?, bottom?, right?: { style?: "single"|"dashed"|"double"|"none", width?: pt, color?: "#RRGGBB" }|null }|null }  // cells in the range (omitted = whole table); width sets the grid columns the range covers; borders: null = no borders on the cell, a side null = that side off',
    keys: ['range', 'fill', 'vAlign', 'width', 'borders'],
    target: 'required',
  },
  setTableStyle: {
    name: 'setTableStyle',
    signature:
      '{ op: "setTableStyle", target, styleId?: string|null, look?: { firstRow?, lastRow?, firstColumn?, lastColumn?, bandedRows?, bandedColumns?: boolean } }  // styleId must be a table style defined in this document (docs read lists them as tableStyles); null removes the style; look toggles Word\'s Table Style Options',
    keys: ['styleId', 'look'],
    target: 'required',
  },
  applyStyle: {
    name: 'applyStyle',
    signature:
      '{ op: "applyStyle", target, styleId }  // paragraph style by id (or name) from styles.xml / define_style; a heading style turns the block into a heading of its level, another paragraph style turns a heading back into a paragraph; list items keep their numbering',
    keys: ['styleId'],
    target: 'required',
    required: ['styleId'],
  },
} as const

export const PUBLIC_DOCS_OPS = Object.values(OP_CONTRACTS).filter((op) => !('hidden' in op))

const string = { type: 'string' }
const boolean = { type: 'boolean' }
const number = { type: 'number' }
const integer = { type: 'integer' }
const index = { ...integer, minimum: 0 }
const blockIndexes = { type: 'array', items: index, minItems: 1 }
const oneOf = (...values: string[]) => ({ type: 'string', enum: values })
const nullable = (schema: Record<string, unknown>) => ({ anyOf: [schema, { type: 'null' }] })
const object = (properties: Record<string, unknown>, required: string[] = []) => ({
  type: 'object',
  properties,
  required,
})
const positions = object({ from: index, to: index }, ['from', 'to'])
const cell = object({ row: index, col: index }, ['row', 'col'])
const gridRange = object({ ...cell.properties, rowEnd: index, colEnd: index }, cell.required)
const target = {
  ...object({
    nodeType: oneOf('docHeading', 'docParagraph', 'docListItem', 'image', 'table'),
    headingLevel: { ...integer, minimum: 1, maximum: 6 },
    containsText: string,
    matchCase: boolean,
    blockIndexes,
    scope: oneOf('selection', 'document'),
    range: positions,
    engine: {
      ...object(
        {
          documentId: string,
          revision: index,
          range: positions,
          expectedType: string,
          expectedHtml: string,
          kind: string,
          cell,
          address: object(
            {
              section: index,
              kind: oneOf('header', 'footer'),
              view: oneOf('default', 'first', 'even'),
              textboxIndex: index,
            },
            ['section'],
          ),
        },
        ['documentId', 'revision', 'range', 'expectedType'],
      ),
      description: 'Copy a revision-bound target from query_document verbatim into target.engine.',
    },
  }),
  minProperties: 1,
  description:
    'Supply at least one selector. Selectors combine with AND. scope: "selection" uses the captured selection. blockIndexes are zero-based. Use { engine: item.target } to address a fresh query_document result.',
}
const patchNumber = nullable(number)
const patchString = nullable(string)
const align = nullable(oneOf('left', 'center', 'right', 'justify'))
const length = { anyOf: [number, string] }
const border = nullable(
  object({
    style: oneOf('single', 'dashed', 'double', 'none'),
    width: number,
    color: string,
  }),
)
const properties: Record<string, Record<string, unknown>> = {
  bold: boolean,
  italic: boolean,
  underline: boolean,
  strike: boolean,
  color: patchString,
  highlight: patchString,
  fontSize: patchNumber,
  fontFamily: patchString,
  baseline: nullable(oneOf('superscript', 'subscript', 'none')),
  link: nullable(object({ url: string }, ['url'])),
  text: string,
  matchCase: boolean,
  matchColor: string,
  matchFontSize: { anyOf: [number, { const: 'default' }] },
  align,
  lineSpacing: patchNumber,
  indentLeft: patchNumber,
  indentRight: patchNumber,
  indentFirstLine: patchNumber,
  spaceBefore: patchNumber,
  spaceAfter: patchNumber,
  pageBreakBefore: boolean,
  shadingFill: patchString,
  borders: patchString,
  tabs: nullable({
    type: 'array',
    items: object(
      {
        pos: length,
        align: oneOf('left', 'center', 'right', 'decimal', 'bar'),
        leader: oneOf('none', 'dot', 'hyphen', 'underscore', 'middleDot'),
        clear: boolean,
      },
      ['pos'],
    ),
  }),
  level: { ...integer, minimum: 0, maximum: 6 },
  find: string,
  replace: string,
  match: oneOf('one', 'all'),
  blockIndexes,
  afterBlockIndex: { ...integer, minimum: -1 },
  kind: oneOf('bullet', 'number'),
  widthPx: patchNumber,
  heightPx: patchNumber,
  type: string,
  args: string,
  result: string,
  position: oneOf('start', 'end'),
  afterText: string,
  name: string,
  scope: oneOf('all', 'SEQ', 'REF', 'DATE'),
  row: index,
  col: index,
  count: { ...integer, minimum: 1 },
  range: gridRange,
  cell,
  fill: patchString,
  vAlign: nullable(oneOf('top', 'center', 'bottom')),
  width: length,
  styleId: string,
  look: object({
    firstRow: boolean,
    lastRow: boolean,
    firstColumn: boolean,
    lastColumn: boolean,
    bandedRows: boolean,
    bandedColumns: boolean,
  }),
}
const tablePosition = { position: oneOf('before', 'after') }
const fieldOverrides: Record<string, Record<string, Record<string, unknown>>> = {
  setImageProperties: { align: nullable(oneOf('left', 'center', 'right')) },
  insertTableRow: tablePosition,
  insertTableColumn: tablePosition,
  setTableCellFormat: {
    borders: nullable(
      object({ all: border, top: border, bottom: border, left: border, right: border }),
    ),
  },
  setTableStyle: { styleId: patchString },
}

export const APPLY_OPS_ITEMS = {
  oneOf: PUBLIC_DOCS_OPS.map((op) => ({
    ...object(
      {
        op: { type: 'string', const: op.name },
        ...(op.target === 'none' ? {} : { target }),
        ...Object.fromEntries(
          op.keys.map((key) => [key, fieldOverrides[op.name]?.[key] ?? properties[key]]),
        ),
      },
      [
        'op',
        ...(op.target === 'required' ? ['target'] : []),
        ...('required' in op ? op.required : []),
      ],
    ),
    additionalProperties: false,
    description: op.signature,
  })),
}
export const APPLY_OPS_USAGE = [
  'Targets combine selectors with AND: nodeType (docHeading, docParagraph, docListItem, image, table), headingLevel, containsText, matchCase, blockIndexes (zero-based), scope ("selection" or "document"), or range ({from,to}). Use at least one selector. For a revision-bound target, pass {target: {engine: item.target}} using an item from query_document.',
  'Operation signatures (? marks optional fields):',
  ...PUBLIC_DOCS_OPS.map((op) => op.signature),
].join('\n')
