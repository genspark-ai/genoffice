import type { AgentToolDef } from '@genoffice/agent-core'

export const PAPER_SIZES: ReadonlyArray<{ name: string; w: number; h: number }> = [
  { name: 'A3', w: 16838, h: 23811 },
  { name: 'A4', w: 11906, h: 16838 },
  { name: 'A5', w: 8391, h: 11906 },
  { name: 'B5', w: 10319, h: 14572 },
  { name: 'Letter', w: 12240, h: 15840 },
  { name: 'Legal', w: 12240, h: 20160 },
  { name: 'Tabloid', w: 15840, h: 24480 },
]

export const LENGTH_HINT =
  'a number of twips or a string with a unit: "2.54cm", "1in", "72pt", "10mm"'

export const PAGE_SETUP_TOOL: AgentToolDef = {
  name: 'set_page_setup',
  description:
    'Change paper size, orientation, margins, text columns, the different-first-page flag or page numbering of the document or of one section (the message context lists the sections). ' +
    'Only the fields given change. Lengths take a unit suffix (cm, mm, in, pt) or twips. Setting the orientation alone swaps the current width and height.',
  inputSchema: {
    type: 'object',
    properties: {
      engineTarget: {
        type: 'object',
        description: 'Revision-bound section target from query_document.',
      },
      section: {
        type: 'integer',
        description:
          'section index (0-based) to change; omit with blockIndex omitted = every section',
      },
      blockIndex: {
        type: 'integer',
        description: 'change the section that contains this block index',
      },
      paper: {
        type: 'string',
        description: `named size: ${PAPER_SIZES.map((p) => p.name).join(', ')}`,
      },
      width: { type: ['string', 'number'], description: `custom page width, ${LENGTH_HINT}` },
      height: { type: ['string', 'number'], description: `custom page height, ${LENGTH_HINT}` },
      orientation: { type: 'string', enum: ['portrait', 'landscape'] },
      margins: {
        type: 'object',
        description: `any of top, right, bottom, left, header, footer, gutter; each ${LENGTH_HINT}`,
        properties: {
          top: { type: ['string', 'number'] },
          right: { type: ['string', 'number'] },
          bottom: { type: ['string', 'number'] },
          left: { type: ['string', 'number'] },
          header: { type: ['string', 'number'], description: 'header distance from the page top' },
          footer: {
            type: ['string', 'number'],
            description: 'footer distance from the page bottom',
          },
          gutter: { type: ['string', 'number'], description: 'extra binding margin' },
        },
      },
      columns: {
        type: 'object',
        description: 'text columns of the section',
        properties: {
          count: { type: 'integer', description: '1-12' },
          spacing: {
            type: ['string', 'number'],
            description: `gap between columns, ${LENGTH_HINT}`,
          },
        },
        required: ['count'],
      },
      titlePg: {
        type: 'boolean',
        description: 'different first page (own header/footer on the first page of the section)',
      },
      pageNumberStart: {
        type: ['integer', 'null'],
        description:
          'restart page numbering at this value; null = continue from the previous section',
      },
      pageNumberFormat: {
        type: 'string',
        description: 'decimal, lowerRoman, upperRoman, lowerLetter, upperLetter, ...',
      },
    },
  },
}

export const SECTION_BREAK_TOOL: AgentToolDef = {
  name: 'insert_section_break',
  description:
    'Insert a section break after a block so the blocks after it form a new section with its own page setup, headers and footers (the new section starts as a copy of the current one; change it with set_page_setup afterwards). ' +
    'type nextPage starts the new section on a new page, continuous keeps the flow (needed for column changes mid-page), evenPage / oddPage start on the next even / odd page.',
  inputSchema: {
    type: 'object',
    properties: {
      afterBlockIndex: {
        type: 'integer',
        description: 'the break goes after this block index; -1 = before the first block',
      },
      type: {
        type: 'string',
        enum: ['nextPage', 'continuous', 'evenPage', 'oddPage'],
        description: 'default nextPage',
      },
    },
    required: ['afterBlockIndex'],
  },
}
