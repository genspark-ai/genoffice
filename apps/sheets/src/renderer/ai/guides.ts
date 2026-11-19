import writingGuide from './prompts/guides/writing.md?raw'
import formattingGuide from './prompts/guides/formatting.md?raw'
import layoutGuide from './prompts/guides/layout.md?raw'
import structureGuide from './prompts/guides/structure.md?raw'
import chartsGuide from './prompts/guides/charts.md?raw'
import dataGuide from './prompts/guides/data.md?raw'
import financialFormattingGuide from './prompts/guides/financial-formatting.md?raw'
import dataAttributionGuide from './prompts/guides/data-attribution.md?raw'
import tableGuide from './prompts/guides/table.md?raw'
import pivotGuide from './prompts/guides/pivot.md?raw'
import shapeImageGuide from './prompts/guides/shape-image.md?raw'

/**
 * On-demand prompt guides: the base system prompt stays small, the load_guide
 * tool's description carries this catalog (name + one-line summary), and the
 * model loads the domains it actually needs for the task at hand.
 */

export interface GuideEntry {
  readonly description: string
  readonly content: string
}

export const GUIDE_CATALOG: Readonly<Record<string, GuideEntry>> = {
  writing: {
    description:
      'Content writing: field definitions for set_cell/set_formula/set_range/clear_*, formula rules, share columns and value/unit conventions',
    content: writingGuide,
  },
  formatting: {
    description:
      'Formatting: all format_range fields (incl. borders), number formats, table style rules (headers/alignment/total rows/financial palette)',
    content: formattingGuide,
  },
  layout: {
    description:
      'Sorting & layout: sort_range semantics and limits, merge/unmerge, row heights/column widths, hiding rows/columns',
    content: layoutGuide,
  },
  structure: {
    description:
      'Structural changes: row/column insert-delete, sheet add/delete/duplicate/move/hide, sheet protection, two-batch discipline, automatic formula reference rewriting and #REF!',
    content: structureGuide,
  },
  charts: {
    description:
      'Charts & shapes: add_chart to create charts from data ranges, edit_chart to edit existing charts, add_shape to insert shapes/text boxes',
    content: chartsGuide,
  },
  data: {
    description:
      'Data tools: hyperlinks, AutoFilter, conditional formatting, data validation, defined names',
    content: dataGuide,
  },
  'financial-formatting': {
    description:
      'Financial statement formatting: currency locale locking, red parentheses for negatives, dash for zeros, hierarchy indentation, total rows, GAAP line-item order, accounting palette',
    content: financialFormattingGuide,
  },
  'data-attribution': {
    description:
      'Data source attribution: three tracing methods for filling sheets with external data (search/API/scraping) — Source column / footer rows / dedicated Sources sheet — plus the decision table',
    content: dataAttributionGuide,
  },
  table: {
    description:
      'Structured tables: add_table creates native Table objects (built-in styles/banding/filter) + add/delete_table_row/column row-column editing',
    content: tableGuide,
  },
  pivot: {
    description:
      'Pivot tables: add_pivot creates native PivotTables (real OOXML, recognized and refreshable in Excel) + refresh_pivot recomputes existing pivot tables',
    content: pivotGuide,
  },
  'shape-image': {
    description:
      'Shapes & images: add_shape shapes/text boxes, SPARKLINE in-cell mini charts, embedded IMAGE/HYPERLINK',
    content: shapeImageGuide,
  },
}

export const GUIDE_NAMES = Object.keys(GUIDE_CATALOG)

export function guideCatalogSummary(): string {
  return Object.entries(GUIDE_CATALOG)
    .map(([name, entry]) => `${name} — ${entry.description}`)
    .join('; ')
}

export function loadGuides(
  names: readonly string[],
): { ok: true; content: string } | { ok: false; error: string } {
  const unknown = names.filter((name) => !GUIDE_CATALOG[name])
  if (unknown.length > 0) {
    return {
      ok: false,
      error: `Unknown guide(s): ${unknown.join(', ')}. Available: ${GUIDE_NAMES.join(', ')}`,
    }
  }
  const content = names.map((name) => GUIDE_CATALOG[name]!.content.trim()).join('\n\n---\n\n')
  return { ok: true, content }
}
