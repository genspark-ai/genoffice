/**
 * sheets-skill extension — GenOffice Excel/sheets AI tools as a pi extension.
 *
 * Mirrors `docs-skill.ts` shape:
 *   - `SheetsEditor` interface abstracts the workbook runtime (Univer in production,
 *     mock in tests)
 *   - One `defineTool()` per migrated tool
 *   - `createSheetsSkillExtension()` is the entry point for createOfficeSession
 *
 * Migration strategy:
 *   - Phase 1 (W7): ship the 4 most-called read tools (workbook_context, read_range,
 *     aggregate_range, find_cells) + 1 representative write tool (create_document)
 *   - Phase 2 (post-W7): add the remaining read/write tools as needed
 */

import { Type, type Static } from "@earendil-works/pi-ai";
import { defineTool, type ExtensionAPI } from "@earendil-works/pi-coding-agent";
import type { ReactUIAdapter } from "@genoffice/agent-runtime";

/** Helper: build a TypeBox string enum (union of literals). */
const StringEnum = <T extends readonly string[]>(values: T, opts?: object) =>
  Type.Union(values.map((v) => Type.Literal(v)) as never, opts as never);

// ============================================================================
// SheetsEditor contract
// ============================================================================

export interface SheetsRange {
  /** First row (0-indexed). */
  startRow: number;
  /** Last row (0-indexed, inclusive). */
  endRow: number;
  /** First column (0-indexed). */
  startCol: number;
  /** Last column (0-indexed, inclusive). */
  endCol: number;
  /** Sheet name. */
  sheet: string;
}

export interface CellValue {
  /** Raw value (string, number, boolean, or formula string starting with "="). */
  raw: string | number | boolean | null;
  /** Display format (e.g. "0.00%", "$#,##0.00"). */
  format?: string;
}

export interface WorkbookSummary {
  sheetNames: string[];
  activeSheet: string;
  totalCells: number;
  totalFormulas: number;
}

export interface SheetsEditor {
  // ---- Read API ----
  /** Get a high-level summary of the workbook (sheets, formula count, etc.). */
  getWorkbookSummary(): WorkbookSummary;
  /** Read cell values for a range as a 2D array. */
  readRange(range: SheetsRange): CellValue[][];
  /** Aggregate (sum/avg/count/min/max) over a range. */
  aggregateRange(range: SheetsRange, op: "sum" | "avg" | "count" | "min" | "max"): number | null;
  /** Find cells matching a substring (case-insensitive); returns A1-style refs. */
  findCells(sheet: string, query: string, maxResults?: number): string[];
  /** Get features (merged cells, frozen panes, data validation) for a sheet. */
  getSheetFeatures(sheet: string): { mergedRanges: SheetsRange[]; frozenPanes: { row: number; col: number } | null };

  // ---- Write API ----
  /** Create a new document (delegated to the host shell). */
  createNewDocument?(sheetName: string): void;
}

// ============================================================================
// get_workbook_context
// ============================================================================

const GetWorkbookContextParams = Type.Object({});

export function createGetWorkbookContextTool(opts: { uiAdapter: ReactUIAdapter }) {
  const { uiAdapter } = opts;
  return defineTool<typeof GetWorkbookContextParams, { sheetCount: number; totalCells: number }>({
    name: "get_workbook_context",
    label: "Get Workbook Context",
    description:
      "Get the current workbook summary: sheet names, active sheet, total cell count, formula count. Call this first to know which sheet to operate on.",
    promptSnippet: "get_workbook_context() — workbook summary (sheets, formula count)",
    parameters: GetWorkbookContextParams,
    async execute(_id, _params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<SheetsEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No workbook available." }], details: { sheetCount: 0, totalCells: 0 } };
      }
      const summary = editor.getWorkbookSummary();
      return {
        content: [{
          type: "text",
          text: `Workbook: ${summary.sheetNames.length} sheet(s) [${summary.sheetNames.join(", ")}], active="${summary.activeSheet}", ${summary.totalCells} cells, ${summary.totalFormulas} formulas.`,
        }],
        details: { sheetCount: summary.sheetNames.length, totalCells: summary.totalCells },
      };
    },
  });
}

// ============================================================================
// read_range
// ============================================================================

const ReadRangeParams = Type.Object({
  sheet: Type.String({ description: "Sheet name (case-sensitive)" }),
  startRow: Type.Integer({ minimum: 0 }),
  startCol: Type.Integer({ minimum: 0 }),
  endRow: Type.Integer({ minimum: 0 }),
  endCol: Type.Integer({ minimum: 0 }),
});

type ReadRangeArgs = Static<typeof ReadRangeParams>;

const MAX_RANGE_CELLS = 5_000;

export function createReadRangeTool(opts: { uiAdapter: ReactUIAdapter }) {
  const { uiAdapter } = opts;
  return defineTool<typeof ReadRangeParams, { rows: number; cols: number; truncated: boolean }>({
    name: "read_range",
    label: "Read Range",
    description:
      "Read a rectangular range of cells. Returns the values as a 2D array (row-major). Large ranges are truncated; call again with a smaller range to continue.",
    promptSnippet: "read_range(sheet, startRow, startCol, endRow, endCol) — read cell values",
    parameters: ReadRangeParams,
    async execute(_id, params: ReadRangeArgs, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<SheetsEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No workbook available." }], details: { rows: 0, cols: 0, truncated: false } };
      }
      const rows = params.endRow - params.startRow + 1;
      const cols = params.endCol - params.startCol + 1;
      if (rows * cols > MAX_RANGE_CELLS) {
        return {
          content: [{ type: "text", text: `Range too large (${rows * cols} cells > ${MAX_RANGE_CELLS}). Narrow the range.` }],
          details: { rows, cols, truncated: true },
        };
      }
      const range: SheetsRange = { sheet: params.sheet, startRow: params.startRow, endRow: params.endRow, startCol: params.startCol, endCol: params.endCol };
      const data = editor.readRange(range);
      // Render as TSV (tab-separated values) for easy LLM consumption
      const tsv = data.map((row) => row.map((c) => String(c.raw ?? "")).join("\t")).join("\n");
      return {
        content: [{ type: "text", text: `${range.sheet}!R${params.startRow + 1}C${params.startCol + 1}:R${params.endRow + 1}C${params.endCol + 1}\n${tsv}` }],
        details: { rows, cols, truncated: false },
      };
    },
  });
}

// ============================================================================
// aggregate_range
// ============================================================================

const AggregateRangeParams = Type.Object({
  sheet: Type.String(),
  startRow: Type.Integer({ minimum: 0 }),
  startCol: Type.Integer({ minimum: 0 }),
  endRow: Type.Integer({ minimum: 0 }),
  endCol: Type.Integer({ minimum: 0 }),
  op: StringEnum(["sum", "avg", "count", "min", "max"] as const, { description: "Aggregation function" }),
});

export function createAggregateRangeTool(opts: { uiAdapter: ReactUIAdapter }) {
  const { uiAdapter } = opts;
  return defineTool<typeof AggregateRangeParams, { result: number | null; op: string }>({
    name: "aggregate_range",
    label: "Aggregate Range",
    description:
      "Aggregate a range of numeric cells. op is one of: sum, avg, count, min, max. Non-numeric cells are ignored for sum/avg/min/max; count includes them.",
    promptSnippet: "aggregate_range(sheet, range, op) — sum/avg/count/min/max over a range",
    parameters: AggregateRangeParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<SheetsEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No workbook available." }], details: { result: null, op: params.op } };
      }
      const range: SheetsRange = { sheet: params.sheet, startRow: params.startRow, endRow: params.endRow, startCol: params.startCol, endCol: params.endCol };
      const result = editor.aggregateRange(range, params.op);
      return {
        content: [{ type: "text", text: `${params.op}(${params.sheet}!R${params.startRow + 1}C${params.startCol + 1}:R${params.endRow + 1}C${params.endCol + 1}) = ${result ?? "(empty)"}` }],
        details: { result, op: params.op },
      };
    },
  });
}

// ============================================================================
// find_cells
// ============================================================================

const FindCellsParams = Type.Object({
  sheet: Type.String(),
  query: Type.String({ description: "Substring to search for (case-insensitive)" }),
  maxResults: Type.Optional(Type.Integer({ minimum: 1, description: "Max number of results (default 20)" })),
});

export function createFindCellsTool(opts: { uiAdapter: ReactUIAdapter }) {
  const { uiAdapter } = opts;
  return defineTool<typeof FindCellsParams, { count: number; cells: string[] }>({
    name: "find_cells",
    label: "Find Cells",
    description:
      "Find cells in a sheet whose value contains the query substring (case-insensitive). Returns A1-style cell references.",
    promptSnippet: "find_cells(sheet, query) — locate cells by substring",
    parameters: FindCellsParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<SheetsEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No workbook available." }], details: { count: 0, cells: [] } };
      }
      const cells = editor.findCells(params.sheet, params.query, params.maxResults ?? 20);
      return {
        content: [{ type: "text", text: cells.length === 0 ? "No matches." : `Found ${cells.length} cell(s): ${cells.join(", ")}` }],
        details: { count: cells.length, cells },
      };
    },
  });
}

// ============================================================================
// create_document (sheets variant)
// ============================================================================

const CreateSheetParams = Type.Object({
  sheetName: Type.Optional(Type.String({ description: "Initial sheet name (default: 'Sheet1')" })),
});

export function createCreateSheetTool(opts: { uiAdapter: ReactUIAdapter }) {
  const { uiAdapter } = opts;
  return defineTool<typeof CreateSheetParams, { created: boolean; sheetName: string }>({
    name: "create_document",
    label: "Create Workbook",
    description: "Create a new workbook (delegated to the host shell).",
    promptSnippet: "create_document(sheetName?) — open a new workbook",
    parameters: CreateSheetParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<SheetsEditor>();
      if (!editor || typeof editor.createNewDocument !== "function") {
        return {
          content: [{ type: "text", text: "create_document requires the host app to implement createNewDocument()." }],
          details: { created: false, sheetName: params.sheetName ?? "Sheet1" },
        };
      }
      const sheetName = params.sheetName ?? "Sheet1";
      editor.createNewDocument(sheetName);
      return {
        content: [{ type: "text", text: `Created new workbook with sheet "${sheetName}".` }],
        details: { created: true, sheetName },
      };
    },
  });
}

// ============================================================================
// Extension factory
// ============================================================================

export type SheetsToolName = "get_workbook_context" | "read_range" | "aggregate_range" | "find_cells" | "create_document";

export const ALL_SHEETS_TOOL_NAMES: readonly SheetsToolName[] = [
  "get_workbook_context",
  "read_range",
  "aggregate_range",
  "find_cells",
  "create_document",
];

export interface SheetsSkillOptions {
  uiAdapter: ReactUIAdapter;
  enabledTools?: ReadonlyArray<SheetsToolName>;
}

export function createSheetsSkillExtension(opts: SheetsSkillOptions) {
  const { uiAdapter, enabledTools } = opts;
  const enabled = new Set(enabledTools ?? ALL_SHEETS_TOOL_NAMES);

  return (pi: ExtensionAPI) => {
    if (enabled.has("get_workbook_context")) pi.registerTool(createGetWorkbookContextTool({ uiAdapter }));
    if (enabled.has("read_range")) pi.registerTool(createReadRangeTool({ uiAdapter }));
    if (enabled.has("aggregate_range")) pi.registerTool(createAggregateRangeTool({ uiAdapter }));
    if (enabled.has("find_cells")) pi.registerTool(createFindCellsTool({ uiAdapter }));
    if (enabled.has("create_document")) pi.registerTool(createCreateSheetTool({ uiAdapter }));

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (pi as any).on("before_agent_start", async () => ({
      systemPromptAppend:
        "\n## Spreadsheet Editing Rules\n- Sheet names are case-sensitive. Always call get_workbook_context first if you don't know them.\n- Cell coordinates are 0-based in tool args but 1-based in display (R1C1 = first cell).\n- read_range returns TSV (tab-separated); use it to understand data before mutating.",
    }));
  };
}

export default createSheetsSkillExtension;
