import { z } from 'zod'
import { createFileViaCli } from '../headless-cli'
import { resolveOutputPath } from './document-tools'
import { generateExtension } from './formats'
import type { FamilyDriver, SessionHost } from './session-tools'
import type { McpToolDefinition } from '../mcp-server'
import type { CliRunner } from '../cli-runner'

/**
 * Sheets (xlsx) tool surface for the MCP server.
 *
 * Two paths, mirroring the docx/slides tools:
 * - headless `create_xlsx`: a row matrix handed to the bundled `genoffice` CLI
 *   (`create --type xlsx --from`), gated behind the "background generation"
 *   setting. The CLI writes through the app's xlsx gateway and evaluates
 *   formulas with the sidecar, so strings starting with "=" become real formula
 *   cells with cached values — more than the values-only writer this used to
 *   carry.
 * - a visible grid session: the tools drive a real sheets tab the user
 *   watches. The workbook lives in the renderer (Univer), so this needs a
 *   request/response bridge into the renderer (`sheets-bridge.ts` here, the
 *   renderer half at `apps/sheets/src/renderer/mcp-bridge.ts`) — the same
 *   pattern as the docs session.
 */

export interface SheetsToolDeps {
  /** directory generated files land in when the caller gives no path */
  defaultSaveDir: () => string
  /** expose the headless create_xlsx tool; default true — the shell passes the user's setting */
  background?: boolean
  /** visible-grid control; absent in headless/unit runs, which drops the session tools */
  sheets?: SheetsControl
  /** the bundled genoffice CLI, used by the headless tool */
  cli?: CliRunner
}

/**
 * Visible-grid control implemented in the shell main process
 * (`apps/shell/src/main/mcp/sheets-bridge.ts`): opens a blank sheets tab and
 * forwards commands to the renderer that owns the Univer workbook.
 */
export interface SheetsControl {
  /** open a fresh blank sheets tab; resolves to its webContents id once the renderer is ready */
  openBlankTab: () => Promise<number>
  /** run one command in that tab and resolve its result */
  runCommand: (
    wcId: number,
    command: 'apply_ops' | 'read_sheet' | 'save_sheet',
    payload: unknown,
  ) => Promise<unknown>
}

const XLSX_EXT = `.${generateExtension('xlsx')}`

/** the headless tool: a row matrix -> xlsx through the bundled CLI */
function createHeadlessXlsxTool(deps: SheetsToolDeps): McpToolDefinition {
  return {
    name: 'create_xlsx',
    description:
      'Create an Excel .xlsx file and save it to disk without opening the app UI. `data` is a 2D ' +
      'array of rows; numbers become numeric cells, and a string starting with "=" becomes a ' +
      'formula (the genoffice CLI evaluates it and stores the cached value). Returns the absolute ' +
      'path of the written file.',
    inputSchema: {
      title: z.string().describe('workbook title, used as the file name'),
      data: z
        .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
        .describe('rows of cell values; row 0 becomes spreadsheet row 1'),
      sheetName: z.string().optional().describe('name of the single sheet; default Sheet1'),
      path: z
        .string()
        .optional()
        .describe('absolute output path; default is a new file in the default save folder'),
      overwrite: z
        .boolean()
        .optional()
        .describe('allow replacing an existing file at `path`; default false'),
    },
    handler: async (args) => {
      const title = String(args.title ?? '').trim()
      if (!title) throw new Error('title must not be empty')
      if (!Array.isArray(args.data)) throw new Error('data must be a 2D array of rows')
      if (args.data.some((row) => !Array.isArray(row))) {
        throw new Error('data must be a 2D array of rows')
      }
      if (!deps.cli) throw new Error('headless generation is not available in this build')
      const sheetName =
        typeof args.sheetName === 'string' && args.sheetName.trim()
          ? args.sheetName.trim()
          : 'Sheet1'

      const targetPath = resolveOutputPath({
        defaultSaveDir: deps.defaultSaveDir,
        ext: XLSX_EXT,
        title,
        requestedPath: typeof args.path === 'string' ? args.path : undefined,
        overwrite: args.overwrite === true,
      })

      const rows = (args.data as unknown[][]).map((row) =>
        row.map((cell) => {
          if (cell === null || cell === undefined) return ''
          if (typeof cell === 'number') return Number.isFinite(cell) ? cell : ''
          return String(cell)
        }),
      )
      const { summary, outputPath } = await createFileViaCli(deps.cli, {
        type: 'xlsx',
        input: {
          name: 'table.json',
          content: JSON.stringify({ sheets: [{ name: sheetName, rows }] }),
        },
        out: targetPath,
        overwrite: args.overwrite === true,
      })
      return { path: outputPath, summary }
    },
  }
}

export function createSheetsTools(deps: SheetsToolDeps, host: SessionHost): McpToolDefinition[] {
  return [
    // headless generation is opt-in, same rule as create_docx/create_pptx
    ...(deps.background === false ? [] : [createHeadlessXlsxTool(deps)]),
    ...createGridContentTools(deps, host),
  ]
}

/**
 * The xlsx session lifecycle as seen by the shared create_session / save_session
 * tools. The content tools below address the tab this driver opened.
 */
export function sheetsDriver(sheets: SheetsControl): FamilyDriver {
  return {
    family: 'xlsx',
    openBlankTab: () => sheets.openBlankTab(),
    save: (wcId, path, overwrite) => sheets.runCommand(wcId, 'save_sheet', { path, overwrite }),
  }
}

/**
 * Visible-grid content tools: fill the spreadsheet the shared session opened, so
 * the user watches the grid take shape. The ops are the same zod-validated
 * workbook DSL the built-in AI uses (planFromOps + applyChangePlan), executed by
 * the tab's renderer.
 *
 * Only registered when the shell wired a SheetsControl — headless/unit runs
 * keep the file-only surface.
 */
function createGridContentTools(deps: SheetsToolDeps, host: SessionHost): McpToolDefinition[] {
  const sheets = deps.sheets
  if (!sheets) return []

  const requireActive = (): number => host.require('xlsx')

  return [
    {
      name: 'read_sheet',
      description:
        'Read the visible workbook: sheet names/ids with data extents, and optionally the current ' +
        'values/formulas of specific cells. Sheet ids and A1 addresses are what apply_sheet_ops targets.',
      inputSchema: {
        addresses: z
          .array(z.string())
          .optional()
          .describe(
            'A1 addresses to read (values + formulas); omit to get the workbook overview only',
          ),
        sheetId: z
          .string()
          .optional()
          .describe(
            'sheet to read from (id from a previous overview); default is the active sheet',
          ),
      },
      handler: async (args) => {
        const wc = requireActive()
        const payload = {
          ...(Array.isArray(args.addresses) ? { addresses: args.addresses.map(String) } : {}),
          ...(typeof args.sheetId === 'string' ? { sheetId: args.sheetId } : {}),
        }
        return sheets.runCommand(wc, 'read_sheet', payload)
      },
    },
    {
      name: 'apply_sheet_ops',
      description:
        'Apply workbook DSL operations to the visible spreadsheet (the same vocabulary the built-in ' +
        'AI uses). Common ops: set_cell, set_formula, clear_cell, set_range, clear_range, fill_range, ' +
        'copy_range, convert_to_values, insert_rows, delete_rows, insert_cols, delete_cols, add_sheet, ' +
        'delete_sheet, add_chart, add_table, set_filter, set_hyperlink, add_conditional_format, ' +
        'set_data_validation, set_note. Addresses are A1 on the target sheet. One batch applies as one ' +
        'undo step; a failed batch changes nothing. Use read_sheet for sheet ids and current values first.',
      inputSchema: {
        ops: z.array(z.any()).describe('array of workbook DSL op objects'),
        dryRun: z
          .boolean()
          .optional()
          .describe('plan the batch and report what would change, without modifying the grid'),
      },
      handler: async (args) => {
        const wc = requireActive()
        if (!Array.isArray(args.ops)) throw new Error('ops must be an array')
        const result = (await sheets.runCommand(wc, 'apply_ops', {
          ops: args.ops,
          dryRun: args.dryRun === true,
        })) as { ok?: boolean; reason?: string }
        // a rejected batch is a tool-level error so the caller reacts to it
        if (result?.ok === false) {
          throw new Error(result.reason ?? 'the batch could not be applied')
        }
        return result
      },
    },
  ]
}
