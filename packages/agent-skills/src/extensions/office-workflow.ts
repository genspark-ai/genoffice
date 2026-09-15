/**
 * office-workflow extension — orchestrates a multi-step task that crosses
 * Office apps (sheets → docs / slides) under a single agent tool call.
 *
 * Why this exists (parity with GenOffice legacy behaviour):
 *   The legacy `packages/agent-core/src/agent-loop.ts` had a private
 *   "cross_office_workflow" path that let a single prompt drive Excel data
 *   extraction → docx templating in one transaction. On pi we expose the
 *   same workflow as a regular tool; the actual file I/O is delegated to
 *   injected callbacks so the extension stays host-agnostic and unit-testable.
 *
 * Mechanism:
 *   - The host installs the extension via `installOfficeWorkflow(pi, opts)`.
 *   - `readSpreadsheet(path)` reads Excel cells (host wires in
 *     `@genoffice/xlsx-gateway` or whatever the host uses).
 *   - `composeDocument(templatePath, rows)` / `composeSlides(templatePath, rows, title)`
 *     render the output (host wires in `@genoffice/docx-engine` / `pptx-engine`).
 *   - The extension itself only orchestrates: it calls the readers / composers
 *     in order, validates the result, and reports back.
 */

import { Type, type Static } from "@earendil-works/pi-ai";
import { defineTool, type ExtensionAPI } from "@earendil-works/pi-coding-agent";

/** Helper: build a TypeBox string enum (union of literals). */
const StringEnum = <T extends readonly string[]>(values: T, opts?: object) =>
	Type.Union(values.map((v) => Type.Literal(v)) as never, opts as never);

/** Output formats supported by the cross-office workflow. */
export const WORKFLOW_OUTPUT_FORMATS = ["docx", "slides"] as const;
export type WorkflowOutputFormat = (typeof WORKFLOW_OUTPUT_FORMATS)[number];

/**
 * A single row of data extracted from the spreadsheet. We keep this structural
 * (plain `unknown` payload) so hosts can return whatever shape their reader
 * produces — strings, numbers, formula results, dates, etc.
 */
export interface WorkflowRow {
	rowIndex: number;
	cells: ReadonlyArray<unknown>;
}

/** Result of the spreadsheet read step. */
export interface SpreadsheetReadResult {
	rows: ReadonlyArray<WorkflowRow>;
	/** Absolute path of the source spreadsheet (echoed back). */
	sourcePath: string;
}

/** Result of the document composition step. */
export interface ComposeResult {
	outputPath: string;
	bytesWritten: number;
}

/**
 * Callbacks that bridge the workflow extension to host-side Office engines.
 * Hosts MUST supply all three. Tests supply fakes.
 */
export interface OfficeWorkflowCallbacks {
	readSpreadsheet(path: string): Promise<SpreadsheetReadResult>;
	composeDocument(templatePath: string, rows: ReadonlyArray<WorkflowRow>): Promise<ComposeResult>;
	composeSlides(
		templatePath: string,
		rows: ReadonlyArray<WorkflowRow>,
		title: string,
	): Promise<ComposeResult>;
}

/* ------------------------------------------------------------------ */
/* Parameter schema.                                                   */
/* ------------------------------------------------------------------ */

export const CrossOfficeWorkflowParams = Type.Object({
	spreadsheetPath: Type.String({ description: "Absolute path of the source Excel file" }),
	templateDocPath: Type.String({ description: "Absolute path of the docx / pptx template" }),
	outputFormat: StringEnum(WORKFLOW_OUTPUT_FORMATS, {
		description: "Output format: docx for a Word document, slides for a PowerPoint deck",
	}),
	slidesTitle: Type.Optional(
		Type.String({ description: "Deck title (used only when outputFormat is 'slides')" }),
	),
});

export type CrossOfficeWorkflowArgs = Static<typeof CrossOfficeWorkflowParams>;

export interface CrossOfficeWorkflowDetails {
	format: WorkflowOutputFormat;
	rowsProcessed: number;
	outputPath: string;
	bytesWritten: number;
}

export interface OfficeWorkflowOptions {
	callbacks: OfficeWorkflowCallbacks;
}

/* ------------------------------------------------------------------ */
/* Tool factory.                                                        */
/* ------------------------------------------------------------------ */

export function createOfficeWorkflowTool(opts: OfficeWorkflowOptions) {
	const { callbacks } = opts;

	return defineTool<typeof CrossOfficeWorkflowParams, CrossOfficeWorkflowDetails>({
		name: "cross_office_workflow",
		label: "Cross-Office Workflow",
		description:
			"Compose a workflow across multiple Office apps in a single call: read an Excel spreadsheet, then render either a docx document or a slides deck from a template. " +
			"Use this when the user asks for a multi-asset deliverable (e.g. 'make the quarterly report from this spreadsheet').",
		promptSnippet:
			"cross_office_workflow(spreadsheetPath, templateDocPath, outputFormat, slidesTitle?) — Excel → docx/slides, single transaction",
		promptGuidelines: [
			"Prefer this tool over chaining sheets + docs / sheets + slides tools when the deliverable is a single composed output.",
			"`outputFormat: docx` ignores `slidesTitle`; supply it only for slides output.",
			"The host's reader decides which sheet(s) to include; pass a spreadsheetPath the host knows how to read.",
		],
		parameters: CrossOfficeWorkflowParams,
		async execute(
			_toolCallId,
			params: CrossOfficeWorkflowArgs,
			_signal,
			_onUpdate,
			_ctx,
		) {
			const read = await callbacks.readSpreadsheet(params.spreadsheetPath);
			const rowsProcessed = read.rows.length;

			let composed: ComposeResult;
			if (params.outputFormat === "docx") {
				composed = await callbacks.composeDocument(params.templateDocPath, read.rows);
			} else {
				const title = params.slidesTitle?.trim() || "Quarterly Report";
				composed = await callbacks.composeSlides(params.templateDocPath, read.rows, title);
			}

			const details: CrossOfficeWorkflowDetails = {
				format: params.outputFormat,
				rowsProcessed,
				outputPath: composed.outputPath,
				bytesWritten: composed.bytesWritten,
			};

			return {
				content: [
					{
						type: "text",
						text:
							`Wrote ${params.outputFormat} with ${rowsProcessed} row(s) to ${composed.outputPath} ` +
							`(${composed.bytesWritten} bytes).`,
					},
				],
				details,
			};
		},
	});
}

/* ------------------------------------------------------------------ */
/* Extension installer.                                                 */
/* ------------------------------------------------------------------ */

export interface InstallOfficeWorkflowOptions extends OfficeWorkflowOptions {
	/** Optional human-readable name shown in extension listings. */
	extensionName?: string;
}

/**
 * Register the cross-office workflow tool on a pi extension API.
 * Idempotent in practice but pi will throw if the same name is registered twice.
 */
export function installOfficeWorkflow(pi: ExtensionAPI, opts: InstallOfficeWorkflowOptions): void {
	const tool = createOfficeWorkflowTool({ callbacks: opts.callbacks });
	pi.registerTool(tool);
}
