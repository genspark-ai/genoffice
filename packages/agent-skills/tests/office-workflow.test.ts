/**
 * Tests for the cross-office workflow extension (W13 deliverable).
 *
 * Verifies:
 *   - Tool schema, prompts, and labels are stable
 *   - readSpreadsheet → compose pipeline runs end-to-end for both output formats
 *   - slidesTitle defaults when omitted or blank
 *   - installOfficeWorkflow wires the tool into a pi ExtensionAPI
 *   - Errors from callbacks surface verbatim (no swallowing)
 *   - Empty spreadsheet still produces a (possibly empty) output
 */

import { describe, expect, it, vi } from "vitest";
import type { ExtensionAPI } from "@earendil-works/pi-coding-agent";
import {
	CrossOfficeWorkflowParams,
	createOfficeWorkflowTool,
	installOfficeWorkflow,
	WORKFLOW_OUTPUT_FORMATS,
	type ComposeResult,
	type OfficeWorkflowCallbacks,
	type SpreadsheetReadResult,
	type WorkflowRow,
} from "../src/extensions/office-workflow";

interface CallLog {
	read?: { path: string };
	docx?: { template: string; rowCount: number };
	slides?: { template: string; rowCount: number; title: string };
}

function makeCallbacks(overrides: Partial<OfficeWorkflowCallbacks> = {}): {
	callbacks: OfficeWorkflowCallbacks;
	log: CallLog;
} {
	const log: CallLog = {};
	const rows: WorkflowRow[] = [
		{ rowIndex: 0, cells: ["Q1", 1200] },
		{ rowIndex: 1, cells: ["Q2", 1500] },
		{ rowIndex: 2, cells: ["Q3", 1700] },
	];
	const read: SpreadsheetReadResult = { rows, sourcePath: "/abs/q.xlsx" };
	const compose: ComposeResult = { outputPath: "/abs/out.docx", bytesWritten: 4096 };

	const callbacks: OfficeWorkflowCallbacks = {
		async readSpreadsheet(path: string) {
			log.read = { path };
			return overrides.readSpreadsheet ? overrides.readSpreadsheet(path) : read;
		},
		async composeDocument(template: string, inRows: ReadonlyArray<WorkflowRow>) {
			log.docx = { template, rowCount: inRows.length };
			return overrides.composeDocument ? overrides.composeDocument(template, inRows) : compose;
		},
		async composeSlides(template: string, inRows: ReadonlyArray<WorkflowRow>, title: string) {
			log.slides = { template, rowCount: inRows.length, title };
			return overrides.composeSlides
				? overrides.composeSlides(template, inRows, title)
				: { ...compose, outputPath: "/abs/out.pptx" };
		},
	};
	return { callbacks, log };
}

describe("WORKFLOW_OUTPUT_FORMATS", () => {
	it("is the docx / slides tuple", () => {
		expect(WORKFLOW_OUTPUT_FORMATS).toEqual(["docx", "slides"]);
	});
});

describe("CrossOfficeWorkflowParams schema", () => {
	it("declares spreadsheetPath, templateDocPath, outputFormat, optional slidesTitle", () => {
		expect(CrossOfficeWorkflowParams.type).toBe("object");
		const properties = (CrossOfficeWorkflowParams as { properties: Record<string, unknown> })
			.properties;
		expect(Object.keys(properties).sort()).toEqual([
			"outputFormat",
			"slidesTitle",
			"spreadsheetPath",
			"templateDocPath",
		]);
	});
});

describe("createOfficeWorkflowTool", () => {
	it("declares a tool named cross_office_workflow with the expected shape", () => {
		const { callbacks } = makeCallbacks();
		const tool = createOfficeWorkflowTool({ callbacks });
		expect(tool.name).toBe("cross_office_workflow");
		expect(tool.label).toBe("Cross-Office Workflow");
		expect(tool.parameters).toBe(CrossOfficeWorkflowParams);
		expect(typeof tool.execute).toBe("function");
	});

	it("composes a docx from a spreadsheet when outputFormat=docx", async () => {
		const { callbacks, log } = makeCallbacks();
		const tool = createOfficeWorkflowTool({ callbacks });
		const result = await tool.execute(
			"call-1",
			{
				spreadsheetPath: "/abs/q.xlsx",
				templateDocPath: "/abs/template.docx",
				outputFormat: "docx",
			},
			undefined,
			undefined,
			undefined as never,
		);

		expect(log.read).toEqual({ path: "/abs/q.xlsx" });
		expect(log.docx).toEqual({ template: "/abs/template.docx", rowCount: 3 });
		expect(log.slides).toBeUndefined();

		const details = result.details;
		expect(details.format).toBe("docx");
		expect(details.rowsProcessed).toBe(3);
		expect(details.outputPath).toBe("/abs/out.docx");
		expect(details.bytesWritten).toBe(4096);
		const first = result.content[0];
		expect(first?.type).toBe("text");
		if (first?.type === "text") expect(first.text).toMatch(/docx with 3 row\(s\)/);
	});

	it("composes slides and honours a custom slidesTitle", async () => {
		const { callbacks, log } = makeCallbacks();
		const tool = createOfficeWorkflowTool({ callbacks });
		const result = await tool.execute(
			"call-2",
			{
				spreadsheetPath: "/abs/q.xlsx",
				templateDocPath: "/abs/template.pptx",
				outputFormat: "slides",
				slidesTitle: "Q3 Earnings",
			},
			undefined,
			undefined,
			undefined as never,
		);

		expect(log.slides).toEqual({
			template: "/abs/template.pptx",
			rowCount: 3,
			title: "Q3 Earnings",
		});
		expect(log.docx).toBeUndefined();

		const details = result.details;
		expect(details.format).toBe("slides");
		expect(details.outputPath).toBe("/abs/out.pptx");
	});

	it("defaults slidesTitle to 'Quarterly Report' when omitted", async () => {
		const { callbacks, log } = makeCallbacks();
		const tool = createOfficeWorkflowTool({ callbacks });
		await tool.execute(
			"call-3",
			{
				spreadsheetPath: "/abs/q.xlsx",
				templateDocPath: "/abs/template.pptx",
				outputFormat: "slides",
			},
			undefined,
			undefined,
			undefined as never,
		);
		expect(log.slides?.title).toBe("Quarterly Report");
	});

	it("defaults slidesTitle to 'Quarterly Report' when blank", async () => {
		const { callbacks, log } = makeCallbacks();
		const tool = createOfficeWorkflowTool({ callbacks });
		await tool.execute(
			"call-3b",
			{
				spreadsheetPath: "/abs/q.xlsx",
				templateDocPath: "/abs/template.pptx",
				outputFormat: "slides",
				slidesTitle: "   ",
			},
			undefined,
			undefined,
			undefined as never,
		);
		expect(log.slides?.title).toBe("Quarterly Report");
	});

	it("passes through zero rows from an empty spreadsheet", async () => {
		const { callbacks, log } = makeCallbacks({
			readSpreadsheet: async () => ({ rows: [], sourcePath: "/empty.xlsx" }),
		});
		const tool = createOfficeWorkflowTool({ callbacks });
		const result = await tool.execute(
			"call-4",
			{
				spreadsheetPath: "/empty.xlsx",
				templateDocPath: "/abs/template.docx",
				outputFormat: "docx",
			},
			undefined,
			undefined,
			undefined as never,
		);
		expect(log.docx?.rowCount).toBe(0);
		expect(result.details.rowsProcessed).toBe(0);
	});

	it("surfaces errors from readSpreadsheet verbatim", async () => {
		const boom = new Error("sheet 1 not found");
		const { callbacks } = makeCallbacks({
			readSpreadsheet: async () => {
				throw boom;
			},
		});
		const tool = createOfficeWorkflowTool({ callbacks });
		await expect(
			tool.execute(
				"call-5",
				{
					spreadsheetPath: "/abs/q.xlsx",
					templateDocPath: "/abs/template.docx",
					outputFormat: "docx",
				},
				undefined,
				undefined,
				undefined as never,
			),
		).rejects.toBe(boom);
	});

	it("surfaces errors from composeDocument verbatim", async () => {
		const boom = new Error("template corrupt");
		const { callbacks } = makeCallbacks({
			composeDocument: async () => {
				throw boom;
			},
		});
		const tool = createOfficeWorkflowTool({ callbacks });
		await expect(
			tool.execute(
				"call-6",
				{
					spreadsheetPath: "/abs/q.xlsx",
					templateDocPath: "/abs/template.docx",
					outputFormat: "docx",
				},
				undefined,
				undefined,
				undefined as never,
			),
		).rejects.toBe(boom);
	});

	it("does not call composeDocument when outputFormat is slides", async () => {
		const composeDoc = vi.fn();
		const composeSlides = vi.fn(async () => ({ outputPath: "/x.pptx", bytesWritten: 1 }));
		const callbacks: OfficeWorkflowCallbacks = {
			readSpreadsheet: async () => ({ rows: [], sourcePath: "/x" }),
			composeDocument: composeDoc,
			composeSlides,
		};
		const tool = createOfficeWorkflowTool({ callbacks });
		await tool.execute(
			"call-7",
			{
				spreadsheetPath: "/x",
				templateDocPath: "/t",
				outputFormat: "slides",
				slidesTitle: "T",
			},
			undefined,
			undefined,
			undefined as never,
		);
		expect(composeDoc).not.toHaveBeenCalled();
		expect(composeSlides).toHaveBeenCalledTimes(1);
	});
});

describe("installOfficeWorkflow", () => {
	it("registers a single cross_office_workflow tool on the ExtensionAPI", () => {
		const { callbacks } = makeCallbacks();
		const registerTool = vi.fn();
		const api = { registerTool } as unknown as ExtensionAPI;
		installOfficeWorkflow(api, { callbacks });
		expect(registerTool).toHaveBeenCalledTimes(1);
		const tool = registerTool.mock.calls[0]?.[0] as { name: string };
		expect(tool.name).toBe("cross_office_workflow");
	});
});
