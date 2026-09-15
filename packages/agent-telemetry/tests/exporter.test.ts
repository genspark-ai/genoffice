import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
	CompositeExporter,
	ConsoleExporter,
	InMemoryExporter,
	JsonlFileExporter,
	collectEvents,
	createRecordingTelemetry,
	summarizeSpan,
} from "../src/index";

let root: string;

beforeEach(async () => {
	root = await mkdtemp(path.join(os.tmpdir(), "agent-telemetry-"));
});

describe("InMemoryExporter", () => {
	it("records spans in start order", async () => {
		const exporter = new InMemoryExporter();
		const { context } = createRecordingTelemetry();
		// Note: InMemoryTelemetryContext.startSpan accepts a callback but does
		// not itself route completed spans to an exporter; the test asserts
		// the exporter records whatever we feed it directly.
		await context.startSpan({ name: "alpha" }, async (span) => {
			span.setAttributes({ value: 1 });
		});
		await context.startSpan({ name: "beta" }, async () => {
			// no-op
		});

		const alpha = {
			id: 1,
			parentId: null,
			name: "alpha",
			attributes: { value: 1 },
			events: [],
			status: { status: "ok" } as const,
			settled: true,
		};
		const beta = {
			id: 2,
			parentId: null,
			name: "beta",
			attributes: {},
			events: [],
			status: { status: "ok" } as const,
			settled: true,
		};

		exporter.exportSpan(alpha);
		exporter.exportSpan(beta);

		const spans = exporter.getSpans();
		expect(spans).toHaveLength(2);
		expect(spans[0]?.name).toBe("alpha");
		expect(spans[1]?.name).toBe("beta");
	});

	it("getSpans returns a snapshot by default", () => {
		const exporter = new InMemoryExporter();
		exporter.exportSpan({
			id: 1,
			parentId: null,
			name: "x",
			attributes: {},
			events: [],
			status: { status: "ok" } as const,
			settled: true,
		});
		const snapshot = exporter.getSpans();
		exporter.reset();
		expect(snapshot).toHaveLength(1);
		expect(exporter.getSpans()).toHaveLength(0);
	});

	it("reset() clears the recorder", () => {
		const exporter = new InMemoryExporter({ readonly: false });
		exporter.exportSpan({
			id: 1,
			parentId: null,
			name: "x",
			attributes: {},
			events: [],
			status: { status: "ok" } as const,
			settled: true,
		});
		exporter.reset();
		expect(exporter.getSpans()).toHaveLength(0);
	});
});

describe("ConsoleExporter", () => {
	it("writes each span as one JSON line via the provided sink", () => {
		const lines: string[] = [];
		const exporter = new ConsoleExporter({ sink: (line) => { lines.push(line); } });
		exporter.exportSpan({
			id: 1,
			parentId: null,
			name: "console-span",
			attributes: { app: "docs" },
			events: [],
			status: { status: "ok" } as const,
			settled: true,
		});
		expect(lines).toHaveLength(1);
		const parsed = JSON.parse(lines[0]!);
		expect(parsed.name).toBe("console-span");
		expect(parsed.attributes.app).toBe("docs");
	});
});

describe("JsonlFileExporter", () => {
	it("writes a span as a single line", async () => {
		const filePath = path.join(root, "traces.jsonl");
		const exporter = new JsonlFileExporter({ filePath });
		await exporter.exportSpan({
			id: 1,
			parentId: null,
			name: "file-span",
			attributes: {},
			events: [],
			status: { status: "ok" } as const,
			settled: true,
		});
		const contents = readFileSync(filePath, "utf8");
		const lines = contents.split("\n").filter((l) => l.length > 0);
		expect(lines).toHaveLength(1);
		expect(JSON.parse(lines[0]!).name).toBe("file-span");
	});

	it("appends subsequent spans without overwriting", async () => {
		const filePath = path.join(root, "traces.jsonl");
		const exporter = new JsonlFileExporter({ filePath });
		await exporter.exportSpan({
			id: 1,
			parentId: null,
			name: "first",
			attributes: {},
			events: [],
			status: { status: "ok" } as const,
			settled: true,
		});
		await exporter.exportSpan({
			id: 2,
			parentId: null,
			name: "second",
			attributes: {},
			events: [],
			status: { status: "ok" } as const,
			settled: true,
		});
		const contents = readFileSync(filePath, "utf8");
		const lines = contents.split("\n").filter((l) => l.length > 0);
		expect(lines).toHaveLength(2);
	});

	it("emits child events as separate JSONL lines when includeEvents is true", async () => {
		const filePath = path.join(root, "traces.jsonl");
		const exporter = new JsonlFileExporter({ filePath, includeEvents: true });
		await exporter.exportSpan({
			id: 1,
			parentId: null,
			name: "span-with-events",
			attributes: {},
			events: [
				{ name: "checkpoint", attributes: { step: 1 } },
				{ name: "checkpoint", attributes: { step: 2 } },
			],
			status: { status: "ok" } as const,
			settled: true,
		});
		const lines = readFileSync(filePath, "utf8").split("\n").filter((l) => l.length > 0);
		expect(lines).toHaveLength(3);
		expect(JSON.parse(lines[1]!).kind).toBe("event");
		expect(JSON.parse(lines[1]!).name).toBe("checkpoint");
		expect(JSON.parse(lines[2]!).attributes.step).toBe(2);
	});

	it("omits child events when includeEvents is false", async () => {
		const filePath = path.join(root, "traces.jsonl");
		const exporter = new JsonlFileExporter({ filePath, includeEvents: false });
		await exporter.exportSpan({
			id: 1,
			parentId: null,
			name: "span-with-events",
			attributes: {},
			events: [{ name: "checkpoint", attributes: {} }],
			status: { status: "ok" } as const,
			settled: true,
		});
		const lines = readFileSync(filePath, "utf8").split("\n").filter((l) => l.length > 0);
		expect(lines).toHaveLength(1);
	});

	it("accepts an injected write function (sync) instead of touching disk", () => {
		const lines: string[] = [];
		const exporter = new JsonlFileExporter({ write: (line) => { lines.push(line); } });
		exporter.exportSpan({
			id: 1,
			parentId: null,
			name: "in-memory",
			attributes: {},
			events: [],
			status: { status: "ok" } as const,
			settled: true,
		});
		expect(lines).toHaveLength(1);
	});
});

describe("CompositeExporter", () => {
	it("fans out exportSpan to every child exporter", async () => {
		const mem = new InMemoryExporter();
		const lines: string[] = [];
		const con = new ConsoleExporter({ sink: (line) => { lines.push(line); } });
		const composite = new CompositeExporter([mem, con]);
		await composite.exportSpan({
			id: 1,
			parentId: null,
			name: "fan-out",
			attributes: {},
			events: [],
			status: { status: "ok" } as const,
			settled: true,
		});
		expect(mem.getSpans()).toHaveLength(1);
		expect(lines).toHaveLength(1);
	});

	it("flush() invokes every child exporter's flush hook", async () => {
		let memFlushes = 0;
		let conFlushes = 0;
		const mem = new InMemoryExporter();
		mem.flush = async () => {
			memFlushes += 1;
		};
		const con = new ConsoleExporter();
		con.flush = async () => {
			conFlushes += 1;
		};
		const composite = new CompositeExporter([mem, con]);
		await composite.flush();
		expect(memFlushes).toBe(1);
		expect(conFlushes).toBe(1);
	});
});

describe("summarizeSpan", () => {
	it("returns the span name plus id and status", () => {
		const summary = summarizeSpan({
			id: 7,
			parentId: null,
			name: "office-ai.run",
			attributes: {},
			events: [],
			status: { status: "ok" } as const,
			settled: true,
		});
		expect(summary).toBe("office-ai.run#7 ok");
	});
});

describe("collectEvents", () => {
	it("returns every event from every span in order", () => {
		const events = collectEvents([
			{
				id: 1,
				parentId: null,
				name: "alpha",
				attributes: {},
				events: [{ name: "a1", attributes: {} }],
				status: { status: "ok" } as const,
				settled: true,
			},
			{
				id: 2,
				parentId: null,
				name: "beta",
				attributes: {},
				events: [
					{ name: "b1", attributes: {} },
					{ name: "b2", attributes: {} },
				],
				status: { status: "ok" } as const,
				settled: true,
			},
		]);
		expect(events.map((e) => e.name)).toEqual(["a1", "b1", "b2"]);
	});
});

describe("createRecordingTelemetry", () => {
	it("returns a TelemetryContext that records spans in memory", async () => {
		const { context } = createRecordingTelemetry();
		await context.startSpan({ name: "first" }, async (span) => {
			span.addEvent("begin");
			span.setAttributes({ app: "docs" });
		});
		await context.startSpan({ name: "second" }, async (span) => {
			span.setStatus({ status: "ok" });
		});
		const spans = context.getSpans();
		expect(spans.map((s) => s.name)).toEqual(["first", "second"]);
		expect(spans[0]?.events[0]?.name).toBe("begin");
		expect(spans[0]?.attributes.app).toBe("docs");
	});
});

afterEach(async () => {
	if (root) {
		await rm(root, { recursive: true, force: true }).catch(() => undefined);
	}
});
