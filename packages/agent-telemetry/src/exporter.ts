import { appendFile } from "node:fs/promises";
import path from "node:path";
import type { RecordedTelemetryEvent, RecordedTelemetrySpan } from "@earendil-works/pi-telemetry";

/**
 * A span exporter consumes completed spans. Exporters are synchronous from the
 * exporter's perspective: they receive one {@link RecordedTelemetrySpan} at a
 * time and may do whatever they wish with it (write to disk, ship over the
 * network, log to console).
 *
 * GenOffice ships three concrete exporters:
 * - {@link JsonlFileExporter} for the local-file trace store
 * - {@link ConsoleExporter} for dev / debug
 * - {@link InMemoryExporter} for tests
 *
 * Custom exporters (e.g. an OpenTelemetry collector bridge) only need to
 * implement this interface.
 */
export interface SpanExporter {
	readonly name: string;
	exportSpan(span: RecordedTelemetrySpan): void | Promise<void>;
	/** Optional flush hook called when a recording session is shutting down. */
	flush?(): void | Promise<void>;
}

/* ------------------------------------------------------------------ */
/* In-memory exporter (tests).                                          */
/* ------------------------------------------------------------------ */

export interface InMemoryExporterOptions {
	/** Detach snapshots on `getSpans()` so callers cannot mutate the recorder. */
	readonly?: boolean;
}

export class InMemoryExporter implements SpanExporter {
	readonly name = "in-memory";
	private readonly spans: RecordedTelemetrySpan[] = [];
	private readonly readonly: boolean;

	constructor(options: InMemoryExporterOptions = {}) {
		this.readonly = options.readonly ?? true;
	}

	exportSpan(span: RecordedTelemetrySpan): void {
		this.spans.push(span);
	}

	getSpans(): readonly RecordedTelemetrySpan[] {
		return this.readonly ? this.spans.slice() : this.spans;
	}

	flush(): void {
		// No-op: in-memory exporter does not buffer.
	}

	reset(): void {
		this.spans.length = 0;
	}
}

/* ------------------------------------------------------------------ */
/* Console exporter (dev / debug).                                      */
/* ------------------------------------------------------------------ */

export interface ConsoleExporterOptions {
	/** Override the sink (mainly for tests). Defaults to `console.log`. */
	sink?: (line: string) => void;
	/** Pretty-print instead of compact JSON. */
	pretty?: boolean;
}

export class ConsoleExporter implements SpanExporter {
	readonly name = "console";
	private readonly sink: (line: string) => void;
	private readonly pretty: boolean;

	constructor(options: ConsoleExporterOptions = {}) {
		this.sink = options.sink ?? ((line: string) => console.log(line));
		this.pretty = options.pretty ?? false;
	}

	exportSpan(span: RecordedTelemetrySpan): void {
		const line = this.pretty ? JSON.stringify(span, null, 2) : JSON.stringify(span);
		this.sink(line);
	}

	flush(): void {
		// No-op: console output is unbuffered.
	}
}

/* ------------------------------------------------------------------ */
/* JSONL file exporter (production default).                            */
/* ------------------------------------------------------------------ */

export interface JsonlFileExporterOptions {
	/** Absolute path of the JSONL file. Defaults to `~/.genoffice/ai-traces.jsonl`. */
	filePath?: string;
	/** Override the writer (mainly for tests). Defaults to `fs/promises.appendFile`. */
	write?: (line: string) => void | Promise<void>;
	/** Whether to include child events as separate JSONL lines. Defaults to true. */
	includeEvents?: boolean;
}

export class JsonlFileExporter implements SpanExporter {
	readonly name = "jsonl-file";
	private readonly filePath: string;
	private readonly write: (line: string) => void | Promise<void>;
	private readonly includeEvents: boolean;

	constructor(options: JsonlFileExporterOptions = {}) {
		this.filePath = options.filePath ?? DEFAULT_TRACE_FILE;
		this.write = options.write ?? defaultFileWriter(this.filePath);
		this.includeEvents = options.includeEvents ?? true;
	}

	async exportSpan(span: RecordedTelemetrySpan): Promise<void> {
		await this.write(JSON.stringify(span));
		if (this.includeEvents) {
			for (const event of span.events) {
				await this.write(JSON.stringify({ kind: "event", spanId: span.id, ...event }));
			}
		}
	}
}

function defaultFileWriter(filePath: string): (line: string) => Promise<void> {
	return async (line: string): Promise<void> => {
		await appendFile(filePath, line + "\n", "utf8");
	};
}

/* ------------------------------------------------------------------ */
/* Multi exporter (fan-out).                                            */
/* ------------------------------------------------------------------ */

export class CompositeExporter implements SpanExporter {
	readonly name = "composite";
	private readonly exporters: SpanExporter[];

	constructor(exporters: readonly SpanExporter[]) {
		this.exporters = [...exporters];
	}

	async exportSpan(span: RecordedTelemetrySpan): Promise<void> {
		await Promise.all(this.exporters.map((exporter) => exporter.exportSpan(span)));
	}

	async flush(): Promise<void> {
		await Promise.all(
			this.exporters.map(async (exporter) => {
				if (exporter.flush) await exporter.flush();
			}),
		);
	}

	getExporters(): readonly SpanExporter[] {
		return this.exporters;
	}
}

/* ------------------------------------------------------------------ */
/* Helpers.                                                             */
/* ------------------------------------------------------------------ */

export const DEFAULT_TRACE_DIRECTORY = path.join(
	process.env.HOME ?? path.join(path.sep, "tmp"),
	".genoffice",
);
export const DEFAULT_TRACE_FILE = path.join(DEFAULT_TRACE_DIRECTORY, "ai-traces.jsonl");

/** Coerce a RecordedTelemetrySpan to a short summary string for log lines. */
export function summarizeSpan(span: RecordedTelemetrySpan): string {
	const status = span.status.status;
	return `${span.name}#${span.id} ${status}`;
}

/** Extract every event from a span tree in start order. */
export function collectEvents(spans: readonly RecordedTelemetrySpan[]): RecordedTelemetryEvent[] {
	const events: RecordedTelemetryEvent[] = [];
	for (const span of spans) {
		for (const event of span.events) events.push(event);
	}
	return events;
}
