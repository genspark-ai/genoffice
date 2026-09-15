import {
	InMemoryTelemetryContext,
	NOOP_TELEMETRY_CONTEXT,
	type SpanOptions,
	type TelemetryContext,
	type TelemetrySpan,
} from "@earendil-works/pi-telemetry";

import { InMemoryExporter } from "./exporter";

export {
	CompositeExporter,
	ConsoleExporter,
	InMemoryExporter,
	JsonlFileExporter,
	collectEvents,
	summarizeSpan,
	DEFAULT_TRACE_DIRECTORY,
	DEFAULT_TRACE_FILE,
	type ConsoleExporterOptions,
	type InMemoryExporterOptions,
	type JsonlFileExporterOptions,
	type SpanExporter,
} from "./exporter";

/**
 * Convenience wrapper around {@link TelemetryContext.startSpan} that flattens
 * the pi-telemetry generic into a one-liner. Returns the callback's return
 * value so callers can `await startSpan(ctx, opts, async () => ...)` directly.
 */
export async function startSpan<T>(
	context: TelemetryContext,
	options: SpanOptions,
	callback: (span: TelemetrySpan) => T | Promise<T>,
): Promise<T> {
	return context.startSpan(options, callback);
}

/**
 * Construct a recording telemetry context backed by an {@link InMemoryExporter}.
 * The returned object exposes the exporter for assertion in tests and the
 * underlying {@link TelemetryContext} for host code that wants to call
 * `startSpan` directly.
 */
export interface RecordingTelemetry {
	readonly context: InMemoryTelemetryContext;
	readonly exporter: InMemoryExporter;
}

export function createRecordingTelemetry(options: { readonly?: boolean } = {}): RecordingTelemetry {
	const exporter = new InMemoryExporter({ readonly: options.readonly });
	const context = new InMemoryTelemetryContext();
	return { context, exporter };
}

/** The shared no-op context for production hosts that opt out of telemetry. */
export const noopTelemetry: TelemetryContext = NOOP_TELEMETRY_CONTEXT;

/** Re-export the type so consumers do not need a direct dependency. */
export type { RecordedTelemetrySpan, SpanOptions, TelemetryContext, TelemetrySpan } from "@earendil-works/pi-telemetry";
