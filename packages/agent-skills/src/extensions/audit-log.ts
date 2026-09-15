/**
 * audit-log extension — record every tool call (and its result) into a sink
 * for compliance / post-hoc analysis.
 *
 * Why this exists (parity with GenOffice legacy behaviour):
 *   Plan §5.3 (`enterprise audit`) requires that every tool call be logged
 *   with user, tool name, input, result, and timestamp. pi's `tool_call` /
 *   `tool_result` events fire from the same agent loop but in two halves; W15
 *   pairs them by `toolCallId` so the final record carries both input and
 *   result without losing anything in between.
 *
 * Mechanism:
 *   - Host installs the extension via `installAuditLog(pi, opts)`.
 *   - `pi.on("tool_call", ...)` writes a half-entry with toolName + input.
 *   - `pi.on("tool_result", ...)` fills in the result half on the same id.
 *   - `redactKeys` masks sensitive fields (e.g. `password`, `apiKey`)
 *     before they reach the sink.
 *   - The default sink is `InMemoryAuditSink` (tests). Hosts use
 *     `JsonlAuditSink` for the production JSONL trail.
 */

import { appendFile } from "node:fs/promises";
import path from "node:path";
import type { ExtensionAPI, ToolCallEvent, ToolResultEvent } from "@earendil-works/pi-coding-agent";

/* ------------------------------------------------------------------ */
/* Audit record + sinks.                                                */
/* ------------------------------------------------------------------ */

/** Default user-data directory for the audit trail: `~/.genoffice`. */
export const DEFAULT_AUDIT_DIRECTORY = path.join(process.env.HOME ?? path.join(path.sep, "tmp"), ".genoffice");

/** Default audit JSONL filename: `audit-log.jsonl`. */
export const DEFAULT_AUDIT_FILE = path.join(DEFAULT_AUDIT_DIRECTORY, "audit-log.jsonl");

/**
 * One audit record covers a complete `tool_call` + `tool_result` roundtrip.
 * `input` / `result` fields are redacted copies — never mutate the event
 * directly.
 */
export interface AuditLogEntry {
	id: string;
	startedAt: number;
	completedAt?: number;
	toolName: string;
	toolCallId: string;
	sessionId?: string;
	cwd: string;
	user?: string;
	input: unknown;
	result?: unknown;
	isError?: boolean;
}

export interface AuditSink {
	readonly name: string;
	append(entry: AuditLogEntry): void | Promise<void>;
	flush?(): void | Promise<void>;
}

/** In-memory sink used by tests and short-lived hosts. */
export class InMemoryAuditSink implements AuditSink {
	readonly name = "in-memory";
	private readonly records: AuditLogEntry[] = [];

	async append(entry: AuditLogEntry): Promise<void> {
		this.records.push(entry);
	}

	getEntries(): readonly AuditLogEntry[] {
		return this.records.slice();
	}

	reset(): void {
		this.records.length = 0;
	}
}

/** JSONL file sink used by production hosts. */
export interface JsonlAuditSinkOptions {
	filePath?: string;
	/**
	 * Override the writer (mainly for tests). Defaults to
	 * `fs/promises.appendFile(filePath, line + "\n")`.
	 */
	write?: (line: string) => void | Promise<void>;
}

export class JsonlAuditSink implements AuditSink {
	readonly name = "jsonl-file";
	private readonly filePath: string;
	private readonly write: (line: string) => void | Promise<void>;

	constructor(options: JsonlAuditSinkOptions = {}) {
		this.filePath = options.filePath ?? DEFAULT_AUDIT_FILE;
		this.write = options.write ?? defaultFileWriter(this.filePath);
	}

	async append(entry: AuditLogEntry): Promise<void> {
		await this.write(JSON.stringify(entry));
	}
}

function defaultFileWriter(filePath: string): (line: string) => Promise<void> {
	return async (line: string): Promise<void> => {
		await appendFile(filePath, line + "\n", "utf8");
	};
}

/** Fan-out sink that calls every child. */
export class CompositeAuditSink implements AuditSink {
	readonly name = "composite";
	private readonly sinks: AuditSink[];

	constructor(sinks: readonly AuditSink[]) {
		this.sinks = [...sinks];
	}

	async append(entry: AuditLogEntry): Promise<void> {
		await Promise.all(this.sinks.map((sink) => sink.append(entry)));
	}

	async flush(): Promise<void> {
		await Promise.all(
			this.sinks.map(async (sink) => {
				if (sink.flush) await sink.flush();
			}),
		);
	}
}

/* ------------------------------------------------------------------ */
/* Redaction helpers.                                                   */
/* ------------------------------------------------------------------ */

export const DEFAULT_REDACT_KEYS: readonly string[] = Object.freeze([
	"password",
	"passwd",
	"secret",
	"apiKey",
	"api_key",
	"token",
	"authorization",
	"cookie",
	"set-cookie",
]);

const REDACTED = "[REDACTED]";

/**
 * Deep-clone `value` and replace every field whose key matches `redactKeys`
 * (case-insensitive). Non-object values are returned as-is.
 */
export function redact<T>(value: T, redactKeys: readonly string[] = DEFAULT_REDACT_KEYS): T {
	const needle = new Set(redactKeys.map((k) => k.toLowerCase()));
	const visit = (input: unknown): unknown => {
		if (input === null || input === undefined) return input;
		if (Array.isArray(input)) return input.map(visit);
		if (typeof input !== "object") return input;
		const out: Record<string, unknown> = {};
		for (const [key, val] of Object.entries(input as Record<string, unknown>)) {
			if (needle.has(key.toLowerCase())) {
				out[key] = REDACTED;
			} else {
				out[key] = visit(val);
			}
		}
		return out;
	};
	return visit(value) as T;
}

/* ------------------------------------------------------------------ */
/* Extension installer.                                                 */
/* ------------------------------------------------------------------ */

export interface AuditLogOptions {
	sink: AuditSink;
	/** Extra / override keys to redact (case-insensitive). */
	redactKeys?: readonly string[];
	/** Optional resolver for the acting user. Defaults to `ctx.user` when present. */
	resolveUser?: (ctx: { cwd: string }) => string | undefined;
}

interface PendingHalf {
	toolName: string;
	toolCallId: string;
	sessionId?: string;
	cwd: string;
	user?: string;
	startedAt: number;
	input: unknown;
}

/**
 * Register `tool_call` + `tool_result` handlers that pair events by
 * `toolCallId` and emit one {@link AuditLogEntry} per roundtrip.
 */
export function installAuditLog(pi: ExtensionAPI, opts: AuditLogOptions): void {
	const redactKeys = opts.redactKeys ?? DEFAULT_REDACT_KEYS;
	const resolveUser = opts.resolveUser;
	const pending = new Map<string, PendingHalf>();

	pi.on("tool_call", async (event: ToolCallEvent, ctx: { cwd: string; user?: string }) => {
		const half: PendingHalf = {
			toolName: event.toolName,
			toolCallId: event.toolCallId,
			cwd: ctx.cwd,
			user: resolveUser?.(ctx) ?? ctx.user,
			startedAt: Date.now(),
			input: redact(event.input, redactKeys),
		};
		pending.set(event.toolCallId, half);
		// Emit the half-entry immediately so a crashed tool still has a record.
		await opts.sink.append({
			id: event.toolCallId,
			startedAt: half.startedAt,
			toolName: half.toolName,
			toolCallId: half.toolCallId,
			cwd: half.cwd,
			user: half.user,
			input: half.input,
		});
	});

	pi.on("tool_result", async (event: ToolResultEvent, ctx: { cwd: string }) => {
		const half = pending.get(event.toolCallId);
		if (!half) {
			// Result without a matching call (rare race). Log it as best-effort.
			await opts.sink.append({
				id: event.toolCallId,
				startedAt: Date.now(),
				completedAt: Date.now(),
				toolName: "<unknown>",
				toolCallId: event.toolCallId,
				cwd: ctx.cwd,
				input: redact(event.input, redactKeys),
				result: redact(extractText(event.content), redactKeys),
				isError: event.isError,
			});
			return;
		}
		pending.delete(event.toolCallId);
		await opts.sink.append({
			id: event.toolCallId,
			startedAt: half.startedAt,
			completedAt: Date.now(),
			toolName: half.toolName,
			toolCallId: half.toolCallId,
			sessionId: half.sessionId,
			cwd: half.cwd,
			user: half.user,
			input: half.input,
			result: redact(extractText(event.content), redactKeys),
			isError: event.isError,
		});
	});
}

function extractText(content: ReadonlyArray<{ type: string; text?: string }>): unknown {
	if (content.length === 1 && content[0]?.type === "text") {
		return content[0].text ?? "";
	}
	return content;
}
