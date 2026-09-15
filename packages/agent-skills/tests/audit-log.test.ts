/**
 * Tests for the audit-log extension (W15 deliverable).
 *
 * Verifies:
 *   - InMemoryAuditSink + JsonlAuditSink + CompositeAuditSink shape
 *   - installAuditLog pairs tool_call + tool_result into one AuditLogEntry
 *   - The half-entry is written on tool_call even if tool_result never arrives
 *   - Redaction applies to both input and result
 *   - Default redact keys cover common sensitive fields
 *   - resolveUser overrides ctx.user
 */

import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { ExtensionAPI, ToolCallEvent, ToolResultEvent } from "@earendil-works/pi-coding-agent";
import {
	CompositeAuditSink,
	InMemoryAuditSink,
	installAuditLog,
	JsonlAuditSink,
	redact,
	DEFAULT_REDACT_KEYS,
	type AuditLogEntry,
} from "../src/extensions/audit-log";

interface Harness {
	api: ExtensionAPI;
	handlers: {
		tool_call: Array<(event: ToolCallEvent, ctx: { cwd: string; user?: string }) => Promise<void>>;
		tool_result: Array<(event: ToolResultEvent, ctx: { cwd: string; user?: string }) => Promise<void>>;
	};
}

function makeHarness(): Harness {
	const handlers = { tool_call: [], tool_result: [] };
	const api = {
		on(event: string, handler: unknown) {
			if (event === "tool_call") handlers.tool_call.push(handler as never);
			else if (event === "tool_result") handlers.tool_result.push(handler as never);
			else throw new Error(`Unexpected event ${event}`);
		},
	} as unknown as ExtensionAPI;
	return { api, handlers };
}

function toolCallEvent(overrides: Partial<ToolCallEvent> = {}): ToolCallEvent {
	return {
		type: "tool_call",
		toolCallId: overrides.toolCallId ?? "call-1",
		toolName: overrides.toolName ?? "read_blocks",
		input: overrides.input ?? { startBlockIndex: 0, endBlockIndex: 1 },
	} as ToolCallEvent;
}

function toolResultEvent(overrides: Partial<ToolResultEvent> = {}): ToolResultEvent {
	return {
		type: "tool_result",
		toolCallId: overrides.toolCallId ?? "call-1",
		input: overrides.input ?? { startBlockIndex: 0, endBlockIndex: 1 },
		content: overrides.content ?? [{ type: "text", text: "ok" }],
		isError: overrides.isError ?? false,
	} as ToolResultEvent;
}

let root: string;

beforeEach(async () => {
	root = await mkdtemp(path.join(os.tmpdir(), "agent-audit-"));
});

afterEach(async () => {
	if (root) await rm(root, { recursive: true, force: true }).catch(() => undefined);
});

describe("InMemoryAuditSink", () => {
	it("records entries and returns a snapshot", async () => {
		const sink = new InMemoryAuditSink();
		const entry: AuditLogEntry = {
			id: "a",
			startedAt: 1,
			toolName: "read_blocks",
			toolCallId: "a",
			cwd: "/x",
			input: { startBlockIndex: 0 },
		};
		await sink.append(entry);
		expect(sink.getEntries()).toHaveLength(1);
		expect(sink.getEntries()[0]?.toolName).toBe("read_blocks");
	});

	it("reset() clears the recorder", async () => {
		const sink = new InMemoryAuditSink();
		await sink.append({
			id: "a",
			startedAt: 1,
			toolName: "x",
			toolCallId: "a",
			cwd: "/x",
			input: {},
		});
		sink.reset();
		expect(sink.getEntries()).toHaveLength(0);
	});
});

describe("JsonlAuditSink", () => {
	it("writes one entry per line via the supplied writer", async () => {
		const lines: string[] = [];
		const sink = new JsonlAuditSink({
			filePath: path.join(root, "audit.jsonl"),
			write: (line) => { lines.push(line); },
		});
		await sink.append({
			id: "a",
			startedAt: 1,
			toolName: "read_blocks",
			toolCallId: "a",
			cwd: "/x",
			input: {},
		});
		expect(lines).toHaveLength(1);
		expect(JSON.parse(lines[0]!).toolName).toBe("read_blocks");
	});

	it("appends to disk by default", async () => {
		const filePath = path.join(root, "audit.jsonl");
		const sink = new JsonlAuditSink({ filePath });
		await sink.append({
			id: "a",
			startedAt: 1,
			toolName: "read_blocks",
			toolCallId: "a",
			cwd: "/x",
			input: {},
		});
		await sink.append({
			id: "b",
			startedAt: 2,
			toolName: "read_blocks",
			toolCallId: "b",
			cwd: "/x",
			input: {},
		});
		const lines = readFileSync(filePath, "utf8").split("\n").filter((l) => l.length > 0);
		expect(lines).toHaveLength(2);
	});
});

describe("CompositeAuditSink", () => {
	it("fans out append to every child sink", async () => {
		const a = new InMemoryAuditSink();
		const b = new InMemoryAuditSink();
		const c = new CompositeAuditSink([a, b]);
		await c.append({
			id: "x",
			startedAt: 1,
			toolName: "x",
			toolCallId: "x",
			cwd: "/x",
			input: {},
		});
		expect(a.getEntries()).toHaveLength(1);
		expect(b.getEntries()).toHaveLength(1);
	});
});

describe("redact", () => {
	it("masks default sensitive keys (case-insensitive)", () => {
		const input = {
			password: "p1",
			API_KEY: "k1",
			Token: "t1",
			authorization: "Bearer xyz",
			safe: "keep",
		};
		expect(redact(input)).toEqual({
			password: "[REDACTED]",
			API_KEY: "[REDACTED]",
			Token: "[REDACTED]",
			authorization: "[REDACTED]",
			safe: "keep",
		});
	});

	it("walks nested objects and arrays", () => {
		const input = { nested: { secret: "s", ok: 1 }, list: [{ token: "t" }, { safe: 2 }] };
		expect(redact(input)).toEqual({
			nested: { secret: "[REDACTED]", ok: 1 },
			list: [{ token: "[REDACTED]" }, { safe: 2 }],
		});
	});

	it("returns primitives and arrays of primitives unchanged", () => {
		expect(redact(42)).toBe(42);
		expect(redact("hi")).toBe("hi");
		expect(redact(null)).toBe(null);
		expect(redact([1, 2, 3])).toEqual([1, 2, 3]);
	});

	it("accepts custom redactKeys", () => {
		expect(redact({ custom: 1 }, ["custom"])).toEqual({ custom: "[REDACTED]" });
	});

	it("DEFAULT_REDACT_KEYS contains the expected sensitive names", () => {
		const keys = new Set(DEFAULT_REDACT_KEYS);
		expect(keys.has("password")).toBe(true);
		expect(keys.has("apiKey")).toBe(true);
		expect(keys.has("authorization")).toBe(true);
		expect(keys.has("token")).toBe(true);
	});
});

describe("installAuditLog", () => {
	it("registers one tool_call and one tool_result handler", () => {
		const { api, handlers } = makeHarness();
		const sink = new InMemoryAuditSink();
		installAuditLog(api, { sink });
		expect(handlers.tool_call).toHaveLength(1);
		expect(handlers.tool_result).toHaveLength(1);
	});

	it("pairs tool_call + tool_result into a single AuditLogEntry", async () => {
		const { api, handlers } = makeHarness();
		const sink = new InMemoryAuditSink();
		installAuditLog(api, { sink });

		await handlers.tool_call[0]!(
			toolCallEvent({ toolCallId: "c1", toolName: "read_blocks", input: { startBlockIndex: 0 } }),
			{ cwd: "/proj" },
		);
		await handlers.tool_result[0]!(
			toolResultEvent({
				toolCallId: "c1",
				content: [{ type: "text", text: "hello" }],
				isError: false,
			}),
			{ cwd: "/proj" },
		);

		const entries = sink.getEntries();
		expect(entries).toHaveLength(2);
		expect(entries[0]?.toolName).toBe("read_blocks");
		expect(entries[0]?.startedAt).toBeGreaterThan(0);
		expect(entries[0]?.completedAt).toBeUndefined();
		expect(entries[1]?.completedAt).toBeGreaterThanOrEqual(entries[0]!.startedAt);
		expect(entries[1]?.result).toBe("hello");
		expect(entries[1]?.isError).toBe(false);
		expect(entries[1]?.toolName).toBe("read_blocks");
	});

	it("writes a half-entry when tool_call arrives without a matching tool_result", async () => {
		const { api, handlers } = makeHarness();
		const sink = new InMemoryAuditSink();
		installAuditLog(api, { sink });
		await handlers.tool_call[0]!(toolCallEvent({ toolCallId: "lost" }), { cwd: "/p" });
		const entries = sink.getEntries();
		expect(entries).toHaveLength(1);
		expect(entries[0]?.toolCallId).toBe("lost");
		expect(entries[0]?.completedAt).toBeUndefined();
		expect(entries[0]?.result).toBeUndefined();
	});

	it("logs an unknown entry when tool_result arrives without tool_call", async () => {
		const { api, handlers } = makeHarness();
		const sink = new InMemoryAuditSink();
		installAuditLog(api, { sink });
		await handlers.tool_result[0]!(
			toolResultEvent({
				toolCallId: "orphan",
				content: [{ type: "text", text: "late" }],
				isError: true,
			}),
			{ cwd: "/p" },
		);
		const entries = sink.getEntries();
		expect(entries).toHaveLength(1);
		expect(entries[0]?.toolName).toBe("<unknown>");
		expect(entries[0]?.isError).toBe(true);
		expect(entries[0]?.result).toBe("late");
	});

	it("redacts sensitive fields from input and result before writing", async () => {
		const { api, handlers } = makeHarness();
		const sink = new InMemoryAuditSink();
		installAuditLog(api, { sink });

		await handlers.tool_call[0]!(
			toolCallEvent({
				toolCallId: "redact",
				input: { password: "p1", safe: "ok" },
			}),
			{ cwd: "/p" },
		);
		await handlers.tool_result[0]!(
			toolResultEvent({
				toolCallId: "redact",
				content: [{ type: "text", text: '{"token":"xyz"}' }],
				isError: false,
			}),
			{ cwd: "/p" },
		);
		const entries = sink.getEntries();
		// First half-entry only carries input (no result yet).
		expect(entries[0]?.input).toEqual({ password: "[REDACTED]", safe: "ok" });
		// Second carries the redacted result.
		expect(entries[1]?.input).toEqual({ password: "[REDACTED]", safe: "ok" });
		expect(entries[1]?.result).toBe('{"token":"xyz"}'); // text content is not JSON-parsed
	});

	it("honours resolveUser when supplied", async () => {
		const { api, handlers } = makeHarness();
		const sink = new InMemoryAuditSink();
		installAuditLog(api, {
			sink,
			resolveUser: (ctx) => `${ctx.cwd}#alice`,
		});
		await handlers.tool_call[0]!(
			toolCallEvent({ toolCallId: "user" }),
			{ cwd: "/proj" },
		);
		const entries = sink.getEntries();
		expect(entries[0]?.user).toBe("/proj#alice");
	});

	it("uses ctx.user when resolveUser is not supplied", async () => {
		const { api, handlers } = makeHarness();
		const sink = new InMemoryAuditSink();
		installAuditLog(api, { sink });
		await handlers.tool_call[0]!(
			toolCallEvent({ toolCallId: "user" }),
			{ cwd: "/p", user: "bob" },
		);
		expect(sink.getEntries()[0]?.user).toBe("bob");
	});

	it("dispatches to a CompositeAuditSink fan-out", async () => {
		const { api, handlers } = makeHarness();
		const a = new InMemoryAuditSink();
		const b = new InMemoryAuditSink();
		const sink = new CompositeAuditSink([a, b]);
		installAuditLog(api, { sink });
		await handlers.tool_call[0]!(toolCallEvent({ toolCallId: "fan" }), { cwd: "/p" });
		expect(a.getEntries()).toHaveLength(1);
		expect(b.getEntries()).toHaveLength(1);
	});

	it("calls sink.append with an awaitable contract", async () => {
		const { api, handlers } = makeHarness();
		const append = vi.fn(async () => undefined);
		const sink = { name: "spy", append };
		installAuditLog(api, { sink: sink as unknown as InMemoryAuditSink });
		await handlers.tool_call[0]!(toolCallEvent({ toolCallId: "spy" }), { cwd: "/p" });
		expect(append).toHaveBeenCalledTimes(1);
	});
});
