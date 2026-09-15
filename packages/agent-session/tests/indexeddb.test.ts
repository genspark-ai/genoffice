import "fake-indexeddb/auto";
import { IDBFactory } from "fake-indexeddb";
import type { JsonlSessionEntry, WebSessionBackend } from "../src/indexeddb";
import {
	createWebSessionBackend,
	DEFAULT_DATABASE_NAME,
	DEFAULT_STORE_NAME,
	fromJsonl,
	toJsonl,
} from "../src/indexeddb";
import { beforeEach, describe, expect, it } from "vitest";

function uniqueDatabaseName(): string {
	return `${DEFAULT_DATABASE_NAME}-${Math.random().toString(36).slice(2, 10)}-${Date.now()}`;
}

function buildFactory(): IDBFactory {
	// Each test gets its own IDBFactory so databases do not leak between tests.
	return new IDBFactory();
}

function userMessage(text: string): JsonlSessionEntry {
	return {
		type: "message",
		id: `m-${Math.random().toString(36).slice(2, 10)}`,
		parentId: null,
		timestamp: new Date().toISOString(),
		message: { role: "user", content: text },
	};
}

function assistantMessage(text: string): JsonlSessionEntry {
	return {
		type: "message",
		id: `a-${Math.random().toString(36).slice(2, 10)}`,
		parentId: null,
		timestamp: new Date().toISOString(),
		message: { role: "assistant", content: text },
	};
}

let factory: IDBFactory;
let backend: WebSessionBackend;

beforeEach(async () => {
	factory = buildFactory();
	backend = await createWebSessionBackend({
		databaseName: uniqueDatabaseName(),
		indexedDBFactory: factory,
	});
});

describe("createWebSessionBackend", () => {
	it("uses genoffice-sessions / sessions by default", () => {
		expect(backend.databaseName.startsWith(DEFAULT_DATABASE_NAME)).toBe(true);
		expect(backend.storeName).toBe(DEFAULT_STORE_NAME);
	});

	it("honours a custom databaseName and storeName", async () => {
		const custom = await createWebSessionBackend({
			databaseName: "custom-db",
			storeName: "custom-store",
			indexedDBFactory: factory,
		});
		expect(custom.databaseName).toBe("custom-db");
		expect(custom.storeName).toBe("custom-store");
		custom.close();
	});

	it("throws when IndexedDB is unavailable and no factory is supplied", async () => {
		const saved = (globalThis as { indexedDB?: IDBFactory }).indexedDB;
		try {
			(globalThis as { indexedDB?: IDBFactory }).indexedDB = undefined;
			await expect(
				createWebSessionBackend({ databaseName: "no-idb", indexedDBFactory: undefined }),
			).rejects.toThrow(/IndexedDB is not available/);
		} finally {
			(globalThis as { indexedDB?: IDBFactory }).indexedDB = saved;
		}
	});

	it("saves entries and reads them back unchanged", async () => {
		const entries = [userMessage("hello"), assistantMessage("hi back")];
		await backend.save("session-a", entries);
		const loaded = await backend.load("session-a");
		expect(loaded).toEqual(entries);
	});

	it("replaces prior entries when save() is called twice with the same id", async () => {
		await backend.save("session-b", [userMessage("v1")]);
		const replacement = [userMessage("v2"), assistantMessage("done")];
		await backend.save("session-b", replacement);
		const loaded = await backend.load("session-b");
		expect(loaded).toEqual(replacement);
	});

	it("preserves createdAt across updates but bumps updatedAt", async () => {
		await backend.save("session-c", [userMessage("first")]);
		const firstMeta = (await backend.list()).find((m) => m.id === "session-c");
		expect(firstMeta).toBeDefined();
		const originalCreated = firstMeta!.createdAt;

		await new Promise((resolve) => setTimeout(resolve, 5));
		await backend.save("session-c", [userMessage("second"), assistantMessage("reply")]);
		const secondMeta = (await backend.list()).find((m) => m.id === "session-c");
		expect(secondMeta).toBeDefined();
		expect(secondMeta!.createdAt).toBe(originalCreated);
		expect(secondMeta!.updatedAt).toBeGreaterThanOrEqual(originalCreated);
		expect(secondMeta!.entryCount).toBe(2);
	});

	it("load() returns undefined for unknown sessions", async () => {
		const loaded = await backend.load("does-not-exist");
		expect(loaded).toBeUndefined();
	});

	it("exists() returns true after save() and false after delete()", async () => {
		await backend.save("session-d", [userMessage("only")]);
		expect(await backend.exists("session-d")).toBe(true);
		await backend.delete("session-d");
		expect(await backend.exists("session-d")).toBe(false);
	});

	it("delete() is a no-op for unknown sessions", async () => {
		await expect(backend.delete("never-saved")).resolves.toBeUndefined();
	});

	it("list() returns lightweight metadata sorted by updatedAt desc", async () => {
		await backend.save("alpha", [userMessage("a1")]);
		await new Promise((resolve) => setTimeout(resolve, 3));
		await backend.save("beta", [userMessage("b1"), assistantMessage("b2")]);
		await new Promise((resolve) => setTimeout(resolve, 3));
		await backend.save("gamma", [userMessage("g1")]);

		const meta = await backend.list();
		const ids = meta.map((m) => m.id);
		expect(ids).toEqual(["gamma", "beta", "alpha"]);
		expect(meta.find((m) => m.id === "beta")!.entryCount).toBe(2);
	});

	it("list() reports zero entries after clear()", async () => {
		await backend.save("a", [userMessage("x")]);
		await backend.save("b", [userMessage("y")]);
		await backend.clear();
		const meta = await backend.list();
		expect(meta).toEqual([]);
	});

	it("close() releases the connection without throwing", () => {
		expect(() => backend.close()).not.toThrow();
	});

	it("isolates two backends with different databaseNames", async () => {
		const other = await createWebSessionBackend({
			databaseName: uniqueDatabaseName(),
			indexedDBFactory: factory,
		});
		try {
			const first = [userMessage("from-first")];
			const second = [userMessage("from-second")];
			await backend.save("shared-id", first);
			await other.save("shared-id", second);
			expect(await backend.load("shared-id")).toEqual(first);
			expect(await other.load("shared-id")).toEqual(second);
		} finally {
			other.close();
		}
	});

	it("survives dispose + reopen cycles against the same database", async () => {
		const databaseName = uniqueDatabaseName();
		const first = await createWebSessionBackend({ databaseName, indexedDBFactory: factory });
		const original = [userMessage("round-trip")];
		await first.save("keep", original);
		first.close();

		const second = await createWebSessionBackend({ databaseName, indexedDBFactory: factory });
		try {
			const loaded = await second.load("keep");
			expect(loaded).toEqual(original);
		} finally {
			second.close();
		}
	});
});

describe("JSONL helpers", () => {
	it("toJsonl joins entries with newlines and survives a round-trip", () => {
		const entries = [userMessage("one"), assistantMessage("two")];
		const text = toJsonl(entries);
		expect(text.split("\n")).toHaveLength(2);
		expect(fromJsonl(text)).toEqual(entries);
	});

	it("toJsonl skips entries that fail JSON.stringify without throwing", () => {
		const ok = userMessage("ok");
		const broken = { ...userMessage("broken"), unserializable: { circular: undefined as unknown } };
		// Force the circular reference to throw when stringified.
		(broken.unserializable as { circular: unknown }).circular = broken.unserializable;
		const text = toJsonl([ok, broken as unknown as JsonlSessionEntry]);
		expect(text.split("\n")).toHaveLength(1);
		expect(fromJsonl(text)).toEqual([ok]);
	});

	it("fromJsonl throws on malformed input", () => {
		expect(() => fromJsonl("not-json\n")).toThrow();
	});

	it("fromJsonl skips blank lines", () => {
		const entries = [userMessage("a")];
		const text = "\n" + toJsonl(entries) + "\n\n";
		expect(fromJsonl(text)).toEqual(entries);
	});
});
