import { existsSync, mkdirSync, mkdtempSync, readdirSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import type { Context, SessionMetadata } from "@earendil-works/pi-agent-core";
import { BACKGROUND_CONTEXT } from "@earendil-works/pi-agent-core";
import { beforeEach, describe, expect, it } from "vitest";
import {
	createElectronSessionBackend,
	DEFAULT_DATABASE_FILENAME,
	DEFAULT_USER_HOME_DIR,
	resolveDatabasePath,
} from "../src/sqlite";

/**
 * `SqliteSessionMetadata` is the concrete SessionMetadata subtype returned by the
 * SQLite backend. It is not re-exported from `@earendil-works/pi-session-backend-sqlite-node`,
 * so we describe the relevant fields inline.
 */
type SqliteMetadata = SessionMetadata & { path: string };

function getSqliteMetadata(metadata: SessionMetadata): SqliteMetadata {
	const candidate = metadata as Partial<SqliteMetadata>;
	if (typeof candidate.path !== "string") {
		throw new Error(`Expected SQLite metadata to expose a string path, got ${String(candidate.path)}`);
	}
	return { ...metadata, path: candidate.path };
}

let root: string;

beforeEach(() => {
	root = mkdtempSync(path.join(os.tmpdir(), "agent-session-sqlite-"));
});

describe("createElectronSessionBackend", () => {
	it("uses ~/.genoffice/sessions.sqlite by default", () => {
		const resolved = resolveDatabasePath(undefined, DEFAULT_USER_HOME_DIR);
		expect(resolved).toBe(path.join(os.homedir(), ".genoffice", DEFAULT_DATABASE_FILENAME));
	});

	it("joins relative database paths onto the user-data directory", () => {
		const resolved = resolveDatabasePath("nested/store.sqlite", "/tmp/userdata");
		expect(resolved).toBe(path.join("/tmp/userdata", "nested", "store.sqlite"));
	});

	it("passes absolute database paths through unchanged", () => {
		const absolute = path.join(root, "absolute-store.sqlite");
		const resolved = resolveDatabasePath(absolute, "/tmp/userdata");
		expect(resolved).toBe(absolute);
	});

	it("creates the user-data directory if it does not exist", async () => {
		const userHomeDir = path.join(root, "user-home");
		const backend = await createElectronSessionBackend({ cwd: root, userHomeDir });
		try {
			expect(existsSync(userHomeDir)).toBe(true);
			expect(backend.userHomeDir).toBe(userHomeDir);
			expect(backend.databasePath).toBe(path.join(userHomeDir, DEFAULT_DATABASE_FILENAME));
		} finally {
			await backend.dispose();
		}
	});

	it("respects a custom cwd, userHomeDir, and absolute databasePath", async () => {
		const userHomeDir = path.join(root, "user-home");
		const cwd = path.join(root, "docs");
		mkdirSync(cwd, { recursive: true });
		const databasePath = path.join(root, "another", "sessions.sqlite");
		mkdirSync(path.dirname(databasePath), { recursive: true });

		const backend = await createElectronSessionBackend({ cwd, userHomeDir, databasePath });
		try {
			expect(backend.userHomeDir).toBe(userHomeDir);
			expect(backend.databasePath).toBe(databasePath);
			// The database file is created lazily on the first session; trigger that.
			const session = await backend.repository.create({ id: "toucher" }, BACKGROUND_CONTEXT);
			const sqlitePath = (session.metadata as { path: string }).path;
			expect(existsSync(sqlitePath)).toBe(true);
			await session.close(BACKGROUND_CONTEXT);
		} finally {
			await backend.dispose();
		}
	});

	it("opens a repository that can create a session and append a user message", async () => {
		const backend = await createElectronSessionBackend({ cwd: root, userHomeDir: root });
		const ctx: Context = BACKGROUND_CONTEXT;
		try {
			const session = await backend.repository.create({ id: "smoke-1" }, ctx);
			const branch = await session.createBranch("main", null, ctx);
			const entryId = await branch.appendMessage(
				{ role: "user", content: "hello from agent-session", timestamp: Date.now() },
				ctx,
			);
			expect(typeof entryId).toBe("string");
			expect(entryId.length).toBeGreaterThan(0);

			expect(session.metadata.id).toBe("smoke-1");
			expect(typeof session.metadata.createdAt).toBe("number");

			await session.close(ctx);
		} finally {
			await backend.dispose();
		}
	});

	it("persists data across dispose / re-open cycles", async () => {
		const databasePath = path.join(root, "persisted.sqlite");
		const userHomeDir = root;

		const first = await createElectronSessionBackend({ cwd: root, userHomeDir, databasePath });
		const ctxA: Context = BACKGROUND_CONTEXT;
		const session = await first.repository.create({ id: "roundtrip" }, ctxA);
		const branch = await session.createBranch("main", null, ctxA);
		const firstEntryId = await branch.appendMessage(
			{ role: "user", content: "first message", timestamp: Date.now() },
			ctxA,
		);
		await session.close(ctxA);
		await first.dispose();

		const second = await createElectronSessionBackend({ cwd: root, userHomeDir, databasePath });
		try {
			const ctxB: Context = BACKGROUND_CONTEXT;
			const all = await second.repository.list(undefined, ctxB);
			expect(all.map((m) => m.id)).toContain("roundtrip");

			const target = getSqliteMetadata(all.find((m) => m.id === "roundtrip")!);
			const resumed = await second.repository.open(target, ctxB);
			const entries = await resumed.findEntries(undefined, ctxB);
			const messageEntries = entries.filter((e) => e.type === "message");
			expect(messageEntries.length).toBeGreaterThan(0);
			const firstEntry = messageEntries.find((e) => e.id === firstEntryId);
			expect(firstEntry).toBeDefined();
			await resumed.close(ctxB);
		} finally {
			await second.dispose();
		}
	});

	it("exposes a dispose() that is safe to call multiple times", async () => {
		const backend = await createElectronSessionBackend({ cwd: root, userHomeDir: root });
		await backend.dispose();
		await expect(backend.dispose()).resolves.toBeUndefined();
	});

	it("produces a backend whose repository supports list()", async () => {
		const userHomeDir = path.join(root, "user-home");
		const backend = await createElectronSessionBackend({ cwd: root, userHomeDir });
		const ctx: Context = BACKGROUND_CONTEXT;
		try {
			await backend.repository.create({ id: "alpha" }, ctx);
			await backend.repository.create({ id: "beta" }, ctx);
			const listed = await backend.repository.list(undefined, ctx);
			const ids = listed.map((m) => m.id).sort();
			expect(ids).toEqual(["alpha", "beta"]);
		} finally {
			await backend.dispose();
		}
	});

	it("leaves the user-data directory in place after dispose (host-managed cleanup)", async () => {
		const userHomeDir = path.join(root, "user-home");
		const backend = await createElectronSessionBackend({ cwd: root, userHomeDir });
		await backend.dispose();
		const remaining = readdirSync(userHomeDir);
		// The repository does not delete the directory; the host is responsible.
		// This test pins current behaviour so future refactors surface a deliberate change.
		expect(Array.isArray(remaining)).toBe(true);
	});
});

describe("DEFAULT_* constants", () => {
	it("DEFAULT_USER_HOME_DIR points to ~/.genoffice", () => {
		expect(DEFAULT_USER_HOME_DIR).toBe(path.join(os.homedir(), ".genoffice"));
	});

	it("DEFAULT_DATABASE_FILENAME is sessions.sqlite", () => {
		expect(DEFAULT_DATABASE_FILENAME).toBe("sessions.sqlite");
	});
});

// Best-effort cleanup if the suite bails out before afterEach had a chance to run.
process.on("exit", () => {
	try {
		rmSync(root, { recursive: true, force: true });
	} catch {
		// ignore cleanup errors
	}
});
