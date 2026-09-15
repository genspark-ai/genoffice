/**
 * IndexedDB-backed session store for web hosts.
 *
 * Stores each session as a single record keyed by `sessionId`, with the value
 * being a JSON-serializable array of entries. The on-disk shape is intentionally
 * compatible with the legacy JSONL format that `@earendil-works/pi-coding-agent`
 * uses for file-backed sessions, so future import/export between Electron and
 * Web hosts can stream entries one per line.
 */

export const DEFAULT_DATABASE_NAME = "genoffice-sessions";
export const DEFAULT_STORE_NAME = "sessions";

/**
 * Minimal structural type for a single JSONL session entry.
 * Matches the shape used by `@earendil-works/pi-coding-agent`'s
 * `SessionEntry` (session header + message / thinking_level_change / etc.),
 * but expressed without depending on the agent-core / coding-agent packages.
 */
export interface JsonlSessionEntry {
	type: string;
	id: string;
	parentId: string | null;
	timestamp: string;
	[key: string]: unknown;
}

/** Lightweight metadata projection returned by {@link WebSessionBackend.list}. */
export interface WebSessionMetadata {
	id: string;
	createdAt: number;
	updatedAt: number;
	entryCount: number;
}

export interface WebSessionBackendOptions {
	/** IndexedDB database name. Defaults to `"genoffice-sessions"`. */
	databaseName?: string;
	/** Object store name. Defaults to `"sessions"`. */
	storeName?: string;
	/** Schema version. Defaults to `1`. */
	version?: number;
	/**
	 * Override the IndexedDB factory (mainly for tests). Defaults to
	 * `globalThis.indexedDB`. Production hosts will rely on the platform-provided
	 * factory; `fake-indexeddb` is the typical test substitute.
	 */
	indexedDBFactory?: IDBFactory;
}

export interface WebSessionBackend {
	/** Persist `entries` under `sessionId`, replacing any prior value. */
	save(sessionId: string, entries: JsonlSessionEntry[]): Promise<void>;
	/** Read the entry array previously stored under `sessionId`. */
	load(sessionId: string): Promise<JsonlSessionEntry[] | undefined>;
	/** Whether `sessionId` has any entries persisted. */
	exists(sessionId: string): Promise<boolean>;
	/** Enumerate lightweight metadata for every persisted session. */
	list(): Promise<WebSessionMetadata[]>;
	/** Delete a single session. No-op when the id is unknown. */
	delete(sessionId: string): Promise<void>;
	/** Clear every persisted session. Mainly for tests. */
	clear(): Promise<void>;
	/** Close the underlying IndexedDB connection. */
	close(): void;
	/** Database name actually used. */
	databaseName: string;
	/** Object-store name actually used. */
	storeName: string;
}

interface StoredSessionRecord {
	id: string;
	createdAt: number;
	updatedAt: number;
	entries: JsonlSessionEntry[];
}

function openDatabase(
	factory: IDBFactory,
	databaseName: string,
	storeName: string,
	version: number,
): Promise<IDBDatabase> {
	return new Promise((resolve, reject) => {
		const request = factory.open(databaseName, version);
		request.onupgradeneeded = () => {
			const db = request.result;
			if (!db.objectStoreNames.contains(storeName)) {
				db.createObjectStore(storeName, { keyPath: "id" });
			}
		};
		request.onsuccess = () => resolve(request.result);
		request.onerror = () => reject(request.error ?? new Error("Failed to open IndexedDB"));
		request.onblocked = () => reject(new Error("IndexedDB open blocked by another connection"));
	});
}

function runTransaction<T>(
	db: IDBDatabase,
	storeName: string,
	mode: IDBTransactionMode,
	run: (store: IDBObjectStore) => Promise<T> | T,
): Promise<T> {
	return new Promise<T>((resolve, reject) => {
		const tx = db.transaction(storeName, mode);
		const store = tx.objectStore(storeName);
		let result: T;
		Promise.resolve(run(store))
			.then((value) => {
				result = value;
			})
			.catch(reject);
		tx.oncomplete = () => resolve(result);
		tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
		tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));
	});
}

function promisifyRequest<T>(request: IDBRequest<T>): Promise<T> {
	return new Promise((resolve, reject) => {
		request.onsuccess = () => resolve(request.result);
		request.onerror = () => reject(request.error ?? new Error("IndexedDB request failed"));
	});
}

/**
 * Create an IndexedDB-backed session store. The returned object holds one
 * IndexedDB connection; call {@link WebSessionBackend.close} to release it.
 */
export async function createWebSessionBackend(
	options: WebSessionBackendOptions = {},
): Promise<WebSessionBackend> {
	const databaseName = options.databaseName ?? DEFAULT_DATABASE_NAME;
	const storeName = options.storeName ?? DEFAULT_STORE_NAME;
	const version = options.version ?? 1;
	const factory = options.indexedDBFactory ?? globalThis.indexedDB;

	if (!factory) {
		throw new Error(
			"IndexedDB is not available in this environment; pass `indexedDBFactory` explicitly (e.g. fake-indexeddb for tests).",
		);
	}

	const db = await openDatabase(factory, databaseName, storeName, version);

	const backend: WebSessionBackend = {
		databaseName,
		storeName,

		async save(sessionId, entries) {
			const now = Date.now();
			const existing = await runTransaction(db, storeName, "readonly", (store) =>
				promisifyRequest<StoredSessionRecord | undefined>(store.get(sessionId)),
			);
			const record: StoredSessionRecord = {
				id: sessionId,
				createdAt: existing?.createdAt ?? now,
				updatedAt: now,
				entries,
			};
			await runTransaction(db, storeName, "readwrite", (store) =>
				promisifyRequest(store.put(record)),
			);
		},

		async load(sessionId) {
			const record = await runTransaction(db, storeName, "readonly", (store) =>
				promisifyRequest<StoredSessionRecord | undefined>(store.get(sessionId)),
			);
			return record?.entries;
		},

		async exists(sessionId) {
			return runTransaction(db, storeName, "readonly", (store) =>
				promisifyRequest<number>(store.count(sessionId)).then((n) => n > 0),
			);
		},

		async list() {
			const records = await runTransaction(db, storeName, "readonly", (store) =>
				promisifyRequest<StoredSessionRecord[]>(store.getAll()),
			);
			return records
				.map((record): WebSessionMetadata => ({
					id: record.id,
					createdAt: record.createdAt,
					updatedAt: record.updatedAt,
					entryCount: record.entries.length,
				}))
				.sort((left, right) => right.updatedAt - left.updatedAt);
		},

		async delete(sessionId) {
			await runTransaction(db, storeName, "readwrite", (store) =>
				promisifyRequest(store.delete(sessionId)),
			);
		},

		async clear() {
			await runTransaction(db, storeName, "readwrite", (store) =>
				promisifyRequest(store.clear()),
			);
		},

		close() {
			db.close();
		},
	};

	return backend;
}

/* ------------------------------------------------------------------ */
/* JSONL helpers — round-trip entries with the legacy file-based format. */
/* ------------------------------------------------------------------ */

/**
 * Serialize an entry array as a newline-delimited JSON string. Blank or
 * non-stringifiable lines are skipped to keep the output forward-compatible
 * with the legacy file format.
 */
export function toJsonl(entries: readonly JsonlSessionEntry[]): string {
	const lines: string[] = [];
	for (const entry of entries) {
		try {
			lines.push(JSON.stringify(entry));
		} catch {
			// skip unserializable entries rather than fail the whole export
		}
	}
	return lines.join("\n");
}

/**
 * Parse a newline-delimited JSON string back into an entry array.
 * Throws on the first malformed line to surface corrupt files loudly.
 */
export function fromJsonl(text: string): JsonlSessionEntry[] {
	const entries: JsonlSessionEntry[] = [];
	for (const rawLine of text.split("\n")) {
		const line = rawLine.trim();
		if (line.length === 0) continue;
		entries.push(JSON.parse(line) as JsonlSessionEntry);
	}
	return entries;
}
