import { mkdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import type { Context } from "@earendil-works/pi-agent-core";
import { BACKGROUND_CONTEXT } from "@earendil-works/pi-agent-core";
import {
	createNodeSqliteFactory,
	SqliteSessionRepo,
} from "@earendil-works/pi-session-backend-sqlite-node";

/** Default GenOffice user-data directory under `~/.genoffice`. */
export const DEFAULT_USER_HOME_DIR = path.join(os.homedir(), ".genoffice");

/** Default SQLite database filename inside the user-data directory. */
export const DEFAULT_DATABASE_FILENAME = "sessions.sqlite";

export interface ElectronSessionBackendOptions {
	/** Working directory the agent session reports to its host app. */
	cwd: string;
	/** User-data directory. Defaults to `~/.genoffice`. */
	userHomeDir?: string;
	/**
	 * Absolute path to the SQLite database file, or a relative path that will be
	 * resolved against `userHomeDir`. Defaults to `~/.genoffice/sessions.sqlite`.
	 *
	 * When supplied, every session is stored in this single SQLite database;
	 * when omitted, each session lives in its own `${id}.sqlite` file inside
	 * the user-data directory.
	 */
	databasePath?: string;
	/** Clock override used by the SQLite repository. Defaults to `Date.now`. */
	now?: () => number;
}

export interface ElectronSessionBackend {
	/** Underlying SQLite session repository (one per backend instance). */
	repository: SqliteSessionRepo;
	/** Absolute path to the SQLite database file actually used. */
	databasePath: string;
	/** Absolute path to the GenOffice user-data directory actually used. */
	userHomeDir: string;
	/**
	 * Release the underlying database connection. Idempotent: safe to call
	 * multiple times. Uses {@link BACKGROUND_CONTEXT} from `@earendil-works/pi-agent-core`
	 * (chord context) so the host does not have to assemble one for cleanup.
	 */
	dispose: () => Promise<void>;
}

/**
 * Resolve a database path option to an absolute filesystem path.
 * Relative inputs are joined onto `userHomeDir`; absolute inputs are passed through.
 */
export function resolveDatabasePath(databasePath: string | undefined, userHomeDir: string): string {
	if (!databasePath) {
		return path.join(userHomeDir, DEFAULT_DATABASE_FILENAME);
	}
	return path.isAbsolute(databasePath) ? databasePath : path.join(userHomeDir, databasePath);
}

/**
 * Create a SQLite-backed session store for an Electron host process.
 *
 * GenOffice stores all user data under `~/.genoffice` by default. The factory:
 *   1. ensures the user-data directory exists,
 *   2. resolves the SQLite database path (absolute or relative),
 *   3. opens the database lazily via the Node `node:sqlite` factory,
 *   4. returns a {@link SqliteSessionRepo} ready for `create / open / list / fork`.
 *
 * Disposal is delegated through {@link SqliteSessionRepo.close}, which is
 * idempotent and drains all currently open sessions before releasing the
 * underlying database connection.
 */
export async function createElectronSessionBackend(
	options: ElectronSessionBackendOptions,
): Promise<ElectronSessionBackend> {
	const userHomeDir = options.userHomeDir ?? DEFAULT_USER_HOME_DIR;
	await mkdir(userHomeDir, { recursive: true });

	const databasePath = resolveDatabasePath(options.databasePath, userHomeDir);

	const repository = new SqliteSessionRepo({
		directory: userHomeDir,
		databasePath,
		databaseFactory: createNodeSqliteFactory(),
		...(options.now ? { now: options.now } : {}),
	});

	return {
		repository,
		databasePath,
		userHomeDir,
		dispose: (): Promise<void> => repository.close(BACKGROUND_CONTEXT as Context),
	};
}
