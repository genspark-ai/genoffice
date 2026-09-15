/**
 * skill-market extension — small registry + install/uninstall primitive that
 * powers the GenOffice skills store.
 *
 * Why this exists (parity with GenOffice legacy behaviour):
 *   Plan §5.5 calls for a "skills market" prototype with `genoffice skill
 *   install <name>` and a settings-page UI. W17 delivers the registry +
 *   install primitives so the CLI command and the settings UI both have a
 *   single source of truth. Network fetching / catalog discovery are kept
 *   out of scope; hosts plug a catalog provider at construction time.
 *
 * Mechanism:
 *   - `SkillMarketEntry` is a single row in the registry catalog.
 *   - `createSkillMarket({ catalog, skillsDir, fileSystem? })` returns a
 *     market with `list / search / install / uninstall / installedNames`.
 *   - `install(name)` writes `${skillsDir}/${name}/SKILL.md` with the entry's
 *     `body` and updates an index file at `${skillsDir}/.index.json`. pi's
 *     skill loader reads from the same directory, so a successful install
 *     makes the skill visible to the agent immediately.
 *   - `uninstall(name)` removes the directory and the index entry.
 *
 * The default `fileSystem` is a thin wrapper over `node:fs/promises`; tests
 * pass an in-memory implementation to avoid touching disk.
 */

import { mkdir, readFile, readdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";

/** A single catalog entry the host wires into the market. */
export interface SkillMarketEntry {
	/** Stable skill name; becomes the directory name under `skillsDir`. */
	name: string;
	/** Short description shown in pickers / search results. */
	description: string;
	/** Semantic version string for display only. */
	version?: string;
	/** Tags used by `search()` to filter entries. */
	tags?: readonly string[];
	/** Full skill body — frontmatter block + markdown instructions. */
	body: string;
}

/** Minimal filesystem contract used by the market. Injectable for tests. */
export interface SkillMarketFileSystem {
	mkdir(path: string, options?: { recursive?: boolean }): Promise<void>;
	readFile(path: string, encoding: "utf8"): Promise<string>;
	writeFile(path: string, contents: string, encoding: "utf8"): Promise<void>;
	readdir(path: string): Promise<readonly string[]>;
	rm(path: string, options?: { recursive?: boolean; force?: boolean }): Promise<void>;
}

const defaultFileSystem: SkillMarketFileSystem = {
	mkdir: (p, o) => mkdir(p, o).then(() => undefined),
	readFile: (p, e) => readFile(p, e) as Promise<string>,
	writeFile: (p, c, e) => writeFile(p, c, e),
	readdir: async (p) => (await readdir(p)) as readonly string[],
	rm: (p, o) => rm(p, o).then(() => undefined),
};

/** Default GenOffice skills directory: `~/.genoffice/skills`. */
export const DEFAULT_SKILLS_DIRECTORY = path.join(
	process.env.HOME ?? path.join(path.sep, "tmp"),
	".genoffice",
	"skills",
);

/** File the index is written to. Lives inside `skillsDir`. */
const INDEX_FILE = ".index.json";

export interface InstallRecord {
	name: string;
	version?: string;
	installedAt: number;
	description: string;
}

export interface SkillMarketOptions {
	/** Catalog entries the host wants to expose. */
	catalog: readonly SkillMarketEntry[];
	/** Directory skill files are written into. Defaults to `~/.genoffice/skills`. */
	skillsDir?: string;
	/** Inject a custom filesystem implementation (tests). */
	fileSystem?: SkillMarketFileSystem;
}

export interface SkillMarket {
	list(): readonly SkillMarketEntry[];
	search(query: string): readonly SkillMarketEntry[];
	install(name: string): Promise<InstallRecord>;
	uninstall(name: string): Promise<void>;
	installedNames(): Promise<readonly string[]>;
	installedRecords(): Promise<readonly InstallRecord[]>;
}

/**
 * Build a skill market from a host-supplied catalog. The catalog is held by
 * reference; mutating it after construction changes what `list()` / `search()`
 * return.
 */
export function createSkillMarket(opts: SkillMarketOptions): SkillMarket {
	const fs = opts.fileSystem ?? defaultFileSystem;
	const skillsDir = opts.skillsDir ?? DEFAULT_SKILLS_DIRECTORY;
	const byName = new Map(opts.catalog.map((entry) => [entry.name, entry]));

	async function readIndex(): Promise<InstallRecord[]> {
		try {
			const raw = await fs.readFile(path.join(skillsDir, INDEX_FILE), "utf8");
			const parsed: unknown = JSON.parse(raw);
			if (!Array.isArray(parsed)) return [];
			return parsed.filter(isInstallRecord);
		} catch (error) {
			if (isNotFound(error)) return [];
			throw error;
		}
	}

	async function writeIndex(records: readonly InstallRecord[]): Promise<void> {
		await fs.mkdir(skillsDir, { recursive: true });
		await fs.writeFile(
			path.join(skillsDir, INDEX_FILE),
			JSON.stringify(records, null, 2),
			"utf8",
		);
	}

	async function upsertRecord(
		name: string,
		existing: readonly InstallRecord[],
	): Promise<InstallRecord[]> {
		const entry = byName.get(name);
		if (!entry) {
			throw new Error(`Unknown skill "${name}". Run market.list() to see what's available.`);
		}
		const record: InstallRecord = {
			name: entry.name,
			...(entry.version !== undefined ? { version: entry.version } : {}),
			installedAt: Date.now(),
			description: entry.description,
		};
		const filtered = existing.filter((r) => r.name !== name);
		return [...filtered, record];
	}

	return {
		list() {
			return [...opts.catalog];
		},

		search(query: string) {
			const needle = query.trim().toLowerCase();
			if (needle.length === 0) return [...opts.catalog];
			return opts.catalog.filter((entry) => {
				if (entry.name.toLowerCase().includes(needle)) return true;
				if (entry.description.toLowerCase().includes(needle)) return true;
				return (entry.tags ?? []).some((tag) => tag.toLowerCase().includes(needle));
			});
		},

		async install(name) {
			const entry = byName.get(name);
			if (!entry) {
				throw new Error(`Unknown skill "${name}". Run market.list() to see what's available.`);
			}
			// idempotent install — re-installing overwrites the SKILL.md and
			// refreshes installedAt so the caller can treat it as a fresh
			// install or an in-place upgrade.
			const existing = await readIndex();
			const targetDir = path.join(skillsDir, name);
			await fs.mkdir(targetDir, { recursive: true });
			await fs.writeFile(path.join(targetDir, "SKILL.md"), entry.body, "utf8");
			const next = await upsertRecord(name, existing);
			await writeIndex(next);
			const updated = next.find((record) => record.name === name);
			if (!updated) {
				throw new Error(`Internal error: installed record for "${name}" not found after upsert.`);
			}
			return updated;
		},

		async uninstall(name) {
			const existing = await readIndex();
			if (!existing.some((record) => record.name === name)) {
				throw new Error(`Skill "${name}" is not installed in ${skillsDir}`);
			}
			const targetDir = path.join(skillsDir, name);
			await fs.rm(targetDir, { recursive: true, force: true });
			await writeIndex(existing.filter((record) => record.name !== name));
		},

		async installedNames() {
			const records = await readIndex();
			return records.map((record) => record.name);
		},

		async installedRecords() {
			return readIndex();
		},
	};
}

/* ------------------------------------------------------------------ */
/* Helpers.                                                             */
/* ------------------------------------------------------------------ */

function isInstallRecord(value: unknown): value is InstallRecord {
	if (value === null || typeof value !== "object") return false;
	const candidate = value as Record<string, unknown>;
	return (
		typeof candidate.name === "string" &&
		typeof candidate.installedAt === "number" &&
		typeof candidate.description === "string" &&
		(candidate.version === undefined || typeof candidate.version === "string")
	);
}

function isNotFound(error: unknown): boolean {
	return (
		typeof error === "object" &&
		error !== null &&
		"code" in error &&
		(error as { code?: string }).code === "ENOENT"
	);
}
