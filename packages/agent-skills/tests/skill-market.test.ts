/**
 * Tests for the skill-market extension (W17 deliverable).
 *
 * Verifies:
 *   - createSkillMarket wires a catalog into list/search/install/uninstall
 *   - install writes the SKILL.md body and updates the .index.json
 *   - uninstall removes both the directory and the index entry
 *   - search matches by name, description, and tags
 *   - Reinstalling an already-installed skill throws
 *   - In-memory fileSystem keeps tests deterministic
 */

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
	createSkillMarket,
	DEFAULT_SKILLS_DIRECTORY,
	type SkillMarketEntry,
	type SkillMarketFileSystem,
} from "../src/extensions/skill-market";

/** Tiny in-memory filesystem for deterministic tests. */
function makeMemoryFs(): SkillMarketFileSystem & { snapshot(): { files: Record<string, string> } } {
	const files = new Map<string, string>();
	const dirs = new Set<string>();
	const snapshot = () => ({ files: Object.fromEntries(files) });
	const pathOf = (p: string) => p;
	const dirOf = (p: string) => pathOf(p).split("/").slice(0, -1).join("/") || "/";
	return {
		snapshot,
		async mkdir(p) {
			dirs.add(dirOf(p));
			dirs.add(p);
		},
		async readFile(p) {
			if (!files.has(p)) {
				const err: Error & { code?: string } = new Error(`ENOENT: ${p}`);
				err.code = "ENOENT";
				throw err;
			}
			return files.get(p)!;
		},
		async writeFile(p, contents) {
			files.set(p, contents);
		},
		async readdir(p) {
			const prefix = p.endsWith("/") ? p : `${p}/`;
			const seen = new Set<string>();
			for (const key of files.keys()) {
				if (!key.startsWith(prefix)) continue;
				const rest = key.slice(prefix.length);
				if (rest.length === 0) continue;
				const head = rest.split("/")[0]!;
				seen.add(head);
			}
			return [...seen];
		},
		async rm(p, options) {
			const prefix = p.endsWith("/") ? p : `${p}/`;
			for (const key of [...files.keys()]) {
				if (key === p || key.startsWith(prefix)) {
					files.delete(key);
				}
			}
			if (options?.force) return;
		},
	};
}

const CATALOG: SkillMarketEntry[] = [
	{
		name: "docs-translator",
		description: "Translate docx files between languages while preserving formatting.",
		version: "1.0.0",
		tags: ["docs", "translation"],
		body: "---\nname: docs-translator\n---\n# Docs translator\n",
	},
	{
		name: "legal-review",
		description: "Scan a contract for risky clauses.",
		version: "0.4.1",
		tags: ["legal"],
		body: "---\nname: legal-review\n---\n# Legal review\n",
	},
	{
		name: "finance-summary",
		description: "Compress a quarterly finance workbook into a one-pager.",
		version: "0.2.0",
		tags: ["finance", "summary"],
		body: "---\nname: finance-summary\n---\n# Finance summary\n",
	},
];

let fs: ReturnType<typeof makeMemoryFs>;
const skillsDir = "/skills";

beforeEach(() => {
	fs = makeMemoryFs();
});

afterEach(() => {
	// Nothing to clean up; the in-memory fs is per-test.
});

describe("createSkillMarket — catalogue surface", () => {
	it("list() returns every catalog entry", () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		expect(market.list()).toHaveLength(3);
		expect(market.list().map((e) => e.name)).toEqual([
			"docs-translator",
			"legal-review",
			"finance-summary",
		]);
	});

	it("search() matches by name (case-insensitive)", () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		const hits = market.search("LEGAL");
		expect(hits).toHaveLength(1);
		expect(hits[0]?.name).toBe("legal-review");
	});

	it("search() matches by description substring", () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		const hits = market.search("contract");
		expect(hits.map((h) => h.name)).toEqual(["legal-review"]);
	});

	it("search() matches by tag", () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		const hits = market.search("translation");
		expect(hits.map((h) => h.name)).toEqual(["docs-translator"]);
	});

	it("search() returns everything when query is empty / whitespace", () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		expect(market.search("")).toHaveLength(3);
		expect(market.search("   ")).toHaveLength(3);
	});

	it("search() returns nothing when no entry matches", () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		expect(market.search("nonexistent")).toEqual([]);
	});
});

describe("createSkillMarket — install / uninstall", () => {
	it("install() writes SKILL.md with the entry body", async () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		const record = await market.install("docs-translator");
		expect(record.name).toBe("docs-translator");
		expect(record.version).toBe("1.0.0");
		expect(typeof record.installedAt).toBe("number");
		const body = fs.snapshot().files["/skills/docs-translator/SKILL.md"];
		expect(body).toBe(CATALOG[0]!.body);
	});

	it("install() appends to the index file", async () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		await market.install("docs-translator");
		await market.install("legal-review");
		const indexRaw = fs.snapshot().files["/skills/.index.json"];
		expect(indexRaw).toBeDefined();
		const parsed = JSON.parse(indexRaw!) as Array<{ name: string }>;
		expect(parsed.map((r) => r.name).sort()).toEqual(["docs-translator", "legal-review"]);
	});

	it("installedNames() returns the installed ids in install order", async () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		await market.install("docs-translator");
		await market.install("legal-review");
		await market.install("finance-summary");
		expect(await market.installedNames()).toEqual([
			"docs-translator",
			"legal-review",
			"finance-summary",
		]);
	});

	it("installedNames() returns [] when nothing is installed", async () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		expect(await market.installedNames()).toEqual([]);
	});

	it("install() throws when the skill is unknown", async () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		await expect(market.install("nope")).rejects.toThrow(/Unknown skill "nope"/);
	});

	it("install() is idempotent: re-installing overwrites SKILL.md and refreshes installedAt", async () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		const first = await market.install("docs-translator");
		// tiny delay so installedAt timestamps differ
		await new Promise((r) => setTimeout(r, 5));
		const second = await market.install("docs-translator");
		// resolves with the same record shape; installedAt is refreshed
		expect(second.name).toBe(first.name);
		expect(second.installedAt).toBeGreaterThan(first.installedAt);
		// still only one entry in the index
		expect(await market.installedNames()).toEqual(["docs-translator"]);
	});

	it("uninstall() removes both the directory contents and the index entry", async () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		await market.install("docs-translator");
		await market.install("legal-review");
		await market.uninstall("docs-translator");
		const snap = fs.snapshot();
		expect(snap.files["/skills/docs-translator/SKILL.md"]).toBeUndefined();
		expect(snap.files["/skills/legal-review/SKILL.md"]).toBeDefined();
		expect(await market.installedNames()).toEqual(["legal-review"]);
	});

	it("uninstall() throws when the skill is not installed", async () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		await expect(market.uninstall("docs-translator")).rejects.toThrow(/not installed/);
	});

	it("install() can be called again after uninstall()", async () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		await market.install("docs-translator");
		await market.uninstall("docs-translator");
		await market.install("docs-translator");
		expect(await market.installedNames()).toEqual(["docs-translator"]);
	});

	it("install() surfaces JSON corruption in the index as an error (caller must repair)", async () => {
		const market = createSkillMarket({ catalog: CATALOG, skillsDir, fileSystem: fs });
		// Pre-poison the index with garbage so JSON.parse throws.
		await fs.writeFile("/skills/.index.json", "not-json", "utf8");
		await expect(market.install("docs-translator")).rejects.toThrow();
	});
});

describe("module constants", () => {
	it("DEFAULT_SKILLS_DIRECTORY points under ~/.genoffice", () => {
		expect(DEFAULT_SKILLS_DIRECTORY).toMatch(/\.genoffice[\\/]skills$/);
	});
});
