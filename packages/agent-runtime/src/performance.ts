/**
 * performance — small utilities that support plan §5.6 (parallel tools,
 * provider response cache) and plan §8.4 (latency benchmarks).
 *
 * Why this exists (parity with GenOffice legacy behaviour):
 *   Plan §8.4 lists concrete latency / failure targets. pi's
 *   `ToolExecutionMode = "parallel"` is on by default (see
 *   `@earendil-works/pi-agent-core` types.ts), so W18 does not need to
 *   configure it. What W18 adds is:
 *     - a TTL-aware response cache so hosts can deduplicate identical
 *       provider calls within a short window,
 *     - a lightweight `Benchmark` harness for measuring tool-call and
 *       session-start latencies against the §8.4 targets.
 *
 * Mechanism:
 *   - `createResponseCache({ ttlMs, maxEntries })` returns a cache that
 *     stores `unknown` JSON-serialisable values keyed by an arbitrary
 *     string. Reads past the TTL or after eviction return a miss.
 *   - `recordTiming(label, fn)` wraps an async function and records the
 *     elapsed milliseconds into an externally-supplied `Benchmark`.
 *   - `summarizeBenchmark(b)` returns `{ label, samples, p50, p95, max,
 *     mean }` for the §8.4 dashboard.
 *
 * Note: pi's parallel mode and provider streaming are not re-implemented
 * here; W18 focuses on the **measurement + cache** side of performance.
 */

export const DEFAULT_CACHE_TTL_MS = 30_000;
export const DEFAULT_CACHE_MAX_ENTRIES = 256;

export interface ResponseCacheOptions {
	/** Time-to-live per entry in milliseconds. Defaults to `30_000`. */
	ttlMs?: number;
	/** Soft cap on entries; oldest is evicted past the cap. Defaults to `256`. */
	maxEntries?: number;
	/** Override `Date.now` (tests). */
	now?: () => number;
}

export interface CacheEntry<V> {
	value: V;
	storedAt: number;
	expiresAt: number;
}

export interface ResponseCache<V = unknown> {
	get(key: string): V | undefined;
	set(key: string, value: V): void;
	has(key: string): boolean;
	clear(): void;
	size(): number;
	/** Test / introspection: snapshot of every live entry. */
	entries(): ReadonlyArray<{ key: string; entry: CacheEntry<V> }>;
}

/**
 * Build a TTL-aware in-memory response cache.
 *
 * The cache uses an insertion-ordered map and evicts the oldest entry when
 * `maxEntries` is exceeded. Entries past `ttlMs` are considered misses on
 * `get()` and are also pruned lazily on the next read.
 */
export function createResponseCache<V = unknown>(opts: ResponseCacheOptions = {}): ResponseCache<V> {
	const ttlMs = opts.ttlMs ?? DEFAULT_CACHE_TTL_MS;
	const maxEntries = opts.maxEntries ?? DEFAULT_CACHE_MAX_ENTRIES;
	const now = opts.now ?? Date.now;
	const map = new Map<string, CacheEntry<V>>();

	function isExpired(entry: CacheEntry<V>): boolean {
		return entry.expiresAt <= now();
	}

	return {
		get(key) {
			const entry = map.get(key);
			if (!entry) return undefined;
			if (isExpired(entry)) {
				map.delete(key);
				return undefined;
			}
			return entry.value;
		},
		set(key, value) {
			const storedAt = now();
			map.set(key, {
				value,
				storedAt,
				expiresAt: storedAt + ttlMs,
			});
			// Evict oldest entries past the cap.
			while (map.size > maxEntries) {
				const oldestKey = map.keys().next().value;
				if (oldestKey === undefined) break;
				map.delete(oldestKey);
			}
		},
		has(key) {
			const entry = map.get(key);
			if (!entry) return false;
			if (isExpired(entry)) {
				map.delete(key);
				return false;
			}
			return true;
		},
		clear() {
			map.clear();
		},
		size() {
			return map.size;
		},
		entries() {
			const cutoff = now();
			const out: Array<{ key: string; entry: CacheEntry<V> }> = [];
			for (const [key, entry] of map) {
				if (entry.expiresAt <= cutoff) {
					map.delete(key);
					continue;
				}
				out.push({ key, entry });
			}
			return out;
		},
	};
}

/* ------------------------------------------------------------------ */
/* Benchmark harness.                                                   */
/* ------------------------------------------------------------------ */

export interface TimingSample {
	label: string;
	durationMs: number;
	timestamp: number;
	error?: string;
}

export interface Benchmark {
	record(sample: TimingSample): void;
	getSamples(): readonly TimingSample[];
	reset(): void;
}

export function createBenchmark(): Benchmark {
	const samples: TimingSample[] = [];
	return {
		record(sample) {
			samples.push(sample);
		},
		getSamples() {
			return samples.slice();
		},
		reset() {
			samples.length = 0;
		},
	};
}

/**
 * Run `fn`, record the elapsed milliseconds against `label` in `benchmark`,
 * and return the original return value. Re-throws after recording.
 */
export async function recordTiming<T>(
	benchmark: Benchmark,
	label: string,
	fn: () => Promise<T>,
	now: () => number = Date.now,
): Promise<T> {
	const startedAt = now();
	try {
		const value = await fn();
		benchmark.record({ label, durationMs: now() - startedAt, timestamp: startedAt });
		return value;
	} catch (error) {
		benchmark.record({
			label,
			durationMs: now() - startedAt,
			timestamp: startedAt,
			error: error instanceof Error ? error.message : String(error),
		});
		throw error;
	}
}

export interface BenchmarkSummary {
	label: string;
	samples: number;
	p50: number;
	p95: number;
	max: number;
	mean: number;
	errors: number;
}

/** Aggregate samples by label into p50 / p95 / max / mean stats. */
export function summarizeBenchmark(benchmark: Benchmark): readonly BenchmarkSummary[] {
	const groups = new Map<string, TimingSample[]>();
	for (const sample of benchmark.getSamples()) {
		const group = groups.get(sample.label) ?? [];
		group.push(sample);
		groups.set(sample.label, group);
	}

	const summaries: BenchmarkSummary[] = [];
	for (const [label, list] of groups) {
		const durations = list.map((s) => s.durationMs).sort((a, b) => a - b);
		const errors = list.filter((s) => s.error !== undefined).length;
		summaries.push({
			label,
			samples: list.length,
			p50: percentile(durations, 0.5),
			p95: percentile(durations, 0.95),
			max: durations[durations.length - 1] ?? 0,
			mean: durations.reduce((sum, d) => sum + d, 0) / Math.max(1, durations.length),
			errors,
		});
	}
	return summaries;
}

function percentile(sortedAsc: number[], p: number): number {
	if (sortedAsc.length === 0) return 0;
	const index = Math.min(sortedAsc.length - 1, Math.floor(p * sortedAsc.length));
	return sortedAsc[index]!;
}

/* ------------------------------------------------------------------ */
/* Plan §8.4 reference targets.                                          */
/* ------------------------------------------------------------------ */

export const PERFORMANCE_TARGETS = Object.freeze({
	/** 45-page docx translation end-to-end. */
	translation45PagesMaxMs: 90_000,
	/** Failure rate across tool calls. */
	toolFailureRateMax: 0.02,
	/** 50-turn long-session response latency. */
	longSessionTurnMaxMs: 3_000,
});
