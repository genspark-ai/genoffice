/**
 * Tests for the performance module (W18 deliverable).
 *
 * Verifies:
 *   - createResponseCache basic get/set/clear + TTL expiry + maxEntries eviction
 *   - recordTiming wraps an async function, captures duration, propagates errors
 *   - summarizeBenchmark computes p50 / p95 / max / mean / errors per label
 *   - PERFORMANCE_TARGETS matches plan §8.4
 */

import { describe, expect, it } from "vitest";
import {
	createBenchmark,
	createResponseCache,
	DEFAULT_CACHE_MAX_ENTRIES,
	DEFAULT_CACHE_TTL_MS,
	PERFORMANCE_TARGETS,
	recordTiming,
	summarizeBenchmark,
} from "../src/performance";

/* ------------------------------------------------------------------ */
/* ResponseCache.                                                        */
/* ------------------------------------------------------------------ */

describe("createResponseCache", () => {
	it("returns undefined for unknown keys", () => {
		const cache = createResponseCache<string>();
		expect(cache.get("missing")).toBeUndefined();
	});

	it("stores and retrieves values", () => {
		const cache = createResponseCache<string>();
		cache.set("a", "alpha");
		expect(cache.get("a")).toBe("alpha");
		expect(cache.has("a")).toBe(true);
	});

	it("removes entries via clear()", () => {
		const cache = createResponseCache<number>();
		cache.set("a", 1);
		cache.set("b", 2);
		cache.clear();
		expect(cache.size()).toBe(0);
		expect(cache.get("a")).toBeUndefined();
	});

	it("expires entries past ttlMs", () => {
		let clock = 0;
		const cache = createResponseCache<string>({ ttlMs: 100, now: () => clock });
		cache.set("a", "alpha");
		clock = 99;
		expect(cache.get("a")).toBe("alpha");
		clock = 100;
		expect(cache.get("a")).toBeUndefined();
		expect(cache.has("a")).toBe(false);
	});

	it("uses the default ttl when none is supplied", () => {
		expect(DEFAULT_CACHE_TTL_MS).toBe(30_000);
		const cache = createResponseCache<string>();
		cache.set("a", "alpha");
		expect(cache.get("a")).toBe("alpha");
		expect(cache.has("a")).toBe(true);
	});

	it("evicts the oldest entry past maxEntries", () => {
		const cache = createResponseCache<string>({ maxEntries: 2 });
		cache.set("a", "1");
		cache.set("b", "2");
		cache.set("c", "3");
		expect(cache.size()).toBe(2);
		expect(cache.has("a")).toBe(false);
		expect(cache.has("b")).toBe(true);
		expect(cache.has("c")).toBe(true);
	});

	it("DEFAULT_CACHE_MAX_ENTRIES is 256", () => {
		expect(DEFAULT_CACHE_MAX_ENTRIES).toBe(256);
	});

	it("entries() returns a snapshot and prunes expired entries", () => {
		let clock = 0;
		const cache = createResponseCache<string>({ ttlMs: 100, now: () => clock });
		cache.set("a", "alpha");
		cache.set("b", "beta");
		clock = 50;
		const snapshot = cache.entries();
		expect(snapshot).toHaveLength(2);
		clock = 200;
		expect(cache.entries()).toHaveLength(0);
	});

	it("overwrites an entry with a new storedAt / expiresAt", () => {
		let clock = 0;
		const cache = createResponseCache<string>({ ttlMs: 100, now: () => clock });
		cache.set("a", "v1");
		clock = 90;
		cache.set("a", "v2");
		clock = 110;
		expect(cache.get("a")).toBe("v2");
	});
});

/* ------------------------------------------------------------------ */
/* Benchmark.                                                           */
/* ------------------------------------------------------------------ */

describe("recordTiming", () => {
	it("records the elapsed time and returns the function's value on success", async () => {
		const benchmark = createBenchmark();
		const value = await recordTiming(benchmark, "translate", async () => {
			await new Promise((resolve) => setTimeout(resolve, 5));
			return 42;
		});
		expect(value).toBe(42);
		const samples = benchmark.getSamples();
		expect(samples).toHaveLength(1);
		expect(samples[0]?.label).toBe("translate");
		expect(samples[0]?.durationMs).toBeGreaterThanOrEqual(0);
		expect(samples[0]?.error).toBeUndefined();
	});

	it("records an error and rethrows", async () => {
		const benchmark = createBenchmark();
		await expect(
			recordTiming(benchmark, "fail", async () => {
				throw new Error("boom");
			}),
		).rejects.toThrow("boom");
		const samples = benchmark.getSamples();
		expect(samples).toHaveLength(1);
		expect(samples[0]?.error).toBe("boom");
	});

	it("supports a custom clock for deterministic durations", async () => {
		const benchmark = createBenchmark();
		let clock = 0;
		await recordTiming(
			benchmark,
			"step",
			async () => {
				clock += 50;
			},
			() => {
				const c = clock;
				clock += 0;
				return c;
			},
		);
		const samples = benchmark.getSamples();
		expect(samples[0]?.durationMs).toBe(50);
	});
});

describe("createBenchmark + summarizeBenchmark", () => {
	it("starts empty", () => {
		const b = createBenchmark();
		expect(b.getSamples()).toEqual([]);
	});

	it("reset() clears samples", async () => {
		const b = createBenchmark();
		await recordTiming(b, "x", async () => 1);
		b.reset();
		expect(b.getSamples()).toEqual([]);
	});

	it("summarize groups by label and reports p50/p95/max/mean/errors", async () => {
		const b = createBenchmark();
		const labels = ["translate", "translate", "translate", "translate", "session-start"];
		const durations = [10, 20, 30, 40, 100];
		for (let i = 0; i < labels.length; i++) {
			await recordTiming(
				b,
				labels[i]!,
				async () => {
					await new Promise((resolve) => setTimeout(resolve, durations[i]!));
				},
			);
		}
		const summary = summarizeBenchmark(b);
		const translate = summary.find((s) => s.label === "translate")!;
		expect(translate.samples).toBe(4);
		expect(translate.errors).toBe(0);
		expect(translate.p50).toBeGreaterThanOrEqual(20);
		expect(translate.p95).toBeGreaterThanOrEqual(40);
		expect(translate.max).toBeGreaterThanOrEqual(40);

		const session = summary.find((s) => s.label === "session-start")!;
		expect(session.samples).toBe(1);
		expect(session.errors).toBe(0);
		expect(session.max).toBeGreaterThanOrEqual(100);
	});

	it("counts errors separately", async () => {
		const b = createBenchmark();
		await recordTiming(b, "x", async () => {
			throw new Error("bad");
		}).catch(() => undefined);
		await recordTiming(b, "x", async () => 1);
		const summary = summarizeBenchmark(b);
		expect(summary[0]?.errors).toBe(1);
		expect(summary[0]?.samples).toBe(2);
	});

	it("returns empty summaries for a fresh benchmark", () => {
		const b = createBenchmark();
		expect(summarizeBenchmark(b)).toEqual([]);
	});
});

/* ------------------------------------------------------------------ */
/* Plan §8.4 reference targets.                                          */
/* ------------------------------------------------------------------ */

describe("PERFORMANCE_TARGETS", () => {
	it("matches plan §8.4 (45-page translation, failure rate, 50-turn latency)", () => {
		expect(PERFORMANCE_TARGETS.translation45PagesMaxMs).toBe(90_000);
		expect(PERFORMANCE_TARGETS.toolFailureRateMax).toBe(0.02);
		expect(PERFORMANCE_TARGETS.longSessionTurnMaxMs).toBe(3_000);
	});

	it("is frozen so callers cannot mutate the reference targets", () => {
		expect(Object.isFrozen(PERFORMANCE_TARGETS)).toBe(true);
	});
});
