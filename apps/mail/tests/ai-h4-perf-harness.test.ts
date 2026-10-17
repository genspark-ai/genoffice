import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { beforeEach, describe, expect, it } from 'vitest'
import { resetMailDemoSeed } from '../src/main/mail-read-handlers'
import { configureMailStoreDir, getMailStore, resetMailStoreConfig } from '../src/main/mail-store'
import { closeAllMailStores } from '../src/main/store/store'

/**
 * H4 P-6 scale harness (NOT a CI pin): per-account FTS search over a
 * 100k-message synthetic SQLite store plus the LIKE-fallback path at
 * 100k rows. Gated on RUN_PERF_HARNESS=1; H7 runs it on reference
 * hardware and records the numbers. Budgets: FTS p95 < 200 ms,
 * LIKE-fallback < 1 s at 100k rows. Synthetic mail only.
 */

const RUN = process.env.RUN_PERF_HARNESS === '1'

interface HarnessStore {
  upsertMessageFull(input: Record<string, unknown>): void
  searchLocal(query: string, opts?: { limit?: number }): { ref: { uid: string } }[]
}

function quantile(sorted: number[], q: number): number {
  if (sorted.length === 0) return 0
  return sorted[Math.min(sorted.length - 1, Math.floor(q * sorted.length))]
}

;(RUN ? describe : describe.skip)('H4 P-6 100k scale harness', () => {
  beforeEach(() => {
    resetMailDemoSeed()
    closeAllMailStores()
    resetMailStoreConfig()
  })

  it('measures FTS p95 and LIKE-fallback at 100k rows', () => {
    const dir = mkdtempSync(join(tmpdir(), 'mail-h4-harness-'))
    configureMailStoreDir(dir)
    const store = getMailStore('demo') as unknown as HarnessStore
    expect(typeof store.upsertMessageFull).toBe('function')

    const insertStart = performance.now()
    for (let i = 0; i < 100_000; i += 1) {
      store.upsertMessageFull({
        folderId: 'inbox',
        uid: `m${i}`,
        threadId: `thread-m${i}`,
        from: ['billing@example.com'],
        to: ['owner@example.com'],
        subject: `Invoice record ${i} alpha`,
        snippet: `Snippet for invoice record ${i} alpha.`,
        date: '2026-10-01T09:00:00.000Z',
        bodyText: `Routine body text for record ${i} with standard terms.`,
      })
    }
    const insertMs = performance.now() - insertStart

    // Mid-token query forces the LIKE fallback (no FTS prefix hit).
    const likeStart = performance.now()
    const likeHits = store.searchLocal('cord 999', { limit: 50 })
    const likeMs = performance.now() - likeStart
    expect(likeHits.length).toBeGreaterThan(0)

    const samples: number[] = []
    const queries = ['alpha', 'invoice record', 'record 4242', 'billing', 'routine body']
    for (let i = 0; i < 31; i += 1) {
      const query = queries[i % queries.length]
      const start = performance.now()
      const hits = store.searchLocal(query, { limit: 50 })
      samples.push(performance.now() - start)
      expect(hits.length).toBeGreaterThan(0)
    }
    samples.sort((a, b) => a - b)
    const p50 = quantile(samples, 0.5)
    const p95 = quantile(samples, 0.95)
    const max = samples[samples.length - 1]

    // Harness numbers are the deliverable of this tier-(b) run.
    console.log(
      `[h4-harness] rows=100000 insertMs=${Math.round(insertMs)} ` +
        `ftsP50=${p50.toFixed(1)}ms ftsP95=${p95.toFixed(1)}ms ftsMax=${max.toFixed(1)}ms ` +
        `likeFallbackMs=${likeMs.toFixed(1)}ms likeHits=${likeHits.length}`,
    )
    // Tier (b): harness targets are REPORTED as numbers for H7's scale
    // gate, not CI-asserted here. Budgets: FTS p95 < 200 ms,
    // LIKE-fallback < 1 s at 100k rows.
    expect(likeHits.length).toBeGreaterThan(0)
  }, 900_000)
})
