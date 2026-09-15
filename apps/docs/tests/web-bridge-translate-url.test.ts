/**
 * Regression test for the URL migration:
 *   2026-09-15: `/crmapi/ai/translation/v1/translate*` (Dataflare Spring Boot)
 *           →  `/office-engine/api/ai/translate*` (GenOffice web-server)
 *
 * Reads `apps/docs/src/renderer/web-bridge.ts` as text and asserts the embedded
 * branches reference the new URL. This catches accidental rollbacks during
 * future bridge refactors without requiring jsdom or a running server.
 *
 * Run with: `vitest run tests/web-bridge-translate-url.test.ts`
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'

const __dirname = dirname(fileURLToPath(import.meta.url))
const webBridgePath = resolve(__dirname, '../src/renderer/web-bridge.ts')
const content = readFileSync(webBridgePath, 'utf8')

describe('web-bridge.ts translate URL migration', () => {
  it('embeds send translate-batch requests to the GenOffice HTTP endpoint (sync)', () => {
    const matches = content.match(/requestDataflare\('\/office-engine\/api\/ai\/translate'/g) ?? []
    // Two embedded sync branches: aiTranslate + aiTranslateBatch
    expect(matches.length).toBeGreaterThanOrEqual(2)
  })

  it('embeds send translate-batch-stream requests to the GenOffice SSE endpoint', () => {
    expect(content).toContain("path: '/office-engine/api/ai/translate/stream'")
  })

  it('keeps saveTranslationMemory pointing at the Dataflare multi-tenant memory endpoint', () => {
    // Memory persistence stays in Dataflare (multi-tenant DB-backed) until a
    // tenant-aware GenOffice memory endpoint is introduced.
    expect(content).toContain("'/crmapi/ai/translation/v1/memory'")
  })

  it('does not retain the legacy Dataflare translation endpoints', () => {
    expect(content).not.toContain("'/crmapi/ai/translation/v1/translate'")
    expect(content).not.toContain("'/crmapi/ai/translation/v1/translate/stream'")
  })
})
