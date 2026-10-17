/**
 * The decision-model routes.
 *
 * Several vendors publish more than one size or generation behind the same
 * URL, differing only in the `model` the body carries. These tests pin each
 * route's URL and model id, and pin the one thing that silently rots: the
 * settings dropdown and the code's endpoint table are two lists of the same
 * set, and a route added to only one of them is a selectable option that
 * throws, or a working model nobody can reach.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { bodyModel, endpointUrl, type DecisionCallOptions } from '../src/main/file-index/decider'
import type { DecisionEndpoint } from '../src/shared/home-api'

const here = dirname(fileURLToPath(import.meta.url))

const opts = (over: Partial<DecisionCallOptions> = {}): DecisionCallOptions => ({
  endpoint: 'openrouter',
  key: '',
  customBaseUrl: '',
  customModel: '',
  cloudflareAccountId: '',
  cloudflareModel: '',
  ...over,
})

describe('decision model routes', () => {
  it('sends each vendor its own published model id', () => {
    // The ids are the vendors', not ours: a wrong one is a 400 from Perplexity
    // and "unknown model" from Liquid, which is why they are spelled out here
    // rather than derived.
    expect(bodyModel(opts({ endpoint: 'perplexity' }))).toBe('pplx-decider-v1-27b')
    expect(bodyModel(opts({ endpoint: 'perplexity-v12' }))).toBe('pplx-decider-v1.2')
    expect(bodyModel(opts({ endpoint: 'perplexity-flash' }))).toBe('pplx-decider-flash-v1.2')
    expect(bodyModel(opts({ endpoint: 'liquid' }))).toBe('d1')
    expect(bodyModel(opts({ endpoint: 'liquid-ultra' }))).toBe('d1-ultra')
  })

  it('shares one URL between a vendor s sizes', () => {
    // Same endpoint, same protocol; only the model differs. If these URLs ever
    // diverge it is a vendor changing its API, which is worth catching here.
    const urls = (['perplexity', 'perplexity-v12', 'perplexity-flash'] as const).map((endpoint) =>
      endpointUrl(opts({ endpoint })),
    )
    expect(new Set(urls).size).toBe(1)
    expect(urls[0]).toBe('https://api.perplexity.ai/v1/decisions')

    const liquid = (['liquid', 'liquid-ultra'] as const).map((endpoint) =>
      endpointUrl(opts({ endpoint })),
    )
    expect(new Set(liquid).size).toBe(1)
    expect(liquid[0]).toBe('https://api.liquid.ai/decisions/v1/systemone')
  })

  it('sends Workers AI the bare id in the body and the @cf path in the URL', () => {
    const cloudflare = opts({ endpoint: 'cloudflare-flash', cloudflareAccountId: 'acct 1' })
    expect(endpointUrl(cloudflare)).toBe(
      'https://api.cloudflare.com/client/v4/accounts/acct%201/ai/run/@cf/cloudflare/clef-flash',
    )
    expect(bodyModel(cloudflare)).toBe('clef-flash')
  })

  it('honours a Workers AI model the reader typed over either default', () => {
    const custom = opts({
      endpoint: 'cloudflare',
      cloudflareAccountId: 'acct',
      cloudflareModel: '@cf/cloudflare/clef-omni',
    })
    expect(endpointUrl(custom)).toContain('/ai/run/@cf/cloudflare/clef-omni')
    expect(bodyModel(custom)).toBe('clef-omni')
  })

  it('refuses a Workers AI model outside the Clef family', () => {
    const bad = opts({
      endpoint: 'cloudflare',
      cloudflareAccountId: 'a',
      cloudflareModel: 'evil-model',
    })
    expect(() => bodyModel(bad)).toThrow()
  })
})

describe('the settings dropdown matches the endpoint table', () => {
  const modalSource = readFileSync(join(here, '../src/renderer/src/SettingsModal.tsx'), 'utf8')
  const rerankSource = readFileSync(join(here, '../src/main/file-index/rerank.ts'), 'utf8')

  it('offers exactly every endpoint the code knows, and nothing else', () => {
    // Only the decision-model list, not the other dozen pickers in this file.
    // The type annotation carries its own "]", so the block ends at a "]"
    // that starts a line of its own.
    const start = modalSource.indexOf('const DECISION_ENDPOINTS')
    const block = modalSource.slice(start, modalSource.indexOf('\n]', start))
    const offered = [...block.matchAll(/\{ value: '([a-z0-9-]+)', label:/g)].map(
      (m) => m[1] as DecisionEndpoint,
    )
    const idsStart = rerankSource.indexOf('const ENDPOINT_IDS')
    const known = rerankSource
      .slice(idsStart, rerankSource.indexOf('\n]', idsStart))
      .matchAll(/'([a-z0-9-]+)'/g)
      .map((m) => m[1] as DecisionEndpoint)

    expect([...known].sort()).toEqual([...offered].sort())
    // the four routes this PR added, so the comparison above cannot quietly
    // pass because both lists shrank together
    for (const added of [
      'perplexity-v12',
      'perplexity-flash',
      'liquid-ultra',
      'cloudflare-flash',
    ] as const) {
      expect(offered).toContain(added)
    }
  })

  it('gives every endpoint a key slot in the defaults', () => {
    const keysBlock = rerankSource.slice(
      rerankSource.indexOf('keys: {'),
      rerankSource.indexOf('customBaseUrl'),
    )
    for (const added of [
      'perplexity-v12',
      'perplexity-flash',
      'liquid-ultra',
      'cloudflare-flash',
    ] as const) {
      expect(keysBlock).toContain(`'${added}'`)
    }
  })
})
