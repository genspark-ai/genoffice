import { describe, expect, it } from 'vitest'
import { listCustomModels } from '../src/custom-models'
import { normalizeCustomProfile } from '../src/custom-profile'
import type { AiCustomModelProfile } from '../src/types'

function stub(body: unknown) {
  globalThis.fetch = (async () =>
    new Response(JSON.stringify(body), {
      status: 200,
      headers: { 'content-type': 'application/json' },
    })) as typeof fetch
}

describe('listCustomModels: capability facts from the listing', () => {
  it('reads the rich OpenAI-compatible metadata a gateway publishes', async () => {
    stub({
      data: [
        {
          id: 'cn:deepseek-v4.1-flash',
          name: 'Deepseek-V4.1-Flash',
          context_length: 1_000_000,
          max_output_tokens: 128_000,
          supports_images: true,
        },
        { id: 'cn:auto', supports_images: false, max_output_tokens: 32_000 },
      ],
    })
    const out = await listCustomModels('https://gw.example.com/v1', 'k')
    expect(out.models).toEqual(['cn:deepseek-v4.1-flash', 'cn:auto'])
    expect(out.profiles).toEqual([
      {
        id: 'cn:deepseek-v4.1-flash',
        name: 'Deepseek-V4.1-Flash',
        vision: true,
        maxTokens: 128_000,
        contextWindow: 1_000_000,
      },
      { id: 'cn:auto', vision: false, maxTokens: 32_000 },
    ])
  })

  it('reads the Gemini limits and strips the resource prefix', async () => {
    stub({
      models: [
        { name: 'models/gemini-3.8-flash', inputTokenLimit: 1_048_576, outputTokenLimit: 65_536 },
      ],
    })
    const out = await listCustomModels('https://gw.example.com/v1beta', 'k', 'gemini')
    expect(out.models).toEqual(['gemini-3.8-flash'])
    expect(out.profiles).toEqual([
      { id: 'gemini-3.8-flash', maxTokens: 65_536, contextWindow: 1_048_576 },
    ])
  })

  it('keeps the Anthropic display name and declares nothing else', async () => {
    stub({ data: [{ id: 'claude-sonnet-5', display_name: 'Claude Sonnet 5' }] })
    const out = await listCustomModels('https://gw.example.com', 'k', 'anthropic')
    expect(out.profiles).toEqual([{ id: 'claude-sonnet-5', name: 'Claude Sonnet 5' }])
  })

  it('an entry that carries only an id is not a declaration', async () => {
    stub({ data: [{ id: 'bare-model' }, 'plain-string'] })
    const out = await listCustomModels('https://gw.example.com/v1')
    expect(out.models).toEqual(['bare-model', 'plain-string'])
    expect(out.profiles).toBeUndefined()
  })

  it('a listing with no metadata at all yields no profiles field', async () => {
    stub({ data: [{ id: 'a' }] })
    const out = await listCustomModels('https://gw.example.com/v1')
    expect(out).toEqual({ models: ['a'], defaultModel: '' })
  })

  it('ignores a non-numeric or zero ceiling', async () => {
    stub({ data: [{ id: 'm', max_output_tokens: 0, context_length: 'big' }] })
    const out = await listCustomModels('https://gw.example.com/v1')
    expect(out.profiles).toBeUndefined()
  })

  it('every field the probe derives survives the settings round trip', async () => {
    // regression: contextWindow was emitted by the probe but the normalizer
    // never parsed it — a spread hides the gap from the excess-property check,
    // so only a round trip catches it
    stub({
      data: [
        {
          id: 'm',
          name: 'M',
          context_length: 1_000_000,
          max_output_tokens: 128_000,
          supports_images: true,
        },
      ],
    })
    const live = await listCustomModels('https://gw.example.com/v1', 'k')
    const normalized = normalizeCustomProfile({ probedProfiles: live.profiles })!.probedProfiles
    expect(normalized).toEqual([
      { id: 'm', name: 'M', vision: true, maxTokens: 128_000, contextWindow: 1_000_000 },
    ])
  })
})

describe('probedProfiles: the cache beside the id list', () => {
  const probed: AiCustomModelProfile[] = [
    { id: 'a', vision: true, maxTokens: 8192 },
    { id: 'b', vision: false },
  ]

  it('is replaced wholesale, never merged into the declarations', () => {
    // a probe is a snapshot: the next one overwrites it, so a model the server
    // stops serving stops being described
    const first = normalizeCustomProfile({ probedProfiles: probed })!.probedProfiles
    expect(first).toEqual(probed)
    const second = normalizeCustomProfile({
      probedProfiles: [{ id: 'c', vision: true }],
    })!.probedProfiles
    expect(second).toEqual([{ id: 'c', vision: true }])
  })

  it('leaves the user declarations untouched', () => {
    const out = normalizeCustomProfile({
      modelProfiles: [{ id: 'a', vision: false, maxTokens: 4096 }],
      probedProfiles: probed,
    })!
    expect(out.modelProfiles).toEqual([{ id: 'a', vision: false, maxTokens: 4096 }])
    expect(out.probedProfiles).toEqual(probed)
  })

  it('an empty probe leaves the cache absent, not an empty array', () => {
    expect(normalizeCustomProfile({ probedProfiles: [] })?.probedProfiles).toBeUndefined()
    expect(normalizeCustomProfile({ probedProfiles: 'nope' })?.probedProfiles).toBeUndefined()
  })

  it('caps at the stored-profile bound', () => {
    const many: AiCustomModelProfile[] = Array.from({ length: 250 }, (_, i) => ({ id: `m${i}` }))
    expect(normalizeCustomProfile({ probedProfiles: many })!.probedProfiles).toHaveLength(200)
  })
})
