import { describe, expect, it } from 'vitest'
import { sanitizeGeminiSchema } from '../src/gemini-schema'

describe('sanitizeGeminiSchema', () => {
  it('normalizes nullable type unions recursively', () => {
    expect(
      sanitizeGeminiSchema({
        type: 'object',
        properties: {
          values: { type: 'array', items: { type: ['number', 'null'] } },
        },
      }),
    ).toEqual({
      type: 'object',
      properties: {
        values: { type: 'array', items: { type: 'number', nullable: true } },
      },
    })
  })

  it('removes unsupported object-schema keywords at every level', () => {
    expect(
      sanitizeGeminiSchema({
        $schema: 'https://json-schema.org/draft/2020-12/schema',
        type: 'object',
        additionalProperties: false,
        properties: {
          metadata: { type: 'object', additionalProperties: { type: 'string' } },
        },
      }),
    ).toEqual({
      type: 'object',
      properties: { metadata: { type: 'object' } },
    })
  })

  it('collapses schema combinators and tuple items to one supported schema', () => {
    expect(
      sanitizeGeminiSchema({
        anyOf: [{ type: 'null' }, { type: 'array', items: [{ type: 'string' }] }],
      }),
    ).toEqual({ type: 'array', items: { type: 'string' } })
  })
})
