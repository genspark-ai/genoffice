import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentToolDef } from '@genoffice/agent-core'
import { toGeminiSchema } from '../src/protocols/gemini-schema'
import { streamForProvider } from '../src/stream'
import { okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('toGeminiSchema', () => {
  it('passes a plain object schema through unchanged', () => {
    const schema = {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'keywords' },
        maxResults: { type: 'integer', description: 'default 6' },
      },
      required: ['query'],
    }
    expect(toGeminiSchema(schema)).toEqual(schema)
  })

  it('keeps empty parameter objects as-is (accepted by the API)', () => {
    expect(toGeminiSchema({ type: 'object', properties: {}, required: [] })).toEqual({
      type: 'object',
      properties: {},
      required: [],
    })
  })

  it('turns a nullable type union into type + nullable (docs insert_chart / edit_chart)', () => {
    const out = toGeminiSchema({
      type: 'object',
      properties: {
        categories: { type: 'array', items: { type: ['string', 'null'] } },
        series: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string' },
              values: { type: 'array', items: { type: ['number', 'null'] } },
            },
            required: ['values'],
          },
        },
      },
    })
    expect(out.properties).toEqual({
      categories: { type: 'array', items: { type: 'string', nullable: true } },
      series: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            name: { type: 'string' },
            values: { type: 'array', items: { type: 'number', nullable: true } },
          },
          required: ['values'],
        },
      },
    })
  })

  it('turns a multi-type union into anyOf branches', () => {
    expect(
      toGeminiSchema({ type: ['string', 'number', 'null'], description: 'id or index' }),
    ).toEqual({
      description: 'id or index',
      nullable: true,
      anyOf: [{ type: 'string' }, { type: 'number' }],
    })
    // a bare null type keeps nullable but has nothing to type against
    expect(toGeminiSchema({ type: ['null'] })).toEqual({ nullable: true })
  })

  it('inlines local $ref targets and drops the definitions block (slides paragraphs)', () => {
    const paragraphs = {
      type: 'array',
      description: 'Complete paragraph list',
      items: {
        type: 'object',
        properties: { text: { type: 'string' }, bold: { type: 'boolean' } },
        required: ['text'],
      },
    }
    const out = toGeminiSchema({
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        paragraphs: { $ref: '#/definitions/paragraphs' },
        alt: { $ref: '#/$defs/paragraphs', description: 'override wins' },
      },
      required: ['slideIndex', 'paragraphs'],
      definitions: { paragraphs },
      $defs: { paragraphs },
    })
    expect(out).toEqual({
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        paragraphs,
        alt: { ...paragraphs, description: 'override wins' },
      },
      required: ['slideIndex', 'paragraphs'],
    })
    expect(JSON.stringify(out)).not.toMatch(/\$ref|definitions|\$defs/)
  })

  it('does not loop on unresolvable or cyclic $ref', () => {
    expect(toGeminiSchema({ $ref: '#/definitions/missing' })).toEqual({})
    expect(toGeminiSchema({ $ref: 'https://example.com/remote.json' })).toEqual({})
    const out = toGeminiSchema({
      type: 'object',
      properties: { node: { $ref: '#/definitions/node' } },
      definitions: {
        node: {
          type: 'object',
          properties: { child: { $ref: '#/definitions/node' } },
        },
      },
    })
    expect(out).toEqual({
      type: 'object',
      properties: { node: { type: 'object', properties: { child: {} } } },
    })
  })

  it('strips keywords the proto does not know', () => {
    expect(
      toGeminiSchema({
        $schema: 'http://json-schema.org/draft-07/schema#',
        type: 'object',
        additionalProperties: false,
        properties: {
          ops: { type: 'array', items: { type: 'object', additionalProperties: true } },
          n: { type: 'number', exclusiveMinimum: 0, multipleOf: 2 },
        },
      }),
    ).toEqual({
      type: 'object',
      properties: {
        ops: { type: 'array', items: { type: 'object' } },
        n: { type: 'number' },
      },
    })
  })

  it('maps const / oneOf / numeric enum members onto proto equivalents', () => {
    expect(
      toGeminiSchema({
        type: 'object',
        properties: {
          kind: { type: 'string', const: 'bar' },
          level: { type: 'integer', enum: [1, 2, 3] },
          target: { oneOf: [{ type: 'string' }, { type: ['integer', 'null'] }] },
        },
      }).properties,
    ).toEqual({
      kind: { type: 'string', enum: ['bar'] },
      level: { type: 'integer', enum: ['1', '2', '3'] },
      target: { anyOf: [{ type: 'string' }, { type: 'integer', nullable: true }] },
    })
  })

  it('collapses tuple items and boolean sub-schemas into objects', () => {
    expect(
      toGeminiSchema({
        type: 'object',
        properties: {
          pair: { type: 'array', items: [{ type: 'string' }, { type: 'number' }] },
          anything: true,
        },
      }).properties,
    ).toEqual({
      pair: { type: 'array', items: { anyOf: [{ type: 'string' }, { type: 'number' }] } },
      anything: {},
    })
  })

  it('does not mutate the input schema', () => {
    const schema = {
      type: 'object',
      properties: { a: { type: ['string', 'null'], enum: [1] } },
      definitions: {},
    }
    const snapshot = JSON.stringify(schema)
    toGeminiSchema(schema)
    expect(JSON.stringify(schema)).toBe(snapshot)
  })
})

describe('gemini request body', () => {
  it('sends translated function declarations, not the raw JSON Schema', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream([])))
    vi.stubGlobal('fetch', fetchMock)
    const tools: AgentToolDef[] = [
      {
        name: 'edit_chart',
        description: 'edit',
        inputSchema: {
          type: 'object',
          properties: { categories: { type: 'array', items: { type: ['string', 'null'] } } },
          required: ['categories'],
        },
      },
    ]
    // the empty fixture stream legitimately rejects with "returned no content";
    // this test only inspects the outgoing request body
    await streamForProvider(
      'gemini',
      { apiKey: 'k', model: 'gemini-test' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      tools,
      1024,
      { signal: new AbortController().signal, onDelta: () => {}, onToolCall: () => {} },
    ).catch(() => {})
    const body = JSON.parse(fetchMock.mock.calls[0][1].body as string)
    expect(body.tools).toEqual([
      {
        functionDeclarations: [
          {
            name: 'edit_chart',
            description: 'edit',
            parameters: {
              type: 'object',
              properties: {
                categories: { type: 'array', items: { type: 'string', nullable: true } },
              },
              required: ['categories'],
            },
          },
        ],
      },
    ])
  })
})
