import { describe, expect, it } from 'vitest'
import type { OpCatalog } from '../src/op-catalog'
import {
  compactSchema,
  DESCRIPTION_CHARS,
  forbiddenConstructs,
  opNamesSchema,
  opsSchemaFromCatalog,
  parseSignatureFields,
  shortDescription,
} from '../src/mcp/op-schemas'

describe('signature parser', () => {
  it('reads names, optional markers and the type notation of the guides', () => {
    const { properties, required } = parseSignatureFields(
      'target, level: 0-6, align?: "left"|"center"|null, els:[id,…], box:{x,y,cx,cy}, kind:"textbox"|<preset geometry>, delta: 1|-1, blockIndexes: number[], on?: bool',
    )
    expect(required).toEqual(['target', 'level', 'els', 'box', 'kind', 'delta', 'blockIndexes'])
    expect(properties.target).toEqual({})
    expect(properties.level).toEqual({ type: 'number' })
    expect(properties.align).toEqual({ anyOf: [{ enum: ['left', 'center'] }, { type: 'null' }] })
    expect(properties.els).toEqual({ type: 'array' })
    expect(properties.box).toEqual({
      type: 'object',
      properties: { x: {}, y: {}, cx: {}, cy: {} },
      required: ['x', 'y', 'cx', 'cy'],
    })
    expect(properties.kind).toEqual({ type: 'string' })
    expect(properties.delta).toEqual({ enum: [1, -1] })
    expect(properties.blockIndexes).toEqual({ type: 'array', items: { type: 'number' } })
    expect(properties.on).toEqual({ type: 'boolean' })
  })

  it('treats quoted examples and unknown words as strings or untyped, never as enums', () => {
    const { properties } = parseSignatureFields(
      'width?: twips|"2.5cm"|"1in", color?: "#RRGGBB"|null, patch: ChartEdit, adjust: {<gdName>:number,…}',
    )
    expect(properties.width).toEqual({ anyOf: [{ type: 'string' }, { type: 'number' }] })
    expect(properties.color).toEqual({ anyOf: [{ type: 'string' }, { type: 'null' }] })
    expect(properties.patch).toEqual({})
    expect(properties.adjust).toEqual({ type: 'object' })
  })
})

describe('compactSchema', () => {
  it('rewrites const and type arrays, flattens unions and drops constraints', () => {
    expect(
      compactSchema({
        type: 'object',
        properties: {
          op: { type: 'string', const: 'set_cell' },
          address: { type: 'string', pattern: '^[A-Z]+[0-9]+$', description: 'A1' },
          value: { type: ['string', 'number', 'null'] },
          color: { anyOf: [{ anyOf: [{ type: 'string' }, { type: 'string' }] }, { type: 'null' }] },
          kind: { type: 'string', enum: ['a', 'b'] },
          n: { oneOf: [{ type: 'integer', minimum: 0 }] },
        },
        required: ['op'],
        additionalProperties: false,
        $schema: 'x',
      }),
    ).toEqual({
      type: 'object',
      properties: {
        op: { enum: ['set_cell'] },
        address: { type: 'string' },
        value: { anyOf: [{ type: 'string' }, { type: 'number' }, { type: 'null' }] },
        color: { anyOf: [{ type: 'string' }, { type: 'null' }] },
        kind: { enum: ['a', 'b'] },
        n: { type: 'integer' },
      },
      required: ['op'],
    })
  })

  it('caps object nesting and keeps array items at the depth of the array', () => {
    const deep = {
      type: 'object',
      properties: {
        format: {
          type: 'object',
          properties: {
            fill: {
              type: 'object',
              properties: { pattern: { type: 'string' } },
              required: ['pattern'],
            },
            stops: {
              type: 'array',
              items: { type: 'object', properties: { pos: { type: 'number' } } },
            },
          },
        },
      },
    }
    const out = compactSchema(deep)
    const format = out.properties!.format!
    expect(format.properties!.fill).toEqual({ type: 'object' })
    expect(format.properties!.stops).toEqual({ type: 'array', items: { type: 'object' } })
  })
})

describe('shortDescription', () => {
  it('collapses whitespace and cuts at a word boundary', () => {
    expect(shortDescription('  two\n words ')).toBe('two words')
    const long = shortDescription('word '.repeat(40))!
    expect(long.length).toBeLessThanOrEqual(DESCRIPTION_CHARS)
    expect(long.endsWith('…')).toBe(true)
    expect(shortDescription(undefined)).toBeUndefined()
  })
})

describe('forbiddenConstructs', () => {
  it('flags the constructs some clients reject and ignores property names', () => {
    expect(
      forbiddenConstructs({
        type: 'object',
        properties: {
          format: { type: 'string' },
          const: { type: 'string' },
          a: { type: ['string', 'null'] },
          b: { const: 1 },
          c: { $ref: '#/x' },
          d: { anyOf: [{ oneOf: [{ type: 'string' }] }] },
          e: { enum: [] },
        },
      }),
    ).toEqual(['$.a.type[]', '$.b.const', '$.c.$ref', '$.d.anyOf[0].oneOf', '$.e.enum'])
  })
})

describe('opsSchemaFromCatalog', () => {
  const catalog: OpCatalog = {
    domain: 'slides',
    fingerprint: 'x',
    groups: [],
    ops: [
      {
        op: 'setText',
        group: 'text',
        signature: '{paragraphs:[{runs:[{text,bold?}],align?}]} (group children: add group)',
        doc: 'Replaces the text. More prose that is dropped.\n\n| Field | Type |',
      },
      { op: 'moveSlide', group: 'slide', signature: '{to} — 0-based destination index', doc: '' },
      { op: 'hidden', group: 'slide', signature: '{}', available: false, reason: 'ui only' },
    ],
  }

  it('emits one discriminated variant per callable op with target and a one-line description', () => {
    const schema = opsSchemaFromCatalog(catalog)
    expect(schema.type).toBe('array')
    const variants = schema.items!.anyOf!
    expect(variants.map((v) => v.properties!.op)).toEqual([
      { enum: ['setText'] },
      { enum: ['moveSlide'] },
    ])
    expect(variants[0]).toMatchObject({
      required: ['op', 'target', 'paragraphs'],
      description: 'Replaces the text.',
    })
    expect(variants[0]!.properties!.paragraphs).toEqual({
      type: 'array',
      items: {
        type: 'object',
        properties: {
          runs: {
            type: 'array',
            items: { type: 'object', properties: { text: {}, bold: {} }, required: ['text'] },
          },
          align: {},
        },
        required: ['runs'],
      },
    })
    expect(variants[1]).toMatchObject({
      required: ['op', 'to'],
      description: '0-based destination index',
    })
    expect(forbiddenConstructs(schema)).toEqual([])
    expect(opNamesSchema(catalog).items!.properties!.op).toEqual({
      enum: ['setText', 'moveSlide'],
    })
  })
})
