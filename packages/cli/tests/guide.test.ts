import { describe, expect, it } from 'vitest'
import { OP_DOCS } from '@genoffice/pptx-ops/op-docs'
import { fingerprint, signatureFromSchema } from '../src/op-catalog'
import { SHEET_OP_GROUPS, sheetCatalog } from '../src/formats/xlsx-catalog'
import { SUPPORTED_DSL_OPS } from '../src/formats/xlsx-dsl'
import { REFUSED_DSL_OPS } from '../src/formats/xlsx-gateway-ops'
import { run } from './helpers'

describe('op catalogs', () => {
  it('lists every supported sheet op in exactly one group and every refused op as unavailable', () => {
    const grouped = Object.values(SHEET_OP_GROUPS).flat()
    expect([...grouped].sort()).toEqual([...SUPPORTED_DSL_OPS].sort())
    expect(new Set(grouped).size).toBe(grouped.length)
    const catalog = sheetCatalog()
    for (const op of SUPPORTED_DSL_OPS) {
      const entry = catalog.ops.find((e) => e.op === op)!
      expect(entry.available).toBeUndefined()
      expect(entry.schema?.properties).not.toHaveProperty('sheetId')
    }
    for (const op of Object.keys(REFUSED_DSL_OPS)) {
      expect(catalog.ops.find((e) => e.op === op)).toMatchObject({ available: false })
    }
    expect(catalog.fingerprint).toMatch(/^[0-9a-f]{12}$/)
    expect(catalog.fingerprint).toBe(sheetCatalog().fingerprint)
  })

  it('renders signatures from a schema', () => {
    const sig = signatureFromSchema({
      type: 'object',
      properties: {
        op: { const: 'x' },
        range: { type: 'string' },
        count: { type: 'integer' },
        order: { enum: ['asc', 'desc'] },
        values: { type: 'array', items: { anyOf: [{ type: 'string' }, { type: 'null' }] } },
        on: { anyOf: [{ type: 'boolean' }, { type: 'null' }] },
        colors: { type: 'object', additionalProperties: { type: 'string' } },
      },
      required: ['op', 'range'],
    })
    expect(sig).toBe(
      '{range, count?: n, order?: asc|desc, values?: [string|null], on?: bool|null, colors?: {"<key>": string}}',
    )
    expect(fingerprint({ b: 1, a: [2] })).toBe(fingerprint({ a: [2], b: 1 }))
  })

  it('prints the sheets guide from the schema, one group, or one op with its schema', async () => {
    const all = await run(['guide', 'sheets'])
    expect(all.code).toBe(0)
    for (const op of SUPPORTED_DSL_OPS) expect(all.stdout).toContain(`  ${op} `)
    expect(all.stdout).toContain('refresh_pivot — not available headless')
    expect(all.stdout).toContain('format: {bold?: bool|null')
    const layout = await run(['guide', 'sheets', 'layout'])
    expect(layout.stdout).toContain('set_freeze')
    expect(layout.stdout).not.toContain('set_cell')
    const one = await run(['guide', 'sheets', 'sort_range', '--json'])
    const d = one.json().detail
    expect(d.op).toBe('sort_range')
    expect(d.schema.properties.order.enum).toEqual(['asc', 'desc'])
    expect(d.schema.required).toEqual(['op', 'range', 'byColumn', 'order'])
    const copy = (await run(['guide', 'sheets', 'copy_range', '--json'])).json().detail
    expect(copy.schema.properties).toHaveProperty('sourceSheet')
    expect(copy.schema.required).not.toContain('sourceSheetId')
    expect(copy.schema.required).not.toContain('sheetId')
    const chart = (await run(['guide', 'sheets', 'edit_chart', '--json'])).json().detail
    expect(chart.schema.properties).not.toHaveProperty('seriesData')
    expect(chart.signature).not.toContain('seriesData')
    const json = await run(['guide', 'sheets', '--json'])
    expect(json.json().detail.ops.length).toBe(sheetCatalog().ops.length)
    expect(json.json().detail.fingerprint).toBe(sheetCatalog().fingerprint)
    const typo = await run(['guide', 'sheets', 'set_cel', '--json'])
    expect(typo.code).toBe(1)
    expect(typo.json()).toMatchObject({
      error: 'invalid_argument',
      suggestion: 'did you mean set_cell?',
    })
  })

  it('exposes the slides ops as a catalog and prints a single op', async () => {
    const json = await run(['guide', 'slides', '--json'])
    const detail = json.json().detail
    const callable = Object.entries(OP_DOCS)
      .filter(([, d]) => d.aiCallable !== false && !d.pending)
      .map(([n]) => n)
    expect(detail.ops.map((e: { op: string }) => e.op).sort()).toEqual(callable.sort())
    const one = await run(['guide', 'slides', 'setText'])
    expect(one.code).toBe(0)
    expect(one.stdout).toContain(`  setText ${OP_DOCS.setText!.sig}`)
    expect(one.stdout).toContain(OP_DOCS.setText!.body.slice(0, 40))
  })

  it('builds the docs catalog from the op registry and the tool schemas', async () => {
    const json = await run(['guide', 'docs', '--json'])
    expect(json.code).toBe(0)
    const detail = json.json().detail
    expect(detail.groups.map((g: { name: string }) => g.name)).toEqual(['block', 'content'])
    const byName = new Map(detail.ops.map((e: { op: string }) => [e.op, e]))
    expect(byName.get('findReplace')).toMatchObject({ group: 'block', target: 'optional' })
    expect((byName.get('findReplace') as { keys: string[] }).keys).toContain('replace')
    expect((byName.get('setFont') as { signature: string }).signature).not.toContain('"setFont"')
    expect(byName.get('set_header_footer')).toMatchObject({
      group: 'content',
      signature: '{kind: header|footer, text, view?: default|first|even}',
    })
    expect(
      (byName.get('insert_content') as { schema: any }).schema.properties.afterBlockIndex
        .description,
    ).toContain('end of document')
    expect((byName.get('insert_image') as { schema: any }).schema.properties).toHaveProperty(
      'afterBlockIndex',
    )
    const text = await run(['guide', 'docs', 'content'])
    expect(text.stdout).toContain('insert_chart {kind: bar|line|pie')
    expect(text.stdout).toContain('Only these tags are allowed')
    expect(text.stdout).not.toContain('setFont')
  })

  it('prints fingerprints for all domains', async () => {
    const fp = await run(['guide', '--fingerprint', '--json'])
    expect(fp.code).toBe(0)
    const { fingerprints } = fp.json().detail
    expect(Object.keys(fingerprints)).toEqual(['slides', 'docs', 'sheets'])
    for (const v of Object.values(fingerprints)) expect(v).toMatch(/^[0-9a-f]{12}$/)
    const typo = await run(['guide', 'sheet', '--fingerprint', '--json'])
    expect(typo.code).toBe(1)
    expect(typo.json().error).toBe('invalid_argument')
  })
})
