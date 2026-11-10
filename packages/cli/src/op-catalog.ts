import { createHash } from 'node:crypto'

export type GuideDomain = 'slides' | 'docs' | 'sheets'

/** The subset of JSON Schema the op catalogs use (zod's output and the docs tool schemas). */
export interface JsonSchema {
  type?: string | string[]
  const?: unknown
  enum?: unknown[]
  properties?: Record<string, JsonSchema>
  required?: string[]
  items?: JsonSchema
  anyOf?: JsonSchema[]
  oneOf?: JsonSchema[]
  description?: string
  [key: string]: unknown
}

export interface OpEntry {
  op: string
  group: string
  /** the fields as an agent writes them; printed after the op name */
  signature: string
  note?: string
  /** false: the app accepts the op but the CLI cannot run it headless */
  available?: false
  reason?: string
  /** input schema of the op object (sheets, docs content ops) */
  schema?: JsonSchema
  /** docs block ops: whether `target` is required, optional or refused */
  target?: 'required' | 'optional' | 'none'
  keys?: string[]
  /** slides: the group guide's markdown for this op */
  doc?: string
  examples?: string[]
}

export interface OpGroupEntry {
  name: string
  summary?: string
  ops: string[]
}

export interface OpCatalog {
  domain: GuideDomain
  fingerprint: string
  groups: OpGroupEntry[]
  ops: OpEntry[]
}

export function withFingerprint(catalog: Omit<OpCatalog, 'fingerprint'>): OpCatalog {
  return {
    domain: catalog.domain,
    fingerprint: fingerprint(catalog),
    groups: catalog.groups,
    ops: catalog.ops,
  }
}

export function fingerprint(value: unknown): string {
  return createHash('sha256').update(canonical(value)).digest('hex').slice(0, 12)
}

function canonical(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`
  if (value && typeof value === 'object') {
    const keys = Object.keys(value as Record<string, unknown>).sort()
    return `{${keys
      .map((k) => `${JSON.stringify(k)}:${canonical((value as Record<string, unknown>)[k])}`)
      .join(',')}}`
  }
  return JSON.stringify(value) ?? 'null'
}

const WIDTH = 110

/**
 * `{sheet?, range, format: {bold?: bool|null, …}}` from an object schema: strings stay bare,
 * numbers are `n`, enums list their values, unions join with `|`. Long signatures break one
 * field per line, and a union of objects one variant per line.
 */
export function signatureFromSchema(schema: JsonSchema, skip: readonly string[] = ['op']): string {
  const fields = fieldList(schema, skip)
  const line = `{${fields.join(', ')}}`
  if (line.length <= WIDTH) return line
  return `{\n${fields.map((f) => `    ${f.replace(/\}\|\{/g, '}\n        | {')}`).join(',\n')}\n  }`
}

function fieldList(schema: JsonSchema, skip: readonly string[]): string[] {
  const required = new Set(schema.required ?? [])
  return Object.entries(schema.properties ?? {})
    .filter(([key]) => !skip.includes(key))
    .map(([key, value]) => {
      const type = typeText(value)
      return `${key}${required.has(key) ? '' : '?'}${type ? `: ${type}` : ''}`
    })
}

function typeText(schema: JsonSchema, inUnion = false): string {
  if (schema.const !== undefined) return JSON.stringify(schema.const)
  if (schema.enum) return schema.enum.map(String).join('|')
  const variants = schema.anyOf ?? schema.oneOf
  if (variants) return variants.map((v) => typeText(v, true)).join('|')
  const types = Array.isArray(schema.type) ? schema.type : schema.type ? [schema.type] : []
  if (types.length === 0) return 'any'
  if (types.length > 1) return types.map((t) => scalarLabel(t, true)).join('|')
  const [type] = types
  if (type === 'array') return `[${schema.items ? typeText(schema.items, true) : 'any'}]`
  if (type === 'object') {
    if (
      !schema.properties &&
      schema.additionalProperties &&
      typeof schema.additionalProperties === 'object'
    ) {
      return `{"<key>": ${typeText(schema.additionalProperties as JsonSchema, true)}}`
    }
    return `{${fieldList(schema, []).join(', ')}}`
  }
  return scalarLabel(type!, inUnion)
}

function scalarLabel(type: string, inUnion: boolean): string {
  switch (type) {
    case 'string':
      return inUnion ? 'string' : ''
    case 'number':
    case 'integer':
      return 'n'
    case 'boolean':
      return 'bool'
    default:
      return type
  }
}

/** Group headings with one signature line per op (and its note): the text form of a catalog. */
export function renderGroups(catalog: OpCatalog, group?: string): string[] {
  const lines: string[] = []
  for (const g of catalog.groups) {
    if (group && g.name !== group) continue
    lines.push(
      `${g.name[0]!.toUpperCase()}${g.name.slice(1)}${g.summary ? ` (${g.summary})` : ''}:`,
    )
    for (const op of g.ops) lines.push(...opLines(catalog.ops.find((e) => e.op === op)!))
  }
  return lines
}

export function opLines(entry: OpEntry): string[] {
  if (entry.available === false) return [`  ${entry.op} — not available headless: ${entry.reason}`]
  const head = `  ${entry.op} ${entry.signature}`
  return entry.note ? [head, `    — ${entry.note}`] : [head]
}
