import { loadCatalog } from '../commands/guide'
import type { GuideDomain, JsonSchema, OpCatalog, OpEntry } from '../op-catalog'

/**
 * Typed `ops` schemas for the MCP tools, built from the catalog `guide <domain> --json`
 * prints (the sheets zod schema, the docs op registry, the pptx-ops signatures), so the
 * schema a client sees and the ops `apply` accepts come from one source. Every client
 * loads every tool schema into the model context, so the output is compact: per op the
 * `op` literal, its fields with type and required, one short description. Nothing here
 * is used to validate a call; that stays with the CLI's structured op errors.
 */

export const DESCRIPTION_CHARS = 60

/** Keys that only constrain or annotate; the CLI rejects the value with a better message. */
const DROPPED_KEYS = new Set([
  'description',
  'title',
  'default',
  'examples',
  'pattern',
  'format',
  'minLength',
  'maxLength',
  'minimum',
  'maximum',
  'exclusiveMinimum',
  'exclusiveMaximum',
  'multipleOf',
  'minItems',
  'maxItems',
  'uniqueItems',
  'propertyNames',
  'additionalProperties',
  '$schema',
  'id',
  '$id',
])

/**
 * The client-safe subset of JSON Schema: no `type` arrays, no `const`, no `$ref`, no
 * leaf descriptions; `oneOf` becomes `anyOf`, nested unions are flattened and deduplicated.
 */
/** Object nesting kept under an op; deeper objects are a bare `object` (the guide and the usage line of a rejected op have the rest). */
export const MAX_DEPTH = 2

export function compactSchema(schema: JsonSchema, depth = 0): JsonSchema {
  const variants = unionVariants(schema, depth)
  if (variants) return variants.length === 1 ? variants[0]! : { anyOf: variants }
  if (Array.isArray(schema.type)) {
    if (schema.type.length !== 1) {
      const split = dedupe(schema.type.map((t) => compactSchema({ ...schema, type: t }, depth)))
      return split.length === 1 ? split[0]! : { anyOf: split }
    }
    return compactSchema({ ...schema, type: schema.type[0] }, depth)
  }
  const out: JsonSchema = {}
  for (const [key, value] of Object.entries(schema)) {
    if (DROPPED_KEYS.has(key) || value === undefined) continue
    if (key === 'const') {
      out.enum = [value]
    } else if (key === 'properties') {
      if (depth >= MAX_DEPTH) continue
      out.properties = Object.fromEntries(
        Object.entries(value as Record<string, JsonSchema>).map(([k, v]) => [
          k,
          compactSchema(v, depth + 1),
        ]),
      )
    } else if (key === 'items') {
      out.items = compactSchema(value as JsonSchema, depth)
    } else if (key === 'required') {
      if (depth < MAX_DEPTH && (value as string[]).length) out.required = value as string[]
    } else if (key === 'enum') {
      out.enum = value as unknown[]
    } else if (key === 'type') {
      out.type = value as string
    }
  }
  if (out.enum) delete out.type
  return out
}

function unionVariants(schema: JsonSchema, depth: number): JsonSchema[] | undefined {
  const raw = schema.anyOf ?? schema.oneOf
  if (!raw) return undefined
  const flat: JsonSchema[] = []
  for (const v of raw) {
    const c = compactSchema(v, depth)
    if (c.anyOf && Object.keys(c).length === 1) flat.push(...c.anyOf)
    else flat.push(c)
  }
  return dedupe(flat)
}

function dedupe(schemas: JsonSchema[]): JsonSchema[] {
  const seen = new Set<string>()
  return schemas.filter((s) => {
    const key = JSON.stringify(s)
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}

export function shortDescription(text: string | undefined): string | undefined {
  if (!text) return undefined
  const line = text.replace(/\s+/g, ' ').trim()
  if (!line) return undefined
  if (line.length <= DESCRIPTION_CHARS) return line
  const cut = line.slice(0, DESCRIPTION_CHARS - 1)
  const space = cut.lastIndexOf(' ')
  return `${space > DESCRIPTION_CHARS / 2 ? cut.slice(0, space) : cut}…`
}

interface ParsedFields {
  properties: Record<string, JsonSchema>
  required: string[]
}

const FIELD_RE = /^([A-Za-z_][A-Za-z0-9_]*)(\?)?\s*(?:\([^()]*\))?\s*(?::\s*(.+))?$/s

/**
 * Fields of a `{a, b?: n, c: "x"|"y", d: {…}}` signature (the docs and slides guides). The
 * notation is prose-tolerant: what cannot be read becomes an untyped field or is skipped.
 */
export function parseSignatureFields(body: string): ParsedFields {
  const out: ParsedFields = { properties: {}, required: [] }
  for (const piece of splitTop(body, ',')) {
    const m = FIELD_RE.exec(piece.trim())
    if (!m) continue
    const [, name, optional, type] = m
    if (out.properties[name!]) continue
    out.properties[name!] = type ? typeSchema(type) : {}
    if (!optional) out.required.push(name!)
  }
  return out
}

/** Splits at `sep` outside braces, brackets, parentheses and quotes. */
function splitTop(text: string, sep: string): string[] {
  const parts: string[] = []
  let depth = 0
  let quote = false
  let start = 0
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!
    if (ch === '"') quote = !quote
    if (quote) continue
    if (ch === '{' || ch === '[' || ch === '(') depth++
    else if (ch === '}' || ch === ']' || ch === ')') depth--
    else if (ch === sep && depth === 0) {
      parts.push(text.slice(start, i))
      start = i + 1
    }
  }
  parts.push(text.slice(start))
  return parts.map((p) => p.trim()).filter(Boolean)
}

/** `{…}` groups at the top level of a signature, in order (a note after the last one is ignored). */
function topGroups(text: string): string[] {
  const groups: string[] = []
  let depth = 0
  let start = -1
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!
    if (ch === '{') {
      if (depth === 0) start = i
      depth++
    } else if (ch === '}') {
      depth--
      if (depth === 0 && start >= 0) {
        groups.push(text.slice(start + 1, i))
        start = -1
      }
    } else if (depth === 0 && ch === '—') break
  }
  return groups
}

const NUMBER_WORDS = new Set(['n', 'number', 'integer', 'int', 'EMU', 'pt', 'px', 'twips', 'ms'])
const STRING_WORDS = new Set(['string', 'str', 'text', 'base64', 'dataURL', 'id'])
const BOOLEAN_WORDS = new Set(['bool', 'boolean'])

function typeSchema(text: string): JsonSchema {
  const cleaned = text.trim().replace(/\s*\([^()]*\)$/, '')
  const literals: unknown[] = []
  const others: JsonSchema[] = []
  let string = false
  let unknown = false
  for (const alt of splitTop(cleaned, '|')) {
    const quoted = /^"(.*)"$/s.exec(alt)
    if (quoted) {
      if (/[<#\d]/.test(quoted[1]!)) string = true
      else literals.push(quoted[1])
    } else if (alt === 'null') others.push({ type: 'null' })
    else if (alt.startsWith('{')) others.push(objectSchema(alt))
    else if (alt.startsWith('[')) others.push(arraySchema(alt))
    else if (alt.endsWith('[]')) others.push({ type: 'array', items: typeSchema(alt.slice(0, -2)) })
    else if (/^-?\d+(\.\d+)?$/.test(alt)) literals.push(Number(alt))
    else if (/^-?\d+(\.\d+)?(\.\.|-)-?\d+(\.\d+)?$/.test(alt) || NUMBER_WORDS.has(alt))
      others.push({ type: 'number' })
    else if (BOOLEAN_WORDS.has(alt)) others.push({ type: 'boolean' })
    else if (STRING_WORDS.has(alt)) string = true
    else if (/^Record</.test(alt)) others.push({ type: 'object' })
    else unknown = true
  }
  if (unknown && others.length === 0 && (string || literals.length)) return { type: 'string' }
  if (unknown) return {}
  if (string) {
    others.unshift({ type: 'string' })
  } else if (literals.length) {
    others.unshift({ enum: literals })
  }
  const all = dedupe(others)
  if (all.length === 0) return {}
  return all.length === 1 ? all[0]! : { anyOf: all }
}

function objectSchema(text: string): JsonSchema {
  const [body] = topGroups(text)
  const fields = body === undefined ? undefined : parseSignatureFields(body)
  if (!fields || Object.keys(fields.properties).length === 0) return { type: 'object' }
  return {
    type: 'object',
    properties: fields.properties,
    ...(fields.required.length ? { required: fields.required } : {}),
  }
}

function arraySchema(text: string): JsonSchema {
  const close = text.lastIndexOf(']')
  const inner = text.slice(1, close < 0 ? undefined : close).trim()
  if (
    !inner ||
    /^[A-Za-z_][A-Za-z0-9_]*(,\s*(…|\.\.\.|[A-Za-z_]+))*$/.test(inner) ||
    /^(…|\.\.\.)$/.test(inner)
  ) {
    return { type: 'array' }
  }
  const items = typeSchema(inner)
  return Object.keys(items).length ? { type: 'array', items } : { type: 'array' }
}

/** The target forms are spelled out once, in the ops parameter description, not per op. */
const TARGET: JsonSchema = { type: 'object' }

/** Groups whose ops act on the whole deck or pick their own page; the others address target.slide (+ el). */
const SLIDES_TARGET_OPTIONAL = new Set(['deck', 'slide'])

function variant(
  entry: OpEntry,
  fields: ParsedFields,
  description: string | undefined,
): JsonSchema {
  const out: JsonSchema = {
    type: 'object',
    properties: { op: { enum: [entry.op] }, ...fields.properties },
    required: ['op', ...fields.required],
  }
  if (description) out.description = description
  return out
}

function sheetsVariant(entry: OpEntry): JsonSchema {
  const schema = compactSchema(entry.schema ?? { type: 'object' })
  const { op: _op, ...properties } = schema.properties ?? {}
  const required = (schema.required ?? []).filter((k) => k !== 'op')
  return variant(entry, { properties, required }, shortDescription(entry.note))
}

function docsVariant(entry: OpEntry): JsonSchema {
  if (entry.schema) return sheetsVariant(entry)
  const [body] = topGroups(entry.signature)
  const parsed = body === undefined ? { properties: {}, required: [] } : parseSignatureFields(body)
  const fields: ParsedFields = { properties: {}, required: [] }
  if (entry.target && entry.target !== 'none') {
    fields.properties.target = TARGET
    if (entry.target === 'required') fields.required.push('target')
  }
  for (const key of entry.keys ?? Object.keys(parsed.properties)) {
    if (key === 'target') continue
    fields.properties[key] = parsed.properties[key] ?? {}
    if (parsed.required.includes(key)) fields.required.push(key)
  }
  return variant(entry, fields, shortDescription(entry.note))
}

function slidesVariant(entry: OpEntry): JsonSchema {
  const groups = topGroups(entry.signature).map(parseSignatureFields)
  const fields: ParsedFields = { properties: { target: TARGET }, required: [] }
  if (!SLIDES_TARGET_OPTIONAL.has(entry.group)) fields.required.push('target')
  for (const g of groups) {
    for (const [key, schema] of Object.entries(g.properties)) {
      if (key === 'target' || fields.properties[key]) continue
      fields.properties[key] = schema
      if (groups.length === 1 && g.required.includes(key)) fields.required.push(key)
    }
  }
  return variant(entry, fields, shortDescription(slidesSummary(entry)))
}

/** The first sentence of the op's prose, before any field table or example. */
function slidesSummary(entry: OpEntry): string | undefined {
  const note = / — (.+)$/.exec(entry.signature)?.[1]
  const paragraph = (entry.doc ?? '').trim().split(/\n\s*\n/)[0]
  const prose = paragraph && !/^[|`#]/.test(paragraph) ? paragraph : undefined
  const sentence = prose ? /^(.+?\.)(\s|$)/s.exec(prose.replace(/\s+/g, ' '))?.[1] : undefined
  return sentence ?? prose ?? note
}

/** Op names only: the compact form `create_pptx` advertises, pointing at `slides_apply` for the fields. */
export function opNamesSchema(catalog: OpCatalog): JsonSchema {
  const names = catalog.ops.filter((e) => e.available !== false).map((e) => e.op)
  return {
    type: 'array',
    items: { type: 'object', properties: { op: { enum: names } }, required: ['op'] },
  }
}

/** `{ type: "array", items: { anyOf: [one object per callable op] } }` for a loaded catalog. */
export function opsSchemaFromCatalog(catalog: OpCatalog): JsonSchema {
  const build =
    catalog.domain === 'sheets'
      ? sheetsVariant
      : catalog.domain === 'docs'
        ? docsVariant
        : slidesVariant
  const items = catalog.ops.filter((e) => e.available !== false).map(build)
  return { type: 'array', items: { anyOf: items } }
}

export async function opsSchemaFor(domain: GuideDomain): Promise<JsonSchema> {
  return opsSchemaFromCatalog(await loadCatalog(domain))
}

const SCALAR: JsonSchema = {
  anyOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }],
}

/** `sheet_apply` cells: `{ cell, sheet?, value? | formula?, style? }`; style is format_range's format object. */
export function cellsSchema(): JsonSchema {
  return {
    type: 'array',
    items: {
      type: 'object',
      properties: {
        cell: { type: 'string' },
        sheet: { type: 'string' },
        value: SCALAR,
        formula: { type: 'string' },
        style: { type: 'object' },
      },
      required: ['cell'],
    },
  }
}

/** `create_xlsx` data: a 2-D array of cell values or `{ sheets: [{ name, rows }] }`. */
export function xlsxDataSchema(): JsonSchema {
  const rows: JsonSchema = { type: 'array', items: { type: 'array', items: SCALAR } }
  return {
    anyOf: [
      rows,
      {
        type: 'object',
        properties: {
          sheets: {
            type: 'array',
            items: {
              type: 'object',
              properties: { name: { type: 'string' }, rows },
              required: ['name', 'rows'],
            },
          },
        },
        required: ['sheets'],
      },
    ],
  }
}

export type TypedKey =
  'docs-ops' | 'sheets-ops' | 'slides-ops' | 'slides-op-names' | 'cells' | 'xlsx-data' | 'object'

export type TypedSchemas = Record<TypedKey, JsonSchema>

/** Every typed parameter schema the tool table refers to; the catalogs load once per process. */
export async function loadTypedSchemas(): Promise<TypedSchemas> {
  const [docs, sheets, slides] = await Promise.all(
    (['docs', 'sheets', 'slides'] as const).map(loadCatalog),
  )
  return {
    'docs-ops': opsSchemaFromCatalog(docs),
    'sheets-ops': opsSchemaFromCatalog(sheets),
    'slides-ops': opsSchemaFromCatalog(slides),
    'slides-op-names': opNamesSchema(slides),
    cells: cellsSchema(),
    'xlsx-data': xlsxDataSchema(),
    object: { type: 'object' },
  }
}

const FORBIDDEN = ['const', '$ref', '$defs', 'definitions', '$schema', 'oneOf', 'not', 'format']

/**
 * Constructs some MCP clients reject (Gemini refuses `type` arrays; strict OpenAI schemas
 * refuse `const` and `$ref`): every offending path, empty when the schema is clean.
 */
export function forbiddenConstructs(schema: unknown, path = '$'): string[] {
  if (Array.isArray(schema))
    return schema.flatMap((v, i) => forbiddenConstructs(v, `${path}[${i}]`))
  if (!schema || typeof schema !== 'object') return []
  const out: string[] = []
  for (const [key, value] of Object.entries(schema as Record<string, unknown>)) {
    if (FORBIDDEN.includes(key)) out.push(`${path}.${key}`)
    if (key === 'type' && Array.isArray(value)) out.push(`${path}.type[]`)
    if (key === 'enum' && (!Array.isArray(value) || value.length === 0)) out.push(`${path}.enum`)
    if (key === 'properties' && value && typeof value === 'object') {
      for (const [k, v] of Object.entries(value as Record<string, unknown>))
        out.push(...forbiddenConstructs(v, `${path}.${k}`))
    } else if (key !== 'enum' && key !== 'required')
      out.push(...forbiddenConstructs(value, `${path}.${key}`))
  }
  return out
}
