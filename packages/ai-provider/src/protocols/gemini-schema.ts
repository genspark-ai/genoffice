/**
 * JSON Schema -> Gemini `Schema` translation for function declarations.
 *
 * Tool `inputSchema`s are authored as JSON Schema (what Anthropic and the
 * OpenAI-compatible protocols consume verbatim). Gemini's
 * `FunctionDeclaration.parameters` is not JSON Schema but an OpenAPI-3.0-subset
 * protobuf, and the API front end rejects the whole request with
 * `HTTP 400 Invalid JSON payload received` when it meets anything the proto
 * lacks. Verified against generativelanguage v1beta (2026-09):
 *
 * - `type` must be a single value: `type: ['number', 'null']` -> "Proto field is
 *   not repeating, cannot start list". Nullability is `nullable: true`.
 * - `$ref` / `definitions` / `$defs` / `additionalProperties` / `const` ->
 *   "Unknown name ...: Cannot find field".
 * - `enum` is a repeated string: numeric members -> "Invalid value (TYPE_STRING)".
 * - `items` must be one schema, not a tuple list.
 * - Empty `properties: {}` and property-less `{ type: 'object' }` are accepted,
 *   as are `anyOf`, `oneOf` and any `format` value.
 */

type Json = Record<string, unknown>

/** Fields the Gemini `Schema` proto knows; anything else is dropped. */
const GEMINI_SCHEMA_KEYS = new Set([
  'type',
  'format',
  'title',
  'description',
  'nullable',
  'enum',
  'items',
  'properties',
  'required',
  'anyOf',
  'minimum',
  'maximum',
  'minItems',
  'maxItems',
  'minLength',
  'maxLength',
  'minProperties',
  'maxProperties',
  'pattern',
  'default',
  'example',
  'propertyOrdering',
])

/** Recursion guard for pathological `$ref` cycles (real tool schemas are shallow). */
const MAX_DEPTH = 32

function isJson(v: unknown): v is Json {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}

/** Resolve a local JSON pointer (`#/definitions/x`, `#/$defs/x/items`) against the root. */
function resolveRef(root: Json, ref: string): Json | null {
  if (!ref.startsWith('#/')) return null
  let node: unknown = root
  for (const raw of ref.slice(2).split('/')) {
    const key = raw.replace(/~1/g, '/').replace(/~0/g, '~')
    if (!isJson(node) || !(key in node)) return null
    node = node[key]
  }
  return isJson(node) ? node : null
}

function convert(schema: unknown, root: Json, depth: number, refStack: string[]): Json {
  // JSON Schema allows `true` / `false` / missing sub-schemas; Gemini wants an object
  if (!isJson(schema) || depth > MAX_DEPTH) return {}

  if (typeof schema.$ref === 'string') {
    const ref = schema.$ref
    const target = refStack.includes(ref) ? null : resolveRef(root, ref)
    if (!target) return {}
    // sibling keywords (usually a `description`) override the referenced schema
    const siblings = Object.fromEntries(Object.entries(schema).filter(([k]) => k !== '$ref'))
    return convert({ ...target, ...siblings }, root, depth + 1, [...refStack, ref])
  }

  const out: Json = {}
  for (const [key, value] of Object.entries(schema)) {
    if (GEMINI_SCHEMA_KEYS.has(key)) out[key] = value
  }

  // `type: [...]` -> single `type` (+ `nullable`), or `anyOf` when several remain
  if (Array.isArray(schema.type)) {
    const types = schema.type.filter((t): t is string => typeof t === 'string')
    const nonNull = types.filter((t) => t !== 'null')
    if (types.length !== nonNull.length) out.nullable = true
    if (nonNull.length === 1) out.type = nonNull[0]
    else {
      delete out.type
      if (nonNull.length > 1 && !Array.isArray(schema.anyOf)) {
        out.anyOf = nonNull.map((t) => ({ type: t }))
      }
    }
  }

  // `oneOf` has the same meaning for tool calling; only `anyOf` is in the proto
  if (Array.isArray(schema.oneOf) && !Array.isArray(out.anyOf)) out.anyOf = schema.oneOf
  // `const` -> single-member enum
  if (schema.const !== undefined && !Array.isArray(out.enum)) out.enum = [schema.const]
  if (Array.isArray(out.enum)) out.enum = out.enum.map((v) => String(v))

  if (isJson(out.properties)) {
    out.properties = Object.fromEntries(
      Object.entries(out.properties).map(([name, sub]) => [
        name,
        convert(sub, root, depth + 1, refStack),
      ]),
    )
  }
  if (out.items !== undefined) {
    // draft-07 tuple form `items: [a, b]` -> one schema accepting either
    out.items = Array.isArray(out.items)
      ? { anyOf: out.items.map((sub) => convert(sub, root, depth + 1, refStack)) }
      : convert(out.items, root, depth + 1, refStack)
  }
  if (Array.isArray(out.anyOf)) {
    out.anyOf = out.anyOf.map((sub) => convert(sub, root, depth + 1, refStack))
  }
  return out
}

/**
 * Translate a tool's JSON Schema `inputSchema` into a Gemini function-declaration
 * `parameters` schema. Lossy where the proto has no equivalent (dropped keywords
 * only ever make the schema more permissive); never throws.
 */
export function toGeminiSchema(inputSchema: Record<string, unknown>): Record<string, unknown> {
  return convert(inputSchema, inputSchema, 0, [])
}
