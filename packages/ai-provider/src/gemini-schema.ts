const GEMINI_UNSUPPORTED_SCHEMA_KEYS = [
  '$schema',
  '$defs',
  'definitions',
  'additionalProperties',
  'propertyNames',
] as const

/** Convert JSON Schema emitted by GenOffice tools to Gemini's OpenAPI subset. */
export function sanitizeGeminiSchema(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sanitizeGeminiSchema)
  if (!value || typeof value !== 'object') return value

  const schema = { ...(value as Record<string, unknown>) }
  for (const key of GEMINI_UNSUPPORTED_SCHEMA_KEYS) delete schema[key]

  if (Array.isArray(schema.type)) {
    const types = schema.type.filter((type): type is string => typeof type === 'string')
    schema.type = types.find((type) => type !== 'null') ?? 'string'
    if (types.includes('null')) schema.nullable = true
  }

  for (const keyword of ['anyOf', 'oneOf', 'allOf'] as const) {
    const alternatives = schema[keyword]
    if (!Array.isArray(alternatives)) continue
    const selected =
      alternatives.find(
        (candidate) =>
          candidate &&
          typeof candidate === 'object' &&
          (candidate as Record<string, unknown>).type !== 'null',
      ) ?? alternatives[0]
    delete schema[keyword]
    if (selected && typeof selected === 'object') {
      return sanitizeGeminiSchema({ ...schema, ...(selected as Record<string, unknown>) })
    }
  }

  if (Array.isArray(schema.items)) schema.items = schema.items[0]
  if (schema.items !== undefined) schema.items = sanitizeGeminiSchema(schema.items)

  if (schema.properties && typeof schema.properties === 'object') {
    schema.properties = Object.fromEntries(
      Object.entries(schema.properties as Record<string, unknown>).map(([name, property]) => [
        name,
        sanitizeGeminiSchema(property),
      ]),
    )
  }

  if (Array.isArray(schema.enum)) {
    schema.enum = schema.enum.map((entry) =>
      typeof entry === 'string' ? entry : JSON.stringify(entry),
    )
  }

  return schema
}
