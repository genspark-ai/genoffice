/**
 * Size-bounded serialization for layout-script output (log entries, return
 * values). `JSON.stringify` always materializes the whole string before a
 * truncating slice can run, so a doubling chain like `{x:o, y:o}` (2^20 shared
 * references build in ~20 object literals) forced a multi-megabyte, hundred-ms
 * stringify for a few thousand visible characters. These writers stop walking
 * the value as soon as the character budget is spent (or the depth cap is hit)
 * and mark the output as truncated, mirroring the shape of
 * `slice(0, max) + '…(truncated)'` without ever building the full text.
 */

export const TRUNCATED_MARKER = '…(truncated)'

/** Plain string truncation with the same marker shape the serializer uses. */
export function truncateString(text: string, maxChars: number): string {
  return text.length > maxChars ? text.slice(0, maxChars) + TRUNCATED_MARKER : text
}

/**
 * JSON.stringify that never produces more than `maxChars` characters (plus the
 * truncation marker) and never walks deeper than `maxDepth` levels. Returns
 * `undefined` for `undefined`, like `JSON.stringify`.
 */
export function boundedJsonStringify(
  value: unknown,
  maxChars: number,
  maxDepth = 64,
): string | undefined {
  let out = ''
  let truncated = false

  const write = (text: string): void => {
    if (out.length >= maxChars) {
      truncated = true
      return
    }
    out += out.length + text.length > maxChars ? text.slice(0, maxChars - out.length) : text
  }

  const quote = (text: string): void => {
    write('"')
    for (let i = 0; i < text.length; i++) {
      if (out.length >= maxChars) {
        truncated = true
        return
      }
      const code = text.charCodeAt(i)
      if (code === 0x22) write('\\"')
      else if (code === 0x5c) write('\\\\')
      else if (code === 0x08) write('\\b')
      else if (code === 0x09) write('\\t')
      else if (code === 0x0a) write('\\n')
      else if (code === 0x0c) write('\\f')
      else if (code === 0x0d) write('\\r')
      else if (code < 0x20) write(`\\u${code.toString(16).padStart(4, '0')}`)
      else if (code >= 0xd800 && code <= 0xdbff) {
        const next = text.charCodeAt(i + 1)
        if (next >= 0xdc00 && next <= 0xdfff) {
          // valid surrogate pair: emit both halves raw, as JSON.stringify does
          write(text.slice(i, i + 2))
          i += 1
        } else write(`\\u${code.toString(16).padStart(4, '0')}`)
      } else if (code >= 0xdc00 && code <= 0xdfff) {
        write(`\\u${code.toString(16).padStart(4, '0')}`)
      } else write(text[i])
    }
    write('"')
  }

  const emit = (item: unknown, depth: number): void => {
    if (out.length >= maxChars || depth > maxDepth) {
      truncated = true
      return
    }
    switch (typeof item) {
      case 'string':
        quote(item)
        return
      case 'number':
        write(Number.isFinite(item) ? String(item) : 'null')
        return
      case 'boolean':
        write(item ? 'true' : 'false')
        return
      case 'object': {
        if (item === null) {
          write('null')
          return
        }
        if (Array.isArray(item)) {
          write('[')
          item.forEach((element, index) => {
            if (index > 0) write(',')
            emit(element, depth + 1)
          })
          write(']')
          return
        }
        write('{')
        let first = true
        for (const [key, element] of Object.entries(item as Record<string, unknown>)) {
          // JSON.stringify omits undefined/function-valued properties
          if (
            element === undefined ||
            typeof element === 'function' ||
            typeof element === 'symbol' ||
            typeof element === 'bigint'
          )
            continue
          if (!first) write(',')
          first = false
          quote(key)
          write(':')
          emit(element, depth + 1)
        }
        write('}')
        return
      }
      default:
        // function/symbol/bigint: treated like JSON.stringify's `undefined`
        write('null')
    }
  }

  if (value === undefined) return undefined
  emit(value, 0)
  return truncated ? out + TRUNCATED_MARKER : out
}
