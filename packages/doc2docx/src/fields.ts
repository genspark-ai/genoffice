/**
 * Stage 3c-4: HYPERLINK and page-number fields in the body text.
 *
 * A Word field is \x13 instruction \x14 result \x15 (the separator is
 * optional for fields without a result). This module walks the raw body text,
 * keeps a stack for nested fields, turns HYPERLINK instructions into link
 * descriptors over the result CP range and reports field types it cannot
 * import yet; the result text stays in the body either way.
 */

export interface HyperlinkField {
  /** first CP of the field result (just after the separator) */
  startCp: number
  /** one past the last CP of the field result (the field-end CP) */
  endCp: number
  /** external URL; absent for bookmark-only links */
  target?: string
  /** bookmark anchor for \\l links inside the document */
  anchor?: string
  /** \\o tooltip text */
  tooltip?: string
}

/** A field whose result is kept as a live field in the docx (PAGE, NUMPAGES, SECTIONPAGES). */
export interface SimpleField {
  /** first CP of the field result */
  startCp: number
  /** one past the last CP of the field result */
  endCp: number
  /** the cleaned instruction, e.g. "PAGE" or "PAGE \\* MERGEFORMAT" */
  instruction: string
}

export interface FieldParseResult {
  hyperlinks: HyperlinkField[]
  simpleFields: SimpleField[]
  warnings: string[]
}

const LIVE_FIELD_TYPES = new Set(['PAGE', 'NUMPAGES', 'SECTIONPAGES'])

const FIELD_BEGIN = '\u0013'
const FIELD_SEPARATOR = '\u0014'
const FIELD_END = '\u0015'
/** HYPERLINK switches that take an argument */
const ARGUMENT_SWITCHES = new Set(['l', 'o', 't'])

interface FieldFrame {
  beginCp: number
  separatorCp?: number
  instruction: string
}

/** drop field/control characters from an instruction (spaces and tabs survive) */
function cleanInstruction(instruction: string): string {
  let out = ''
  for (const char of instruction) {
    const code = char.charCodeAt(0)
    if (code < 0x20 && char !== '\t' && char !== ' ') continue
    out += char === '\t' ? ' ' : char
  }
  return out.trim()
}

/** read one HYPERLINK argument: a quoted string or a bare token */
function readArgument(text: string, start: number): { value: string; next: number } | undefined {
  let index = start
  while (index < text.length && text[index] === ' ') index++
  if (index >= text.length) return undefined
  const quote = text[index]
  if (quote === '"' || quote === "'") {
    const end = text.indexOf(quote, index + 1)
    if (end < 0) return undefined
    return { value: text.slice(index + 1, end), next: end + 1 }
  }
  const end = text.indexOf(' ', index)
  const next = end < 0 ? text.length : end
  return { value: text.slice(index, next), next }
}

interface HyperlinkInstruction {
  target?: string
  anchor?: string
  tooltip?: string
}

/** parse a cleaned 'HYPERLINK ...' instruction into its target/anchor/tooltip */
function parseHyperlinkInstruction(instruction: string): HyperlinkInstruction | undefined {
  const match = /^HYPERLINK\b/i.exec(instruction)
  if (!match) return undefined
  const rest = instruction.slice(match[0].length)
  let target: string | undefined
  let anchor: string | undefined
  let tooltip: string | undefined
  let index = 0
  while (index < rest.length) {
    while (index < rest.length && rest[index] === ' ') index++
    if (index >= rest.length) break
    if (rest[index] === '\\') {
      const switchChar = rest[index + 1]?.toLowerCase()
      index += 2
      if (switchChar === undefined) break
      if (!ARGUMENT_SWITCHES.has(switchChar)) continue
      const argument = readArgument(rest, index)
      if (!argument) break
      index = argument.next
      if (switchChar === 'l') anchor = anchor ?? argument.value
      else if (switchChar === 'o') tooltip = tooltip ?? argument.value
      else if (switchChar === 't') target = target ?? argument.value
      continue
    }
    const argument = readArgument(rest, index)
    if (!argument) break
    index = argument.next
    if (target === undefined && !argument.value.startsWith('\\')) target = argument.value
  }
  if (target === undefined && anchor === undefined) return undefined
  return {
    ...(target !== undefined ? { target } : {}),
    ...(anchor !== undefined ? { anchor } : {}),
    ...(tooltip !== undefined ? { tooltip } : {}),
  }
}

/**
 * Parse the fields of the body text. Hyperlinks are described by their result
 * CP range; unsupported field types and malformed fields are warnings and keep
 * their result as plain text.
 */
export function parseFields(rawText: string): FieldParseResult {
  const hyperlinks: HyperlinkField[] = []
  const simpleFields: SimpleField[] = []
  const warnings: string[] = []
  const unsupportedTypes = new Set<string>()
  const stack: FieldFrame[] = []
  const finishField = (frame: FieldFrame, endCp: number): void => {
    const instruction = cleanInstruction(frame.instruction)
    const type = instruction.split(/\s+/)[0]?.toUpperCase() ?? ''
    if (frame.separatorCp === undefined) {
      if (type !== '')
        warnings.push('a ' + type + ' field has no result separator and was imported as plain text')
      return
    }
    if (LIVE_FIELD_TYPES.has(type)) {
      // an empty result (field never updated) still becomes a live field
      simpleFields.push({ startCp: frame.separatorCp + 1, endCp, instruction })
      return
    }
    if (!/^HYPERLINK\b/i.test(instruction)) {
      if (type !== '' && !unsupportedTypes.has(type)) {
        unsupportedTypes.add(type)
        warnings.push(
          'the ' + type + ' field is not supported yet; its result text is imported as plain text',
        )
      }
      return
    }
    const parsed = parseHyperlinkInstruction(instruction)
    if (!parsed) {
      warnings.push('a HYPERLINK field is malformed and was imported as plain text')
      return
    }
    const startCp = frame.separatorCp + 1
    if (endCp <= startCp) return
    hyperlinks.push({ startCp, endCp, ...parsed })
  }
  for (let cp = 0; cp < rawText.length; cp++) {
    const char = rawText[cp]
    if (char === FIELD_BEGIN) {
      stack.push({ beginCp: cp, instruction: '' })
      continue
    }
    if (char === FIELD_SEPARATOR) {
      const frame = stack[stack.length - 1]
      if (frame && frame.separatorCp === undefined) frame.separatorCp = cp
      continue
    }
    if (char === FIELD_END) {
      const frame = stack.pop()
      if (frame) finishField(frame, cp)
      continue
    }
    const frame = stack[stack.length - 1]
    if (frame && frame.separatorCp === undefined) frame.instruction += char
  }
  if (stack.length > 0)
    warnings.push('the document ends inside a field; its text was imported as plain text')
  return { hyperlinks, simpleFields, warnings: [...new Set(warnings)] }
}

/** true when the CP lies inside a field range (instruction or result) */
export function cpInField(rawText: string, cp: number): boolean {
  let depth = 0
  for (let index = 0; index < cp && index < rawText.length; index++) {
    const char = rawText[index]
    if (char === FIELD_BEGIN) depth++
    else if (char === FIELD_END) depth = Math.max(0, depth - 1)
  }
  return depth > 0
}
