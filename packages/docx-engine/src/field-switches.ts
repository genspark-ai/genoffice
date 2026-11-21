/** One `\x` switch of a field instruction, with the argument that follows it. */
export interface FieldSwitch {
  name: string
  arg?: string
}

export interface FieldInstruction {
  /** the field keyword (TOC, INDEX, TOA), upper case */
  keyword: string
  switches: FieldSwitch[]
}

const TOKEN = /"((?:[^"\\]|\\.)*)"|(\S+)/g

/**
 * A field instruction as keyword + switches. A quoted token is always the
 * argument of the switch before it; a bare token is one unless it starts
 * with a backslash.
 */
export function parseFieldInstruction(instr: string): FieldInstruction {
  const out: FieldInstruction = { keyword: '', switches: [] }
  let m: RegExpExecArray | null
  TOKEN.lastIndex = 0
  while ((m = TOKEN.exec(instr)) !== null) {
    const quoted = m[1]
    const bare = m[2]
    if (quoted === undefined && bare.startsWith('\\') && bare.length > 1) {
      out.switches.push({ name: bare.slice(1) })
      continue
    }
    const value = quoted !== undefined ? quoted.replace(/\\(["\\])/g, '$1') : bare
    const last = out.switches[out.switches.length - 1]
    if (!out.keyword && out.switches.length === 0) out.keyword = value.toUpperCase()
    else if (last && last.arg === undefined) last.arg = value
  }
  return out
}

export function formatFieldInstruction(instr: FieldInstruction): string {
  const parts = instr.switches.map((s) =>
    s.arg === undefined ? `\\${s.name}` : `\\${s.name} "${s.arg.replace(/(["\\])/g, '\\$1')}"`,
  )
  return ` ${[instr.keyword, ...parts].join(' ')} `
}

/** the switch's argument; '' when it has none, undefined when the switch is absent */
export function fieldSwitchArg(instr: FieldInstruction, name: string): string | undefined {
  const found = instr.switches.find((s) => s.name === name)
  return found ? (found.arg ?? '') : undefined
}

/**
 * The instruction with one switch set (true: bare, string: with that
 * argument) or removed (false); the other switches keep their order.
 */
export function withFieldSwitch(
  instr: FieldInstruction,
  name: string,
  value: string | boolean,
): FieldInstruction {
  const rest = instr.switches.filter((s) => s.name !== name)
  if (value === false) return { ...instr, switches: rest }
  const next: FieldSwitch = value === true ? { name } : { name, arg: value }
  const at = instr.switches.findIndex((s) => s.name === name)
  if (at === -1) return { ...instr, switches: [...rest, next] }
  return { ...instr, switches: instr.switches.map((s, i) => (i === at ? next : s)) }
}

/** the text of a field's instruction runs in a paragraph fragment, up to its separator */
export function fieldInstructionOfXml(xml: string): string {
  const end = xml.search(/<w:fldChar\b[^>]*w:fldCharType="separate"/)
  const head = end === -1 ? xml : xml.slice(0, end)
  return [...head.matchAll(/<w:instrText(?:\s[^>]*)?>([^<]*)<\/w:instrText>/g)]
    .map((m) => m[1])
    .join('')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}
