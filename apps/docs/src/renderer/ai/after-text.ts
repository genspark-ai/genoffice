import type { Mark, Node as PmNode } from '@tiptap/pm/model'

const FIELD_MARKS = new Set(['instrField', 'refField'])

/** right after the first `afterText` in the live text whose end is not inside a field result (inserting there would split the field) */
export function posAfterText(
  block: PmNode,
  contentStart: number,
  afterText: string,
  blockLabel: string,
): { pos: number } | { error: string } {
  const chars: number[] = []
  const fields: Array<Mark | undefined> = []
  let text = ''
  block.descendants((child, offset) => {
    if (!child.isText || child.marks.some((m) => m.type.name === 'del')) return true
    const field = child.marks.find((m) => FIELD_MARKS.has(m.type.name))
    const start = contentStart + offset
    for (let i = 0; i < (child.text ?? '').length; i++) {
      chars.push(start + i)
      fields.push(field)
    }
    text += child.text ?? ''
    return true
  })
  let splits: Mark | undefined
  for (let at = text.indexOf(afterText); at >= 0; at = text.indexOf(afterText, at + 1)) {
    const end = at + afterText.length - 1
    const field = fields[end]
    const next = fields[end + 1]
    if (field && next && field.eq(next)) {
      splits ??= field
      continue
    }
    return { pos: chars[end]! + 1 }
  }
  if (splits) {
    const label = splits.type.name === 'refField' ? `REF ${splits.attrs.name}` : splits.attrs.instr
    return {
      error: `afterText "${afterText}" ends inside the field { ${label} }; include the whole field result or choose text outside it`,
    }
  }
  return {
    error: `afterText "${afterText}" does not occur in ${blockLabel} ("${text.slice(0, 60)}")`,
  }
}
