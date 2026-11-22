/**
 * Insert > File pastes the other document as clipboard markup, which has no
 * room for what an object or a note reference really is. While that paste is
 * being parsed, the nodes the markup cannot describe are looked up here by
 * the key the markup does carry.
 */
type Attrs = Record<string, unknown>

export interface CarriedNodes {
  /** protected blocks by the index they had in their own document (`data-idx`) */
  blocks: Map<string, Attrs>
  /** note references by the note id they were given in this document, in document order */
  notes: Map<string, Attrs[]>
}

let carried: CarriedNodes | null = null

/** runs a paste with `nodes` at hand for the parse rules */
export function whileCarrying<T>(nodes: CarriedNodes, paste: () => T): T {
  carried = nodes
  try {
    return paste()
  } finally {
    carried = null
  }
}

export const carriedBlock = (index: string | null): Attrs | null =>
  (index ? carried?.blocks.get(index) : null) ?? null

/** the reference `el` stands for: marks of one note are told apart by their order in the markup */
export function carriedNoteRef(el: Element): Attrs | null {
  const id = el.getAttribute('data-note-ref')
  const marks = id ? carried?.notes.get(id) : undefined
  if (!marks) return null
  if (marks.length === 1) return marks[0]
  let root = el
  while (root.parentElement) root = root.parentElement
  const same = [...root.querySelectorAll('sup[data-note-ref]')].filter(
    (mark) => mark.getAttribute('data-note-ref') === id,
  )
  return marks[same.indexOf(el)] ?? marks[0]
}

/** attributes that say nothing about what a carried block is */
const NOT_CONTENT = new Set(['genPristine', 'groupId', 'groupBox'])

/**
 * What a carried block looked like when it arrived. While it still looks
 * that way its XML is written as it came; once edited it is saved the way
 * an object made here is.
 */
export function blockFingerprint(attrs: Attrs): string {
  // an attribute left at its default may be absent, null or false, depending on who built the node
  const set = (key: string): boolean => attrs[key] != null && attrs[key] !== false
  return JSON.stringify(
    Object.keys(attrs)
      .filter((key) => !NOT_CONTENT.has(key) && set(key))
      .sort()
      .map((key) => [key, attrs[key]]),
  )
}
