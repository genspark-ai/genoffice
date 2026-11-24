const KEY = 'aidocs.styleSets'

/** a style set the user saved: the styles of a document, kept on this computer under a name */
export interface CustomStyleSet {
  id: string
  name: string
  /** the w:style elements, as `styleSetXmlOf` keeps them */
  xml: string
}

const isSet = (value: unknown): value is CustomStyleSet =>
  typeof value === 'object' &&
  value !== null &&
  typeof (value as CustomStyleSet).id === 'string' &&
  typeof (value as CustomStyleSet).name === 'string' &&
  typeof (value as CustomStyleSet).xml === 'string'

export function customStyleSets(): CustomStyleSet[] {
  try {
    const parsed: unknown = JSON.parse(localStorage.getItem(KEY) ?? '[]')
    return Array.isArray(parsed) ? parsed.filter(isSet) : []
  } catch {
    return []
  }
}

function store(sets: CustomStyleSet[]): boolean {
  try {
    localStorage.setItem(KEY, JSON.stringify(sets))
    return true
  } catch {
    return false
  }
}

/** Keeps `xml` under `name`; a set of that name is replaced in its place. False when it could not be kept. */
export function saveCustomStyleSet(name: string, xml: string): boolean {
  const title = name.trim()
  if (!title || !xml) return false
  const sets = customStyleSets()
  const at = sets.findIndex((s) => s.name.toLowerCase() === title.toLowerCase())
  if (at >= 0) sets[at] = { ...sets[at]!, name: title, xml }
  else {
    const taken = new Set(sets.map((s) => s.id))
    let n = sets.length + 1
    while (taken.has(`set${n}`)) n++
    sets.push({ id: `set${n}`, name: title, xml })
  }
  return store(sets)
}

export function deleteCustomStyleSet(id: string): void {
  store(customStyleSets().filter((s) => s.id !== id))
}
