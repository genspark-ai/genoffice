import type { FontTableEntry } from './types'
import { attrsOf, childrenOf, findChild, nameOf, xmlParser, type XNode } from './xml-utils'

export const FONT_TABLE_PART_PATH = 'word/fontTable.xml'

/** word/fontTable.xml → substitution hints (altName / PANOSE / family / pitch) */
export function parseFontTable(xml: string): FontTableEntry[] {
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(xml) as XNode[]
  } catch {
    return []
  }
  const root = parsed.find((n) => nameOf(n) === 'w:fonts')
  if (!root) return []
  const out: FontTableEntry[] = []
  for (const node of childrenOf(root)) {
    if (nameOf(node) !== 'w:font') continue
    const name = attrsOf(node)['w:name']
    if (!name) continue
    const val = (tag: string) => attrsOf(findChild(node, tag) ?? {})['w:val'] || undefined
    const entry: FontTableEntry = { name }
    const altName = val('w:altName')
    if (altName) entry.altName = altName
    const panose = val('w:panose1')
    if (panose) entry.panose = panose
    const family = val('w:family')
    if (family) entry.family = family
    const pitch = val('w:pitch')
    if (pitch) entry.pitch = pitch
    out.push(entry)
  }
  return out
}
