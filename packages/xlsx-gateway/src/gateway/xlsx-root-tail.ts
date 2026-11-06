export interface RootTail {
  /// The root element's own extLst (a nested one inside a conditional
  /// formatting rule or a cache field never qualifies), or null.
  readonly extLst: {
    readonly open: number
    readonly close: number
    readonly selfClosing: boolean
  } | null
  /// Offset of the root's closing tag.
  readonly rootClose: number
}

const TAGS =
  /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<!DOCTYPE\b[\s\S]*?>|<\/?[A-Za-z_][\w:.-]*(?:[^>"']|"[^"]*"|'[^']*')*\/?>/g

/// Locates where a schema-last child belongs in `root`: before the root's own
/// extLst, or before the root's closing tag. Null when the part is malformed
/// or the root is self-closing.
export function rootTail(xml: string, root: string): RootTail | null {
  const stack: string[] = []
  let extLst: RootTail['extLst'] = null
  let extLstOpen = -1
  for (const token of xml.matchAll(TAGS)) {
    const tag = token[0]
    if (tag.startsWith('<!--') || tag.startsWith('<![CDATA[') || tag.startsWith('<?')) continue
    const name = /^<\/?([^\s/>]+)/.exec(tag)?.[1]
    if (!name || tag.startsWith('<!')) return null
    if (tag.startsWith('</')) {
      if (stack.pop() !== name) return null
      if (stack.length === 1 && name === 'extLst' && extLstOpen !== -1) {
        extLst = { open: extLstOpen, close: token.index, selfClosing: false }
      }
      if (stack.length === 0) return { extLst, rootClose: token.index }
      continue
    }
    if (stack.length === 0 && name !== root) return null
    if (stack.length === 1 && name === 'extLst') {
      if (tag.endsWith('/>')) {
        extLst = { open: token.index, close: token.index + tag.length, selfClosing: true }
        continue
      }
      extLstOpen = token.index
    }
    if (!tag.endsWith('/>')) stack.push(name)
  }
  return null
}

/// Inserts an `<ext>` element into the root's extLst, creating the list before
/// the root's closing tag when there is none. Null when the part is malformed.
export function insertRootExt(xml: string, root: string, element: string): string | null {
  const tail = rootTail(xml, root)
  if (tail === null) return null
  const { extLst } = tail
  if (extLst === null) {
    return xml.slice(0, tail.rootClose) + `<extLst>${element}</extLst>` + xml.slice(tail.rootClose)
  }
  if (extLst.selfClosing) {
    return xml.slice(0, extLst.open) + `<extLst>${element}</extLst>` + xml.slice(extLst.close)
  }
  return xml.slice(0, extLst.close) + element + xml.slice(extLst.close)
}
