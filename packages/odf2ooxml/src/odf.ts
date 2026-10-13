/**
 * ODF package loading: the zip container, its two XML parts, and the body
 * lookup both converters share. A file is accepted by the shape of its
 * `office:body` (the mimetype entry lies too often to be trusted), and any
 * structural surprise — missing part, malformed XML — returns null so the
 * caller reports an unreadable file instead of converting half of one.
 */

import JSZip from 'jszip'
import { XNode, firstChild, nodeChildren, nodeName, parseOdfXml } from './xml'

export interface OdfPackage {
  contentRoots: XNode[]
  stylesRoots: XNode[] | null
  /** read one zip entry by its path as spelled in xlink:href */
  read: (path: string) => Promise<Uint8Array | null>
}

export async function loadOdfPackage(bytes: Uint8Array): Promise<OdfPackage | null> {
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(bytes)
  } catch {
    return null
  }
  const content = await zip.file('content.xml')?.async('string')
  if (!content) return null
  const contentRoots = parseOdfXml(content)
  if (!contentRoots) return null
  const stylesXml = await zip.file('styles.xml')?.async('string')
  const stylesRoots = stylesXml ? parseOdfXml(stylesXml) : null
  return {
    contentRoots,
    stylesRoots,
    read: async (path: string) => {
      const clean = path.replace(/^\.\//, '')
      const entry = zip.file(clean) ?? zip.file(decodeURIComponent(clean))
      if (!entry) return null
      try {
        return await entry.async('uint8array')
      } catch {
        return null
      }
    },
  }
}

/** the office:text / office:presentation element under office:body, or null */
export function odfBody(
  pack: OdfPackage,
  kind: 'office:text' | 'office:presentation',
): XNode | null {
  for (const root of pack.contentRoots) {
    const body = firstChild(root, 'office:body')
    if (!body) continue
    const element = firstChild(body, kind)
    if (element) return element
  }
  return null
}

/** the office:automatic-styles element of a document part (content.xml or styles.xml) */
export function automaticStyles(pack: OdfPackage): XNode[] {
  const out: XNode[] = []
  for (const root of [...pack.contentRoots, ...(pack.stylesRoots ?? [])]) {
    const node = firstChild(root, 'office:automatic-styles')
    if (node) out.push(node)
  }
  return out
}

/** element children only, with their names — the body walkers' main iteration */
export function childElements(node: XNode): Array<{ name: string; node: XNode }> {
  const out: Array<{ name: string; node: XNode }> = []
  for (const child of nodeChildren(node)) {
    const name = nodeName(child)
    if (name && name !== '#comment') out.push({ name, node: child })
  }
  return out
}
