/**
 * Resolve an OPC relationship target against the part that owns the .rels.
 * Some Windows producers emit backslash separators; OPC uses forward slashes,
 * so normalize before splitting. `..` is clamped at the zip root so a hostile
 * `../../..` chain cannot read as a deeper traversal than root.
 */
export function resolveTarget(basePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = basePart.slice(0, basePart.lastIndexOf('/')).split('/').filter(Boolean)
  for (const seg of target.replace(/\\/g, '/').split('/')) {
    if (seg === '.' || seg === '') continue
    if (seg === '..') {
      if (parts.length > 0) parts.pop()
    } else parts.push(seg)
  }
  return parts.join('/')
}
