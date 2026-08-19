import type JSZip from 'jszip'

/**
 * Relationship kinds whose references live directly in document XML and whose
 * targets are safe for this engine to lifecycle-manage.
 */
const DOCUMENT_OWNED_REL_TYPES = new Set([
  'chart',
  'diagramcolors',
  'diagramdata',
  'diagramdrawing',
  'diagramlayout',
  'diagramquickstyle',
  'hyperlink',
  'image',
  'oleobject',
])

const RELATIONSHIP_TAG_RE =
  /<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*(?:\/\s*>|>\s*<\/(?:[A-Za-z_][\w.-]*:)?Relationship\s*>)/g
const OVERRIDE_TAG_RE =
  /<(?:[A-Za-z_][\w.-]*:)?Override\b[^>]*(?:\/\s*>|>\s*<\/(?:[A-Za-z_][\w.-]*:)?Override\s*>)/g

interface Relationship {
  id: string
  type: string
  target: string
  external: boolean
}

interface RelationshipEdge {
  source: string
  target: string
}

function xmlAttr(tag: string, name: string): string | undefined {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const match = new RegExp(`(?:^|\\s)${escaped}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag)
  const value = match?.[1] ?? match?.[2]
  return value === undefined ? undefined : decodeXmlEntities(value)
}

function decodeXmlEntities(value: string): string {
  const codePoint = (match: string, raw: string, radix: number) => {
    const value = parseInt(raw, radix)
    return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
      ? String.fromCodePoint(value)
      : match
  }
  return value
    .replace(/&#x([0-9a-f]+);/gi, (match, hex: string) => codePoint(match, hex, 16))
    .replace(/&#(\d+);/g, (match, decimal: string) => codePoint(match, decimal, 10))
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
}

function relationshipsOf(xml: string): Relationship[] {
  const relationships: Relationship[] = []
  for (const tag of xml.match(RELATIONSHIP_TAG_RE) ?? []) {
    const id = xmlAttr(tag, 'Id')
    const type = xmlAttr(tag, 'Type')
    const target = xmlAttr(tag, 'Target')
    if (!id || !type || !target) continue
    relationships.push({
      id,
      type,
      target,
      external: xmlAttr(tag, 'TargetMode')?.toLowerCase() === 'external',
    })
  }
  return relationships
}

function relationshipTypeName(type: string): string {
  return (type.slice(type.lastIndexOf('/') + 1) || type).toLowerCase()
}

function relationshipSourcePath(relsPath: string): string | null {
  if (relsPath === '_rels/.rels') return ''
  const match = /^(.*\/)?_rels\/([^/]+)\.rels$/.exec(relsPath)
  return match ? `${match[1] ?? ''}${match[2]}` : null
}

function relationshipPartPath(sourcePath: string): string {
  const slash = sourcePath.lastIndexOf('/')
  const dir = slash >= 0 ? sourcePath.slice(0, slash + 1) : ''
  const filename = slash >= 0 ? sourcePath.slice(slash + 1) : sourcePath
  return `${dir}_rels/${filename}.rels`
}

function decodePartUri(uri: string): string {
  try {
    return decodeURIComponent(uri)
  } catch {
    return uri
  }
}

function normalizePartPath(path: string): string | null {
  const parts: string[] = []
  for (const segment of path.replace(/\\/g, '/').split('/')) {
    if (!segment || segment === '.') continue
    if (segment === '..') {
      if (parts.length === 0) return null
      parts.pop()
    } else {
      parts.push(segment)
    }
  }
  return parts.join('/')
}

function resolveTargetPath(sourcePath: string, target: string): string | null {
  const withoutFragment = target.split('#', 1)[0]
  if (!withoutFragment || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(withoutFragment)) return null
  const decoded = decodePartUri(withoutFragment)
  if (decoded.startsWith('/')) return normalizePartPath(decoded.slice(1))
  const slash = sourcePath.lastIndexOf('/')
  const dir = slash >= 0 ? sourcePath.slice(0, slash + 1) : ''
  return normalizePartPath(dir + decoded)
}

function isOwnedPart(path: string): boolean {
  return (
    path.startsWith('word/media/') ||
    path.startsWith('word/charts/') ||
    path.startsWith('word/embeddings/') ||
    path.startsWith('word/diagrams/')
  )
}

/**
 * Find relationship ids actually referenced by the source XML. Prefixes are
 * discovered from namespace declarations so strict OOXML and non-"r" prefixes
 * remain safe. A missing declaration still gets the conventional "r" fallback.
 */
function referencedRelationshipIds(xml: string): Set<string> {
  const prefixes = new Set(['r'])
  const namespaceRe =
    /\bxmlns:([A-Za-z_][\w.-]*)\s*=\s*(?:"([^"]*\/relationships)"|'([^']*\/relationships)')/g
  let namespace: RegExpExecArray | null
  while ((namespace = namespaceRe.exec(xml)) !== null) prefixes.add(namespace[1])

  const ids = new Set<string>()
  for (const prefix of prefixes) {
    const escaped = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const attrRe = new RegExp(
      `\\b${escaped}:[A-Za-z_][\\w.-]*\\s*=\\s*(?:"([^"]*)"|'([^']*)')`,
      'g',
    )
    let attr: RegExpExecArray | null
    while ((attr = attrRe.exec(xml)) !== null) ids.add(decodeXmlEntities(attr[1] ?? attr[2]))
  }
  // Legacy VML/Office producers occasionally put a package relationship id in
  // a non-r namespace attribute (for example o:relid). Preserve any id that is
  // still present as an exact XML attribute value rather than risking damage.
  const attributeRe = /\s[A-Za-z_][\w.:-]*\s*=\s*(?:"([^"]*)"|'([^']*)')/g
  let attribute: RegExpExecArray | null
  while ((attribute = attributeRe.exec(xml)) !== null) {
    const value = decodeXmlEntities(attribute[1] ?? attribute[2])
    if (/^rId[\w.-]*$/i.test(value)) ids.add(value)
  }
  return ids
}

async function relationshipEdges(zip: JSZip): Promise<RelationshipEdge[]> {
  const edges: RelationshipEdge[] = []
  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith('.rels')) continue
    const source = relationshipSourcePath(path)
    const file = zip.file(path)
    if (source === null || !file) continue
    for (const rel of relationshipsOf(await file.async('string'))) {
      if (rel.external) continue
      const target = resolveTargetPath(source, rel.target)
      if (target) edges.push({ source, target })
    }
  }
  return edges
}

function addReachableOwnedParts(
  starts: Iterable<string>,
  outgoing: Map<string, string[]>,
  allowed: Set<string> | null = null,
): Set<string> {
  const reached = new Set<string>()
  const pending = [...starts]
  while (pending.length > 0) {
    const part = pending.pop()!
    if (reached.has(part) || (allowed && !allowed.has(part))) continue
    reached.add(part)
    for (const target of outgoing.get(part) ?? []) {
      if (isOwnedPart(target) && (!allowed || allowed.has(target))) pending.push(target)
    }
  }
  return reached
}

/**
 * Prune resources made unreachable by the final main-document XML.
 *
 * The pass is intentionally narrow:
 * - only known drawing/hyperlink relationship kinds are pruned at the document root;
 * - only owned Word resource directories are recursively removed;
 * - inbound references from any surviving part retain a shared target and its dependencies;
 * - unknown/custom parts are never selected for deletion.
 */
export async function cleanupDocxOwnedResources(zip: JSZip, documentPath: string): Promise<void> {
  const documentFile = zip.file(documentPath)
  const relsPath = relationshipPartPath(documentPath)
  const relsFile = zip.file(relsPath)
  if (!documentFile || !relsFile) return

  const documentXml = await documentFile.async('string')
  const referencedIds = referencedRelationshipIds(documentXml)
  const candidateRoots = new Set<string>()
  const relsXml = await relsFile.async('string')
  const cleanedRelsXml = relsXml.replace(RELATIONSHIP_TAG_RE, (tag) => {
    const id = xmlAttr(tag, 'Id')
    const type = xmlAttr(tag, 'Type')
    if (
      !id ||
      !type ||
      referencedIds.has(id) ||
      !DOCUMENT_OWNED_REL_TYPES.has(relationshipTypeName(type))
    ) {
      return tag
    }
    const external = xmlAttr(tag, 'TargetMode')?.toLowerCase() === 'external'
    const target = xmlAttr(tag, 'Target')
    if (!external && target) {
      const targetPath = resolveTargetPath(documentPath, target)
      if (targetPath && isOwnedPart(targetPath)) candidateRoots.add(targetPath)
    }
    return ''
  })
  if (cleanedRelsXml === relsXml) return
  zip.file(relsPath, cleanedRelsXml, { date: relsFile.date })

  const edges = await relationshipEdges(zip)
  const outgoing = new Map<string, string[]>()
  for (const edge of edges) {
    const targets = outgoing.get(edge.source)
    if (targets) targets.push(edge.target)
    else outgoing.set(edge.source, [edge.target])
  }

  const candidates = addReachableOwnedParts(
    [...candidateRoots].filter((path) => zip.file(path) !== null),
    outgoing,
  )
  for (const path of [...candidates]) {
    if (!zip.file(path)) candidates.delete(path)
  }
  if (candidates.size === 0) return

  // Any inbound edge from outside the candidate subgraph makes that target
  // shared. Retain it and every owned dependency reachable from it.
  const sharedRoots = new Set<string>()
  for (const edge of edges) {
    if (candidates.has(edge.target) && !candidates.has(edge.source)) sharedRoots.add(edge.target)
  }
  const retained = addReachableOwnedParts(sharedRoots, outgoing, candidates)
  const removed = new Set<string>()
  for (const path of candidates) {
    if (retained.has(path)) continue
    removed.add(path)
    zip.remove(path)
    const dependencyRelsPath = relationshipPartPath(path)
    if (zip.file(dependencyRelsPath)) {
      removed.add(dependencyRelsPath)
      zip.remove(dependencyRelsPath)
    }
  }
  if (removed.size === 0) return

  const contentTypesPath = '[Content_Types].xml'
  const contentTypesFile = zip.file(contentTypesPath)
  if (!contentTypesFile) return
  const contentTypesXml = await contentTypesFile.async('string')
  const cleanedContentTypesXml = contentTypesXml.replace(OVERRIDE_TAG_RE, (tag) => {
    const partName = xmlAttr(tag, 'PartName')
    if (!partName) return tag
    const path = normalizePartPath(decodePartUri(partName.replace(/^\//, '')))
    return path && removed.has(path) ? '' : tag
  })
  if (cleanedContentTypesXml !== contentTypesXml) {
    zip.file(contentTypesPath, cleanedContentTypesXml, { date: contentTypesFile.date })
  }
}
