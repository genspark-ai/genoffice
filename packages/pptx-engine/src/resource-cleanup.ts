/**
 * Relationship-aware cleanup for package parts dropped by slide mutations.
 *
 * A media/chart part cannot be removed merely because one shape stopped using it:
 * duplicated slides and copied objects may point at the same target. Cleanup first
 * removes only relationships no longer used by the owning XML, then collects the
 * now-unreachable resource subgraph.
 */
import type { Slide } from './types'
import { PackageArchive, relsPathFor, resolveTarget } from './zip'

export interface ResourceCleanupStats {
  relationshipsRemoved: number
  partsRemoved: number
  bytesRemoved: number
}

interface ResourceCleanupContext {
  archive: PackageArchive
  deck: { slides: Slide[] }
}

const EMPTY_STATS: ResourceCleanupStats = {
  relationshipsRemoved: 0,
  partsRemoved: 0,
  bytesRemoved: 0,
}

const RELATIONSHIP_TAG = /<Relationship\b[^>]*\/\s*>/g
const ATTRIBUTE = /\s[\w:.-]+\s*=\s*(["'])(.*?)\1/g
const ID_ATTRIBUTE = /\bId\s*=\s*(["'])(.*?)\1/
const PART_NAME_ATTRIBUTE = /\bPartName\s*=\s*(["'])(.*?)\1/
const OVERRIDE_TAG = /<Override\b[^>]*\/\s*>/g
const MODEL_3D_REF = /aislides-3d:(ppt\/media\/[^"'&<\s]+)/g

/**
 * Parts below these directories are owned by a slide object/slide and can be
 * collected once no retained package part points at them.
 */
const OWNED_RESOURCE_PREFIXES = [
  'ppt/media/',
  'ppt/charts/',
  'ppt/embeddings/',
  'ppt/diagrams/',
  'ppt/drawings/',
  'ppt/oleObjects/',
  'ppt/activeX/',
  'ppt/ctrlProps/',
  'ppt/controls/',
  'ppt/tags/',
  'ppt/notesSlides/',
  'ppt/comments/',
  'ppt/threadedComments/',
  'ppt/printerSettings/',
  'ppt/vmlDrawings/',
  'ppt/ink/',
  'ppt/model3d/',
  'ppt/models3d/',
] as const

function isOwnedResourcePath(path: string): boolean {
  return OWNED_RESOURCE_PREFIXES.some((prefix) => path.startsWith(prefix))
}

function attributeValues(xml: string): Set<string> {
  const out = new Set<string>()
  for (const match of xml.matchAll(ATTRIBUTE)) out.add(match[2]!)
  return out
}

function model3dRefs(xml: string): Set<string> {
  const out = new Set<string>()
  for (const match of xml.matchAll(MODEL_3D_REF)) out.add(match[1]!)
  return out
}

function currentSlideXml(slide: Slide): string {
  const parts = [slide.bodyPrefix]
  for (const element of slide.elements) {
    parts.push(element.anchor.originalXml)
    if (element.anchor.gapAfter) parts.push(element.anchor.gapAfter)
  }
  parts.push(slide.bodySuffix)
  return parts.join('')
}

function sourcePartForRels(path: string): string | null {
  if (path === '_rels/.rels') return ''
  const marker = '/_rels/'
  const markerAt = path.lastIndexOf(marker)
  if (markerAt < 0 || !path.endsWith('.rels')) return null
  const dir = path.slice(0, markerAt)
  const name = path.slice(markerAt + marker.length, -'.rels'.length)
  return `${dir}/${name}`
}

function stripRelationships(xml: string, ids: ReadonlySet<string>): string {
  return xml.replace(RELATIONSHIP_TAG, (tag) => {
    const id = ID_ATTRIBUTE.exec(tag)?.[2]
    return id && ids.has(id) ? '' : tag
  })
}

function removeContentTypeOverrides(
  archive: PackageArchive,
  removedParts: ReadonlySet<string>,
): void {
  if (removedParts.size === 0) return
  const path = '[Content_Types].xml'
  const xml = archive.readText(path)
  if (!xml) return
  const next = xml.replace(OVERRIDE_TAG, (tag) => {
    const partName = PART_NAME_ATTRIBUTE.exec(tag)?.[2]?.replace(/^\//, '')
    return partName && removedParts.has(partName) ? '' : tag
  })
  if (next !== xml) archive.entries.set(path, Buffer.from(next, 'utf8'))
}

function packageRelationshipGraph(archive: PackageArchive): {
  outgoing: Map<string, Set<string>>
  incoming: Map<string, Set<string>>
} {
  const outgoing = new Map<string, Set<string>>()
  const incoming = new Map<string, Set<string>>()
  for (const path of archive.entries.keys()) {
    if (!path.endsWith('.rels')) continue
    const source = sourcePartForRels(path)
    if (source == null || (source !== '' && !archive.entries.has(source))) continue
    for (const rel of archive.readRels(source).values()) {
      if (rel.targetMode === 'External') continue
      const target = resolveTarget(source, rel.target)
      let targets = outgoing.get(source)
      if (!targets) outgoing.set(source, (targets = new Set()))
      targets.add(target)
      let sources = incoming.get(target)
      if (!sources) incoming.set(target, (sources = new Set()))
      sources.add(source)
    }
  }
  return { outgoing, incoming }
}

/**
 * Remove the candidate resource closure only when it has no incoming edge from a
 * retained package part. If a shared chart/image is retained, its dependencies
 * are retained too.
 */
function removeUnreferencedOwnedParts(
  archive: PackageArchive,
  initialCandidates: Iterable<string>,
  xmlOverrides: ReadonlyMap<string, string> = new Map(),
): ResourceCleanupStats {
  const { outgoing, incoming } = packageRelationshipGraph(archive)
  const closure = new Set<string>()
  const pending = [...initialCandidates]
  while (pending.length) {
    const path = pending.pop()!
    if (closure.has(path) || !isOwnedResourcePath(path) || !archive.entries.has(path)) continue
    closure.add(path)
    for (const target of outgoing.get(path) ?? []) {
      if (isOwnedResourcePath(target)) pending.push(target)
    }
  }
  if (closure.size === 0) return { ...EMPTY_STATS }

  const retained = new Set<string>()
  for (const path of closure) {
    for (const source of incoming.get(path) ?? []) {
      if (!closure.has(source)) {
        retained.add(path)
        break
      }
    }
  }

  // GenOffice's simplified 3D placeholder records the model path in cNvPr@descr
  // rather than an OOXML relationship, so include those direct references.
  for (const [path, bytes] of archive.entries) {
    if (!path.endsWith('.xml') || closure.has(path)) continue
    const xml = xmlOverrides.get(path) ?? Buffer.from(bytes).toString('utf8')
    for (const target of model3dRefs(xml)) {
      if (closure.has(target)) retained.add(target)
    }
  }

  const propagate = [...retained]
  while (propagate.length) {
    const source = propagate.pop()!
    for (const target of outgoing.get(source) ?? []) {
      if (closure.has(target) && !retained.has(target)) {
        retained.add(target)
        propagate.push(target)
      }
    }
  }

  const removed = new Set<string>()
  let bytesRemoved = 0
  for (const path of closure) {
    if (retained.has(path)) continue
    const bytes = archive.entries.get(path)
    if (!bytes) continue
    bytesRemoved += bytes.byteLength
    removed.add(path)
    archive.entries.delete(path)
    archive.entries.delete(relsPathFor(path))
  }
  removeContentTypeOverrides(archive, removed)
  return {
    relationshipsRemoved: 0,
    partsRemoved: removed.size,
    bytesRemoved,
  }
}

/**
 * Remove selected relationship rows and collect their now-unreferenced owned
 * targets. The package graph is rebuilt after the rows are removed, so another
 * relationship to the same target keeps that target and its dependency closure.
 */
function removeRelationshipsAndCollectOwnedTargets(
  archive: PackageArchive,
  sourcePart: string,
  relationshipIds: Iterable<string>,
  extraCandidates: Iterable<string> = [],
  xmlOverrides: ReadonlyMap<string, string> = new Map(),
): ResourceCleanupStats {
  const rels = archive.readRels(sourcePart)
  const ids = new Set<string>()
  const candidates = new Set(extraCandidates)
  for (const id of relationshipIds) {
    const rel = rels.get(id)
    if (!rel) continue
    ids.add(id)
    if (rel.targetMode !== 'External') {
      const target = resolveTarget(sourcePart, rel.target)
      if (isOwnedResourcePath(target)) candidates.add(target)
    }
  }

  if (ids.size) {
    const path = relsPathFor(sourcePart)
    const xml = archive.readText(path)
    if (xml) archive.entries.set(path, Buffer.from(stripRelationships(xml, ids), 'utf8'))
  }

  const collected = removeUnreferencedOwnedParts(archive, candidates, xmlOverrides)
  return {
    relationshipsRemoved: ids.size,
    partsRemoved: collected.partsRemoved,
    bytesRemoved: collected.bytesRemoved,
  }
}

/**
 * Compare XML from before and after a slide mutation. Relationship ids that
 * disappeared from the live slide are pruned, then their unshared owned targets
 * are collected. Callers pass the fully patched current slide XML so pending
 * in-memory edits participate in the liveness check.
 */
export function cleanupSupersededSlideResources(
  opened: ResourceCleanupContext,
  slide: Slide,
  previousXml: string,
  currentXml: string,
): ResourceCleanupStats {
  const { archive } = opened
  const rels = archive.readRels(slide.path)
  const previousValues = attributeValues(previousXml)
  const currentValues = attributeValues(currentXml)
  const liveSlideXml = new Map(
    opened.deck.slides.map((deckSlide) => [deckSlide.path, currentSlideXml(deckSlide)]),
  )
  liveSlideXml.set(slide.path, currentXml)
  const ids = new Set<string>()
  const candidates = new Set<string>()
  for (const id of rels.keys()) {
    if (!previousValues.has(id) || currentValues.has(id)) continue
    ids.add(id)
  }

  const currentModels = model3dRefs(currentXml)
  for (const target of model3dRefs(previousXml)) {
    if (!currentModels.has(target)) candidates.add(target)
  }

  return removeRelationshipsAndCollectOwnedTargets(
    archive,
    slide.path,
    ids,
    candidates,
    liveSlideXml,
  )
}

/**
 * Called after an element has left slide.elements. It removes relationship rows
 * used only by that element, then collects unshared media/chart dependency parts.
 */
export function cleanupDeletedElementResources(
  opened: ResourceCleanupContext,
  slide: Slide,
  removedElementXml: string,
): ResourceCleanupStats {
  return cleanupSupersededSlideResources(opened, slide, removedElementXml, currentSlideXml(slide))
}

/**
 * Remove one explicit relationship and collect its internal owned target only
 * when no retained package relationship still points at it. This covers
 * relationship-owned parts (such as a slide's comments part) that are not
 * referenced from the source XML.
 */
export function removeRelationshipAndCollectOwnedTarget(
  archive: PackageArchive,
  sourcePart: string,
  relationshipId: string,
): ResourceCleanupStats {
  return removeRelationshipsAndCollectOwnedTargets(archive, sourcePart, [relationshipId])
}

/**
 * Delete a package part and collect resource parts that were owned exclusively by
 * it. Used by slide deletion after the presentation relationship is removed.
 */
export function removePartWithOwnedResources(
  archive: PackageArchive,
  partPath: string,
): ResourceCleanupStats {
  const candidates = new Set<string>()
  for (const rel of archive.readRels(partPath).values()) {
    if (rel.targetMode === 'External') continue
    const target = resolveTarget(partPath, rel.target)
    if (isOwnedResourcePath(target)) candidates.add(target)
  }
  const xml = archive.readText(partPath)
  if (xml) {
    for (const target of model3dRefs(xml)) candidates.add(target)
  }

  let bytesRemoved = archive.entries.get(partPath)?.byteLength ?? 0
  archive.entries.delete(partPath)
  archive.entries.delete(relsPathFor(partPath))
  removeContentTypeOverrides(archive, new Set([partPath]))

  const collected = removeUnreferencedOwnedParts(archive, candidates)
  bytesRemoved += collected.bytesRemoved
  return {
    relationshipsRemoved: 0,
    partsRemoved: 1 + collected.partsRemoved,
    bytesRemoved,
  }
}
