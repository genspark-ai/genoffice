/// Surgical edits to visuals that already live in the file: remove an
/// anchor, or move/resize one by rewriting its from/to markers. Visuals are
/// located by the sidecar's (drawingPath, drawingIndex) pair — the index
/// counts every anchor element in document order, matching visuals.rs.

import type { WorkbookVisualEdit } from '../shared/desktop-api'
import { relsPathFor, resolveRelTarget, type MutablePackage } from './xlsx-drawing-add'
import {
  parseRelationships,
  partPathForRels,
  removePartOverride,
  removeRelationshipById,
  type ParsedRelationship,
} from './xlsx-sheets'

export class VisualEditError extends Error {}

const ANCHOR_PATTERN = /<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g
const ATTRIBUTE_PATTERN = /\s[\w:.-]+="([^"]*)"/g
const CHART_OWNED_RELATIONSHIP_TYPES =
  /\/(?:chartStyle|chartColorStyle|package|oleObject|theme|themeOverride|chartUserShapes|image|externalLink)$/
const CHART_OWNED_PATHS =
  /^xl\/(?:charts|media|embeddings|theme|drawings|externalLinks|oleObjects)\//

type RemovedAnchorRelationship = {
  readonly id: string
  readonly kind: 'chart' | 'image'
}

type OwnedPartRoot = {
  readonly path: string
  readonly recursive: boolean
}

export async function applyVisualEdits(
  pkg: MutablePackage,
  edits: readonly WorkbookVisualEdit[],
  touchedEntries: Set<string>,
): Promise<void> {
  const byPath = new Map<string, WorkbookVisualEdit[]>()
  for (const edit of edits) {
    const group = byPath.get(edit.drawingPath) ?? []
    group.push(edit)
    byPath.set(edit.drawingPath, group)
  }
  for (const [drawingPath, group] of byPath) {
    if (!(await pkg.has(drawingPath))) {
      throw new VisualEditError(`Workbook is missing ${drawingPath}.`)
    }
    let xml = await pkg.readText(drawingPath)
    // Removals splice text out, so process high indexes first — lower
    // indexes keep their document positions.
    const ordered = [...group].sort((left, right) => right.drawingIndex - left.drawingIndex)
    const seen = new Set<number>()
    const removedRelationships: RemovedAnchorRelationship[] = []
    for (const edit of ordered) {
      if (seen.has(edit.drawingIndex)) {
        throw new VisualEditError('Duplicate edits target the same drawing anchor.')
      }
      seen.add(edit.drawingIndex)
      xml = applyOneEdit(xml, edit, removedRelationships)
    }
    if (removedRelationships.length > 0) {
      await cleanupRemovedAnchorRelationships(
        pkg,
        drawingPath,
        xml,
        removedRelationships,
        touchedEntries,
      )
    }
    pkg.write(drawingPath, xml)
    touchedEntries.add(drawingPath)
    await cleanupEmptyDrawingHookup(pkg, drawingPath, xml, touchedEntries)
  }
}

/**
 * Drop relationship rows no longer referenced by the drawing, then collect
 * their unshared package targets. Chart targets recursively own style, color,
 * embedded-workbook, theme-override, user-shape, and image relationships.
 */
async function cleanupRemovedAnchorRelationships(
  pkg: MutablePackage,
  drawingPath: string,
  remainingDrawingXml: string,
  removedRelationships: readonly RemovedAnchorRelationship[],
  touchedEntries: Set<string>,
): Promise<void> {
  const relsPath = relsPathFor(drawingPath)
  if (!(await pkg.has(relsPath))) {
    throw new VisualEditError('The drawing is missing its relationships part.')
  }
  let relsXml = await pkg.readText(relsPath)
  const roots: OwnedPartRoot[] = []
  const kindsById = new Map<string, RemovedAnchorRelationship['kind']>()
  for (const removed of removedRelationships) {
    const previous = kindsById.get(removed.id)
    if (previous !== undefined && previous !== removed.kind) {
      throw new VisualEditError('One drawing relationship is used by incompatible visuals.')
    }
    kindsById.set(removed.id, removed.kind)
  }
  for (const [relId, kind] of kindsById) {
    // A second anchor can intentionally reuse the same relationship. Removing
    // one anchor must not break the survivor.
    if (xmlHasAttributeValue(remainingDrawingXml, relId)) continue
    const relationship = parseRelationships(relsXml).find((entry) => entry.id === relId)
    if (!relationship) {
      throw new VisualEditError(`The deleted ${kind} has no drawing relationship.`)
    }
    const expectedType = kind === 'chart' ? /\/chart$/ : /\/image$/
    if (!expectedType.test(relationship.type)) {
      throw new VisualEditError(
        `The deleted ${kind} uses an unsupported drawing relationship (${relationship.type}).`,
      )
    }
    const withoutRelationship = removeRelationshipById(relsXml, relId)
    if (withoutRelationship === relsXml) {
      throw new VisualEditError(`The deleted ${kind} relationship could not be removed safely.`)
    }
    relsXml = withoutRelationship
    if (relationship.external) continue
    const targetPath = resolvePackageTarget(drawingPath, relationship.target)
    if (kind === 'chart') {
      if (!/^xl\/charts\/[^/]+\.xml$/.test(targetPath) || !(await pkg.has(targetPath))) {
        throw new VisualEditError('The deleted chart relationship has an invalid package target.')
      }
      roots.push({ path: targetPath, recursive: true })
    } else if (/^xl\/media\//.test(targetPath) && (await pkg.has(targetPath))) {
      roots.push({ path: targetPath, recursive: false })
    }
  }
  const removedParts = await collectUnreferencedOwnedParts(
    pkg,
    roots,
    new Map([[relsPath, relsXml]]),
  )
  const contentTypesPath = '[Content_Types].xml'
  let contentTypes = removedParts.size === 0 ? '' : await pkg.readText(contentTypesPath)
  const originalContentTypes = contentTypes
  for (const path of removedParts) contentTypes = removePartOverride(contentTypes, path)

  // All validation and relationship-graph reads above complete before the
  // package overlay is mutated, so unsupported chart dependencies fail closed.
  pkg.write(relsPath, relsXml)
  touchedEntries.add(relsPath)
  for (const path of removedParts) pkg.remove(path)
  if (contentTypes !== originalContentTypes) {
    pkg.write(contentTypesPath, contentTypes)
    touchedEntries.add(contentTypesPath)
  }
}

/**
 * Mirrors sheet-delete's relationship closure/retention pass: collect owned
 * descendants, then pull back any target reached from a surviving package part.
 */
async function collectUnreferencedOwnedParts(
  pkg: MutablePackage,
  roots: readonly OwnedPartRoot[],
  relationshipOverrides: ReadonlyMap<string, string> = new Map(),
): Promise<Set<string>> {
  const closure = new Set<string>()
  const queue = [...roots]
  while (queue.length > 0) {
    const current = queue.pop() as OwnedPartRoot
    if (closure.has(current.path) || !(await pkg.has(current.path))) continue
    closure.add(current.path)
    const childRelsPath = relsPathFor(current.path)
    if (!(await pkg.has(childRelsPath))) continue
    closure.add(childRelsPath)
    if (!current.recursive) continue
    const childRelationships =
      relationshipOverrides.get(childRelsPath) ?? (await pkg.readText(childRelsPath))
    for (const relationship of parseRelationships(childRelationships)) {
      if (relationship.external) continue
      const targetPath = resolvePackageTarget(current.path, relationship.target)
      if (
        !CHART_OWNED_RELATIONSHIP_TYPES.test(relationship.type) ||
        !CHART_OWNED_PATHS.test(targetPath)
      ) {
        throw new VisualEditError(
          `The chart carries an unsupported relationship (${relationship.type}) — deletion aborted.`,
        )
      }
      queue.push({ path: targetPath, recursive: true })
    }
  }

  const retained = new Set<string>()
  const packagePaths = await pkg.paths()
  for (const relsPath of packagePaths) {
    if (!relsPath.endsWith('.rels') || closure.has(relsPath)) continue
    const owner = partPathForRels(relsPath)
    const relationships = relationshipOverrides.get(relsPath) ?? (await pkg.readText(relsPath))
    for (const relationship of parseRelationships(relationships)) {
      if (relationship.external) continue
      const target = resolvePackageTarget(owner, relationship.target)
      if (closure.has(target)) retained.add(target)
    }
  }
  const propagate = [...retained]
  while (propagate.length > 0) {
    const source = propagate.pop() as string
    const childRelsPath = relsPathFor(source)
    if (!closure.has(childRelsPath) || retained.has(childRelsPath)) continue
    retained.add(childRelsPath)
    const childRelationships =
      relationshipOverrides.get(childRelsPath) ?? (await pkg.readText(childRelsPath))
    for (const relationship of parseRelationships(childRelationships)) {
      if (relationship.external) continue
      const target = resolvePackageTarget(source, relationship.target)
      if (closure.has(target) && !retained.has(target)) {
        retained.add(target)
        propagate.push(target)
      }
    }
  }
  return new Set([...closure].filter((path) => !retained.has(path)))
}

/**
 * Once the final anchor is gone, remove the worksheet hookup and empty drawing
 * only when the relationship graph is unambiguous and carries no unsupported
 * relationship rows. Otherwise the empty drawing is preserved.
 */
async function cleanupEmptyDrawingHookup(
  pkg: MutablePackage,
  drawingPath: string,
  drawingXml: string,
  touchedEntries: Set<string>,
): Promise<void> {
  const inner = /<xdr:wsDr\b[^>]*>([\s\S]*)<\/xdr:wsDr>/.exec(drawingXml)?.[1]
  if (inner === undefined || inner.trim() !== '') return
  const drawingRelsPath = relsPathFor(drawingPath)
  if (
    (await pkg.has(drawingRelsPath)) &&
    /<Relationship\b/.test(await pkg.readText(drawingRelsPath))
  ) {
    return
  }

  const incoming: {
    owner: string
    relsPath: string
    relationship: ParsedRelationship
  }[] = []
  for (const candidateRelsPath of await pkg.paths()) {
    if (!candidateRelsPath.endsWith('.rels') || candidateRelsPath === drawingRelsPath) continue
    const owner = partPathForRels(candidateRelsPath)
    for (const relationship of parseRelationships(await pkg.readText(candidateRelsPath))) {
      if (
        !relationship.external &&
        resolvePackageTarget(owner, relationship.target) === drawingPath
      ) {
        incoming.push({ owner, relsPath: candidateRelsPath, relationship })
      }
    }
  }
  if (incoming.length !== 1) return
  const incomingReference = incoming[0]
  if (incomingReference === undefined) return
  const { owner, relsPath, relationship } = incomingReference
  if (
    !/^xl\/worksheets\/[^/]+\.xml$/.test(owner) ||
    !/\/drawing$/.test(relationship.type) ||
    relationship.id === undefined ||
    !(await pkg.has(owner))
  ) {
    return
  }
  const worksheetXml = await pkg.readText(owner)
  const hookups = [...worksheetXml.matchAll(/<(?:[\w.-]+:)?drawing\b[^>]*\/>/g)].filter((match) =>
    xmlHasAttributeValue(match[0], relationship.id as string),
  )
  if (hookups.length !== 1 || hookups[0]?.index === undefined) return
  const hookup = hookups[0]
  const withoutHookup =
    worksheetXml.slice(0, hookup.index) + worksheetXml.slice(hookup.index + hookup[0].length)
  if (xmlHasAttributeValue(withoutHookup, relationship.id)) return
  const worksheetRels = await pkg.readText(relsPath)
  const withoutRelationship = removeRelationshipById(worksheetRels, relationship.id)
  if (withoutRelationship === worksheetRels) return
  const contentTypesPath = '[Content_Types].xml'
  const contentTypes = await pkg.readText(contentTypesPath)
  const stripped = removePartOverride(contentTypes, drawingPath)

  pkg.write(owner, withoutHookup)
  touchedEntries.add(owner)
  if (/<Relationship\b/.test(withoutRelationship)) {
    pkg.write(relsPath, withoutRelationship)
    touchedEntries.add(relsPath)
  } else {
    pkg.remove(relsPath)
  }
  pkg.remove(drawingPath)
  if (await pkg.has(drawingRelsPath)) pkg.remove(drawingRelsPath)
  if (stripped !== contentTypes) {
    pkg.write(contentTypesPath, stripped)
    touchedEntries.add(contentTypesPath)
  }
}

function resolvePackageTarget(fromPart: string, target: string): string {
  return target.startsWith('/') ? target.slice(1) : resolveRelTarget(fromPart, target)
}

function xmlHasAttributeValue(xml: string, value: string): boolean {
  for (const match of xml.matchAll(ATTRIBUTE_PATTERN)) {
    if (match[1] === value) return true
  }
  return false
}

function applyOneEdit(
  xml: string,
  edit: WorkbookVisualEdit,
  removedRelationships: RemovedAnchorRelationship[],
): string {
  const anchors = [...xml.matchAll(ANCHOR_PATTERN)]
  const match = anchors[edit.drawingIndex]
  if (!match) {
    throw new VisualEditError(
      `Drawing anchor #${edit.drawingIndex} was not found — the file may have changed.`,
    )
  }
  const anchorXml = match[0]
  const kind = match[1]
  if (edit.remove) {
    // A chart's graphicFrame anchor cascades: its rel, part, and override
    // are collected here and removed after the drawing XML is final.
    if (anchorXml.includes('<xdr:graphicFrame')) {
      const relId = /<c:chart\b[^>]*\br:id="([^"]+)"/.exec(anchorXml)?.[1]
      if (!relId) {
        throw new VisualEditError(
          'This graphic frame is not a chart — deleting it is not supported.',
        )
      }
      removedRelationships.push({ id: relId, kind: 'chart' })
    } else if (anchorXml.includes('<xdr:pic')) {
      for (const match of anchorXml.matchAll(/\s[\w.-]+:(?:embed|link)="([^"]+)"/g)) {
        const id = match[1]
        if (id !== undefined) removedRelationships.push({ id, kind: 'image' })
      }
    }
    return xml.slice(0, match.index) + xml.slice(match.index + anchorXml.length)
  }
  const anchor = edit.anchor
  if (!anchor) throw new VisualEditError('A visual edit needs a removal or a new anchor.')
  if (kind === 'absoluteAnchor') {
    throw new VisualEditError('This visual uses an absolute anchor — moving it is not supported.')
  }
  const from =
    `<xdr:from><xdr:col>${anchor.fromColumn}</xdr:col>` +
    `<xdr:colOff>${anchor.fromColumnOffset}</xdr:colOff>` +
    `<xdr:row>${anchor.fromRow}</xdr:row>` +
    `<xdr:rowOff>${anchor.fromRowOffset}</xdr:rowOff></xdr:from>`
  let patched = anchorXml.replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, () => from)
  if (patched === anchorXml && !anchorXml.includes('<xdr:from>')) {
    throw new VisualEditError('Drawing anchor has no from marker — moving it is not supported.')
  }
  if (kind === 'twoCellAnchor') {
    const to =
      `<xdr:to><xdr:col>${anchor.toColumn}</xdr:col>` +
      `<xdr:colOff>${anchor.toColumnOffset}</xdr:colOff>` +
      `<xdr:row>${anchor.toRow}</xdr:row>` +
      `<xdr:rowOff>${anchor.toRowOffset}</xdr:rowOff></xdr:to>`
    // An unchanged edge replaces to an identical string, so presence must be
    // checked directly (an NW resize touches only the from marker).
    const withTo = patched.replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, () => to)
    if (withTo === patched && !patched.includes('<xdr:to>')) {
      throw new VisualEditError('Drawing anchor has no to marker — moving it is not supported.')
    }
    patched = withTo
  }
  if (edit.frameSize) {
    // A rotated shape resized through its AABB: the anchor holds the rotated
    // bounds, so the true frame (first a:ext of the shape's xfrm) must be
    // rewritten alongside or the reload re-derives the old size.
    const ext = `<a:ext cx="${edit.frameSize.width}" cy="${edit.frameSize.height}"/>`
    const withExt = patched.replace(
      /(<a:xfrm\b[^>]*>[\s\S]*?)<a:ext\b[^>]*\/>/,
      (_, head: string) => `${head}${ext}`,
    )
    // An unchanged ext replaces to an identical string, so presence must be
    // checked directly (same caveat as the from/to markers above).
    if (withExt === patched && !/<a:xfrm\b[^>]*>[\s\S]*?<a:ext\b/.test(patched)) {
      throw new VisualEditError(
        'Drawing anchor has no frame extent — resizing it is not supported.',
      )
    }
    patched = withExt
  }
  return xml.slice(0, match.index) + patched + xml.slice(match.index + anchorXml.length)
}
