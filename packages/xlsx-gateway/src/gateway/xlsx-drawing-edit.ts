/// Surgical edits to visuals that already live in the file: remove an
/// anchor, or move/resize one by rewriting its from/to markers. Visuals are
/// located by the sidecar's (drawingPath, drawingIndex) pair — the index
/// counts every anchor element in document order, matching visuals.rs.

import type { WorkbookVisualEdit } from '../shared/edit-schemas'
import {
  appendRelationship,
  outlineXml,
  relsPathFor,
  resolveRelTarget,
  solidFillXml,
  type MutablePackage,
} from './xlsx-drawing-add'
import {
  parseRelationships,
  partPathForRels,
  removePartOverride,
  removeRelationshipById,
  type ParsedRelationship,
} from './xlsx-sheets'

export class VisualEditError extends Error {}

// The spreadsheetDrawing namespace is conventionally bound to `xdr:`, but
// openpyxl writes it as the DEFAULT namespace (no prefix) — and the sidecar
// counts anchors by local name, so the index pairing must see those too.
const ANCHOR_PATTERN =
  /<([A-Za-z_][\w.-]*:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/\1\2>/g
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

const HYPERLINK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

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
    const xml = await pkg.readText(drawingPath)
    const anchors = [...xml.matchAll(ANCHOR_PATTERN)]
    const byIndex = new Map<number, WorkbookVisualEdit>()
    for (const edit of group) {
      if (byIndex.has(edit.drawingIndex)) {
        throw new VisualEditError('Duplicate edits target the same drawing anchor.')
      }
      if (!anchors[edit.drawingIndex]) {
        throw new VisualEditError(
          `Drawing anchor #${edit.drawingIndex} was not found — the file may have changed.`,
        )
      }
      byIndex.set(edit.drawingIndex, edit)
    }
    const relsPath = relsPathFor(drawingPath)
    const hyperlinkRelIds = new Map<number, string>()
    for (const edit of group) {
      if (!edit.hyperlink || edit.remove) continue
      hyperlinkRelIds.set(
        edit.drawingIndex,
        await appendRelationship(pkg, relsPath, HYPERLINK_REL_TYPE, edit.hyperlink, 'External'),
      )
      touchedEntries.add(relsPath)
    }
    const removedRelationships: RemovedAnchorRelationship[] = []
    const droppedHyperlinkRelIds: string[] = []
    const patched = anchors.map((match, index) => {
      const edit = byIndex.get(index)
      return edit
        ? applyOneEdit(
            match,
            edit,
            removedRelationships,
            hyperlinkRelIds.get(index),
            droppedHyperlinkRelIds,
          )
        : match[0]
    })
    // Document order is z-order. Anchors carrying a zIndex permute among
    // their own slots (sorted by zIndex); every other anchor — and all the
    // markup between anchors, e.g. mc:AlternateContent wrappers — stays put.
    const movable = anchors
      .map((match, index) => ({ match, index }))
      .filter(
        ({ match, index }) =>
          byIndex.get(index)?.zIndex !== undefined &&
          patched[index] !== null &&
          !insideAlternateContent(xml, match.index),
      )
    const sorted = [...movable].sort(
      (left, right) =>
        byIndex.get(left.index)!.zIndex! - byIndex.get(right.index)!.zIndex! ||
        left.index - right.index,
    )
    const placed = [...patched]
    movable.forEach(({ index }, slot) => {
      placed[index] = patched[sorted[slot]!.index]!
    })
    let next = xml
    for (let index = anchors.length - 1; index >= 0; index -= 1) {
      const match = anchors[index]!
      const text = placed[index]
      if (text === match[0]) continue
      next = next.slice(0, match.index) + (text ?? '') + next.slice(match.index + match[0].length)
    }
    if (removedRelationships.length > 0) {
      await cleanupRemovedAnchorRelationships(
        pkg,
        drawingPath,
        next,
        removedRelationships,
        touchedEntries,
      )
    }
    const orphaned = droppedHyperlinkRelIds.filter((id) => !xmlHasAttributeValue(next, id))
    if (orphaned.length > 0 && (await pkg.has(relsPath))) {
      let relsXml = await pkg.readText(relsPath)
      for (const id of orphaned) relsXml = removeRelationshipById(relsXml, id)
      pkg.write(relsPath, relsXml)
      touchedEntries.add(relsPath)
    }
    pkg.write(drawingPath, next)
    touchedEntries.add(drawingPath)
    await cleanupEmptyDrawingHookup(pkg, drawingPath, next, touchedEntries)
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
    const targetPath = resolveRelTarget(drawingPath, relationship.target)
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
      const targetPath = resolveRelTarget(current.path, relationship.target)
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
      const target = resolveRelTarget(owner, relationship.target)
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
      const target = resolveRelTarget(source, relationship.target)
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
  const inner = /<(?:[A-Za-z_][\w.-]*:)?wsDr\b[^>]*>([\s\S]*)<\/(?:[A-Za-z_][\w.-]*:)?wsDr>/.exec(
    drawingXml,
  )?.[1]
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
      if (!relationship.external && resolveRelTarget(owner, relationship.target) === drawingPath) {
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

/// True when the position sits inside an open mc:AlternateContent (slicer
/// and timeline anchors live in its Choice/Fallback branches).
function insideAlternateContent(xml: string, position: number): boolean {
  const prefix = xml.slice(0, position)
  const opens = prefix.match(/<mc:AlternateContent\b/g)?.length ?? 0
  const closes = prefix.match(/<\/mc:AlternateContent>/g)?.length ?? 0
  return opens > closes
}

function xmlHasAttributeValue(xml: string, value: string): boolean {
  for (const match of xml.matchAll(ATTRIBUTE_PATTERN)) {
    if (match[1] === value) return true
  }
  return false
}

const escapeRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

const escapeXmlAttribute = (value: string): string =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/// Sets (or with null removes) one attribute on an opening tag string.
export function setTagAttribute(tag: string, name: string, value: string | null): string {
  const pattern = new RegExp(`\\s${escapeRegExp(name)}="[^"]*"`)
  const stripped = tag.replace(pattern, '')
  if (value === null) return stripped
  const insertAt = stripped.endsWith('/>') ? stripped.length - 2 : stripped.length - 1
  return `${stripped.slice(0, insertAt)} ${name}="${escapeXmlAttribute(value)}"${stripped.slice(insertAt)}`
}

/// Degrees clockwise → a:xfrm/@rot (60000ths of a degree in [0, 360°)).
export function rotationAttribute(degrees: number): string | null {
  const normalized = ((Math.round(degrees * 60_000) % 21_600_000) + 21_600_000) % 21_600_000
  return normalized === 0 ? null : String(normalized)
}

const XFRM_OPEN = /<a:xfrm\b[^>]*>/
const CNVPR_PATTERN = /<([A-Za-z_][\w.-]*:)?cNvPr\b[^>]*?(\/>|>[\s\S]*?<\/\1cNvPr>)/

/// One anchor's edits applied to its XML text; null when the anchor is
/// removed. Hyperlink relationships are allocated by the caller.
function applyOneEdit(
  match: RegExpMatchArray,
  edit: WorkbookVisualEdit,
  removedRelationships: RemovedAnchorRelationship[],
  hyperlinkRelId: string | undefined,
  droppedHyperlinkRelIds: string[],
): string | null {
  const anchorXml = match[0]
  const p = match[1] ?? ''
  const kind = match[2]
  if (edit.remove) {
    // A chart's graphicFrame anchor cascades: its rel, part, and override
    // are collected here and removed after the drawing XML is final.
    if (anchorXml.includes(`<${p}graphicFrame`)) {
      const relId = /<(?:[A-Za-z_][\w.-]*:)?chart\b[^>]*\br:id="([^"]+)"/.exec(anchorXml)?.[1]
      if (!relId) {
        throw new VisualEditError(
          'This graphic frame is not a chart — deleting it is not supported.',
        )
      }
      removedRelationships.push({ id: relId, kind: 'chart' })
    } else if (anchorXml.includes(`<${p}pic`)) {
      for (const match of anchorXml.matchAll(/\s[\w.-]+:(?:embed|link)="([^"]+)"/g)) {
        const id = match[1]
        if (id !== undefined) removedRelationships.push({ id, kind: 'image' })
      }
    }
    return null
  }
  let patched = anchorXml
  const pre = escapeRegExp(p)
  // Re-applying the current paint rebuilds identical XML; that is a no-op,
  // not a failure.
  if (edit.fillColor !== undefined || edit.lineColor !== undefined) {
    patched = repaintShape(patched, p, edit.fillColor, edit.lineColor)
  }
  const anchor = edit.anchor
  if (anchor) {
    if (kind === 'absoluteAnchor') {
      throw new VisualEditError('This visual uses an absolute anchor — moving it is not supported.')
    }
    const from =
      `<${p}from><${p}col>${anchor.fromColumn}</${p}col>` +
      `<${p}colOff>${anchor.fromColumnOffset}</${p}colOff>` +
      `<${p}row>${anchor.fromRow}</${p}row>` +
      `<${p}rowOff>${anchor.fromRowOffset}</${p}rowOff></${p}from>`
    const moved = patched.replace(new RegExp(`<${pre}from>[\\s\\S]*?</${pre}from>`), () => from)
    if (moved === patched && !patched.includes(`<${p}from>`)) {
      throw new VisualEditError('Drawing anchor has no from marker — moving it is not supported.')
    }
    patched = moved
    if (kind === 'twoCellAnchor') {
      const to =
        `<${p}to><${p}col>${anchor.toColumn}</${p}col>` +
        `<${p}colOff>${anchor.toColumnOffset}</${p}colOff>` +
        `<${p}row>${anchor.toRow}</${p}row>` +
        `<${p}rowOff>${anchor.toRowOffset}</${p}rowOff></${p}to>`
      // An unchanged edge replaces to an identical string, so presence must be
      // checked directly (an NW resize touches only the from marker).
      const withTo = patched.replace(new RegExp(`<${pre}to>[\\s\\S]*?</${pre}to>`), () => to)
      if (withTo === patched && !patched.includes(`<${p}to>`)) {
        throw new VisualEditError('Drawing anchor has no to marker — moving it is not supported.')
      }
      patched = withTo
    }
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
    if (withExt === patched && !/<a:xfrm\b[^>]*>[\s\S]*?<a:ext\b/.test(patched)) {
      throw new VisualEditError(
        'Drawing anchor has no frame extent — resizing it is not supported.',
      )
    }
    patched = withExt
  }
  if (edit.rotation !== undefined || edit.flipH !== undefined || edit.flipV !== undefined) {
    const open = XFRM_OPEN.exec(patched)
    if (!open) {
      throw new VisualEditError(
        'Drawing anchor has no frame transform — rotating it is not supported.',
      )
    }
    let tag = open[0]
    if (edit.rotation !== undefined)
      tag = setTagAttribute(tag, 'rot', rotationAttribute(edit.rotation))
    if (edit.flipH !== undefined) tag = setTagAttribute(tag, 'flipH', edit.flipH ? '1' : null)
    if (edit.flipV !== undefined) tag = setTagAttribute(tag, 'flipV', edit.flipV ? '1' : null)
    patched = patched.slice(0, open.index) + tag + patched.slice(open.index + open[0].length)
  }
  if (edit.altText !== undefined || edit.hyperlink !== undefined) {
    const element = CNVPR_PATTERN.exec(patched)
    if (!element) {
      throw new VisualEditError('Drawing anchor has no non-visual properties element.')
    }
    const prefix = element[1] ?? ''
    const selfClosing = element[2] === '/>'
    const openEnd = selfClosing ? element[0].length - 2 : element[0].indexOf('>') + 1
    let open = selfClosing ? `${element[0].slice(0, openEnd)}>` : element[0].slice(0, openEnd)
    let inner = selfClosing ? '' : element[0].slice(openEnd, element[0].lastIndexOf('</'))
    if (edit.altText !== undefined) {
      open = setTagAttribute(open, 'descr', edit.altText === '' ? null : edit.altText)
    }
    if (edit.hyperlink !== undefined) {
      for (const link of inner.matchAll(/<a:hlinkClick\b[^>]*?\br:id="([^"]+)"/g)) {
        if (link[1] !== undefined) droppedHyperlinkRelIds.push(link[1])
      }
      inner = inner.replace(/<a:hlinkClick\b[^>]*?(?:\/>|>[\s\S]*?<\/a:hlinkClick>)/g, '')
      if (hyperlinkRelId !== undefined) {
        inner =
          '<a:hlinkClick xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
          ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
          ` r:id="${escapeXmlAttribute(hyperlinkRelId)}"/>` +
          inner
      }
    }
    const rebuilt = inner === '' ? `${open.slice(0, -1)}/>` : `${open}${inner}</${prefix}cNvPr>`
    patched =
      patched.slice(0, element.index) + rebuilt + patched.slice(element.index + element[0].length)
  }
  if (edit.editAs !== undefined) {
    if (kind !== 'twoCellAnchor') {
      throw new VisualEditError('Only two-cell anchors carry a placement mode.')
    }
    const open = /^<[^>]*>/.exec(patched)!
    const tag = setTagAttribute(open[0], 'editAs', edit.editAs === 'twoCell' ? null : edit.editAs)
    patched = tag + patched.slice(open[0].length)
  }
  return patched
}

const SP_PR_FILL =
  /<a:(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:\1>)/g
const SP_PR_LINE = /<a:ln\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:ln>)/
const GEOMETRY_END = /<a:prstGeom\b[^>]*\/>|<\/a:prstGeom>|<\/a:custGeom>/
const XFRM_END = /<\/a:xfrm>/

/// Rewrites the shape's own spPr paint. An explicit solidFill/noFill beats
/// any xdr:style fillRef, so the style block is left alone. The outline keeps
/// its width/dash/arrowheads and only swaps the fill inside `a:ln`.
function repaintShape(
  anchorXml: string,
  p: string,
  fillColor: string | undefined,
  lineColor: string | undefined,
): string {
  const pre = escapeRegExp(p)
  // Group anchors flatten to several visuals sharing one index; the first
  // spPr would be an arbitrary child's.
  if (
    !new RegExp(`<${pre}sp[\\s>]`).test(anchorXml) ||
    new RegExp(`<${pre}grpSp[\\s>]`).test(anchorXml)
  ) {
    throw new VisualEditError('Only standalone shapes can be repainted.')
  }
  const spPr = new RegExp(`(<${pre}spPr\\b[^>]*>)([\\s\\S]*?)(</${pre}spPr>)`).exec(anchorXml)
  if (!spPr)
    throw new VisualEditError('Drawing shape has no spPr — repainting it is not supported.')
  const [whole, open, inner, close] = spPr as unknown as [string, string, string, string]
  const existingLine = SP_PR_LINE.exec(inner)?.[0] ?? ''
  const existingFill = inner.replace(SP_PR_LINE, '').match(SP_PR_FILL)?.join('') ?? ''
  const body = inner.replace(SP_PR_LINE, '').replace(SP_PR_FILL, '')
  const fill = fillColor === undefined ? existingFill : solidFillXml(fillColor)
  let line = existingLine
  if (lineColor !== undefined) {
    const parts = /^<a:ln\b([^>]*?)\/?>([\s\S]*?)(?:<\/a:ln>)?$/.exec(existingLine)
    line = parts
      ? `<a:ln${parts[1]}>${solidFillXml(lineColor)}${(parts[2] ?? '').replace(SP_PR_FILL, '')}</a:ln>`
      : outlineXml(lineColor)
  }
  // Paint follows the geometry in the spPr sequence (xfrm, geom, fill, ln).
  const geometry = GEOMETRY_END.exec(body) ?? XFRM_END.exec(body)
  const at = geometry ? geometry.index + geometry[0].length : 0
  const rebuilt = `${open}${body.slice(0, at)}${fill}${line}${body.slice(at)}${close}`
  return anchorXml.slice(0, spPr.index) + rebuilt + anchorXml.slice(spPr.index + whole.length)
}
