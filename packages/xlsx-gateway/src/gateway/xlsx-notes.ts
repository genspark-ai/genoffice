/// Legacy-comment (note) writer: replaces a worksheet's full comment set —
/// the comments part, the note shapes of the VML drawing Excel needs to show
/// them, the worksheet rels, the `<legacyDrawing>` element, and the
/// [Content_Types].xml entries. Non-note VML shapes (checkboxes, buttons)
/// survive a rewrite untouched.

import { encodeXlsxEscapes } from './xlsx-escapes'
import {
  allocatePartPath,
  appendRelationship,
  registerContentTypeDefault,
  relativeTarget,
  relsPathFor,
  resolveRelTarget,
} from './xlsx-drawing-add'
import { ensureRelationshipNamespace } from './xlsx-namespace'
import { nextFreeRelationshipId } from './xlsx-sheets'
import {
  parseLegacyComments,
  syncThreadedComments,
  THREADED_COMMENTS_REL_TYPE,
  type LegacyCommentEntry,
  type SheetThread,
} from './xlsx-threaded-comments'

export type { SheetThread, SheetThreadReply } from './xlsx-threaded-comments'

export class NoteEditError extends Error {}

/**
 * A picture painted onto a note.
 *
 * Excel shows a picture in a note by painting it onto the note's shape as a
 * picture fill, so the bytes live in `xl/media` and the note's shape reaches
 * them through the VML part's own relationships.
 */
export interface NotePicture {
  readonly mediaType: 'image/png' | 'image/jpeg' | 'image/gif'
  readonly base64: string
  /**
   * The picture's pixel size. The shape is given the picture's own
   * proportions: a fill is stretched to the shape's frame, so a shape the size
   * of the text box would squash the photograph into it.
   */
  readonly widthPx: number
  readonly heightPx: number
  /// The file name, carried into the picture's title.
  readonly name?: string | undefined
}

export interface SheetNote {
  readonly row: number
  readonly column: number
  readonly author: string
  readonly text: string
  /// Threaded comment metadata (text is the root text); `null` turns a
  /// threaded cell back into a plain note.
  readonly thread?: SheetThread | null | undefined
  /**
   * The picture this note shows.
   *
   * `undefined` says nothing about the picture: whatever the note was already
   * carrying is kept. `null` takes the picture away.
   */
  readonly picture?: NotePicture | null | undefined
}

interface MutableNotePackage {
  paths(): Promise<readonly string[]>
  has(path: string): Promise<boolean>
  readText(path: string): Promise<string>
  write(path: string, content: string): void
  add(path: string, content: string): void
  addBinary(path: string, bytes: Uint8Array): void
  remove(path: string): void
}

const COMMENTS_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
const VML_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
const CONTENT_TYPES_PATH = '[Content_Types].xml'

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function columnName(index: number): string {
  let label = ''
  for (let i = index; i >= 0; i = Math.floor(i / 26) - 1) {
    label = String.fromCharCode(65 + (i % 26)) + label
  }
  return label
}

function worksheetRelsPath(worksheetPath: string): string {
  return worksheetPath.replace(/^(xl\/worksheets\/)([^/]+)$/, '$1_rels/$2.rels')
}

function relTarget(relsXml: string, type: string): string | null {
  const pattern = new RegExp(`<Relationship\\b[^>]*Type="${type}"[^>]*/?>`)
  const found = pattern.exec(relsXml)
  if (!found) return null
  const target = / Target="([^"]*)"/.exec(found[0])
  return target?.[1] ?? null
}

async function nextFreePath(
  pkg: MutableNotePackage,
  template: (index: number) => string,
): Promise<string> {
  for (let index = 1; index < 10_000; index += 1) {
    const candidate = template(index)
    if (!(await pkg.has(candidate))) return candidate
  }
  throw new NoteEditError('No free part name for the comments part.')
}

function cellRef(note: SheetNote): string {
  return `${columnName(note.column)}${note.row + 1}`
}

function buildCommentsXml(
  notes: readonly SheetNote[],
  overrides: ReadonlyMap<string, LegacyCommentEntry>,
): string {
  const authors: string[] = []
  const authorId = (author: string): number => {
    const existing = authors.indexOf(author)
    if (existing !== -1) return existing
    authors.push(author)
    return authors.length - 1
  }
  const comments = notes
    .map((note) => {
      const ref = cellRef(note)
      const entry = overrides.get(ref) ?? note
      return (
        `<comment ref="${ref}" authorId="${authorId(entry.author)}">` +
        `<text><t xml:space="preserve">${escapeXml(encodeXlsxEscapes(entry.text))}</t></text></comment>`
      )
    })
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<authors>${authors.map((author) => `<author>${escapeXml(encodeXlsxEscapes(author))}</author>`).join('')}</authors>` +
    `<commentList>${comments}</commentList></comments>`
  )
}

const VML_HEADER =
  '<xml xmlns:v="urn:schemas-microsoft-com:vml"' +
  ' xmlns:o="urn:schemas-microsoft-com:office:office"' +
  ' xmlns:x="urn:schemas-microsoft-com:office:excel">' +
  '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
  '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"' +
  ' path="m,l,21600r21600,l21600,xe">' +
  '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>' +
  '</v:shapetype>'

const FIRST_NOTE_SHAPE_ID = 1025

function nextShapeId(vmlXml: string): number {
  let next = FIRST_NOTE_SHAPE_ID
  for (const match of vmlXml.matchAll(/\bid="_x0000_s(\d+)"/g)) {
    const id = Number(match[1])
    if (id >= next) next = id + 1
  }
  return next
}

/// A note's picture, as the shape paints it: the fill element that reaches
/// the media, and the size the shape has to be given for the picture to keep
/// its proportions. The two travel together — a fill without the size would
/// stretch the photograph into the text box's frame.
interface NoteShapePicture {
  readonly fill: string
  readonly width?: string | undefined
  readonly height?: string | undefined
}

/// Excel puts a picture into a note by painting that note's shape with a
/// picture fill — a `<v:fill type="frame">` element carrying the image
/// relationship on `o:relid` (`r:id` on some producers) — resolved through
/// the VML part's own rels. The rewrite below throws every Note shape away
/// and builds fresh ones, so without this map the media is left related but
/// painted by nothing: the picture is still in the package and no longer
/// shown by anything. The whole fill element travels with the map so the
/// relationship, the title, and the recolor/rotate flags survive together.
/// (WPS writes an empty `<v:imagedata>` on every note; it carries no
/// relationship and matches nothing here.)
function noteShapePictures(vmlXml: string): Map<string, NoteShapePicture> {
  const pictures = new Map<string, NoteShapePicture>()
  for (const shape of vmlXml.matchAll(/<v:shape\b[\s\S]*?<\/v:shape>/g)) {
    if (!shape[0].includes('ObjectType="Note"')) continue
    let tag: string | null = null
    const fill = /<v:fill\b[^>]*\/>/.exec(shape[0])?.[0]
    if (fill && fill.includes('type="frame"')) {
      // Excel writes o:relid; some producers use r:id on the same element.
      if (/\b(?:o:)?relid="/.test(fill) || /\br:id="/.test(fill)) tag = fill
    }
    if (tag === null) {
      const image = /<v:imagedata\b[^>]*\br:id="([^"]+)"[^>]*\/>/.exec(shape[0])
      if (image?.[1]) tag = image[0]
    }
    if (tag === null) continue
    const row = /<x:Row>(\d+)<\/x:Row>/.exec(shape[0])
    const column = /<x:Column>(\d+)<\/x:Column>/.exec(shape[0])
    if (!row || !column) continue
    // The shape's own size is the picture's size: a fill is stretched to the
    // frame it paints, so keeping the frame is what keeps the photograph's
    // proportions.
    const style = / style="([^"]*)"/.exec(shape[0])?.[1] ?? ''
    pictures.set(`${column[1]}:${row[1]}`, {
      fill: tag,
      width: /(?:^|;)\s*width:([^;]+)/.exec(style)?.[1],
      height: /(?:^|;)\s*height:([^;]+)/.exec(style)?.[1],
    })
  }
  return pictures
}

function noteImageKey(note: SheetNote): string {
  return `${note.column}:${note.row}`
}

const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const IMAGE_EXTENSIONS: Record<NotePicture['mediaType'], string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/gif': 'gif',
}

/// Excel's own ceiling for a note's shape. A photograph larger than that is
/// scaled down — proportionally, so the picture keeps its shape.
const MAX_NOTE_PICTURE_PT = 255

/// The frame a picture is given: its own size in points (96 px per inch,
/// 72 pt per inch), scaled down to Excel's ceiling.
function notePictureSize(picture: NotePicture): { width: string; height: string } {
  const widthPx = Math.max(1, Math.round(picture.widthPx))
  const heightPx = Math.max(1, Math.round(picture.heightPx))
  let widthPt = (widthPx * 72) / 96
  let heightPt = (heightPx * 72) / 96
  const scale = Math.min(1, MAX_NOTE_PICTURE_PT / Math.max(widthPt, heightPt))
  widthPt *= scale
  heightPt *= scale
  const points = (value: number): string => `${Math.max(1, Math.round(value * 100) / 100)}pt`
  return { width: points(widthPt), height: points(heightPt) }
}

/// Writes a picture into a note: the media part, the relationship that
/// reaches it, and the fill the note's shape paints it with.
///
/// The relationship lives in the VML part's OWN rels — a different id space
/// from the worksheet's — so the id is allocated there and cannot collide
/// with a worksheet relationship.
async function writeNotePicture(
  pkg: MutableNotePackage,
  vmlPath: string,
  picture: NotePicture,
  touchedEntries: Set<string>,
): Promise<NoteShapePicture> {
  const extension = IMAGE_EXTENSIONS[picture.mediaType]
  const mediaPath = await allocatePartPath(pkg, 'xl/media/image', `.${extension}`)
  pkg.addBinary(mediaPath, Uint8Array.from(Buffer.from(picture.base64, 'base64')))
  touchedEntries.add(mediaPath)
  await registerContentTypeDefault(pkg, extension, picture.mediaType, touchedEntries)
  const relId = await appendRelationship(
    pkg,
    relsPathFor(vmlPath),
    IMAGE_REL_TYPE,
    relativeTarget(vmlPath, mediaPath),
  )
  const title = picture.name ? ` o:title="${escapeXml(picture.name)}"` : ''
  const size = notePictureSize(picture)
  return {
    fill: `<v:fill o:relid="${relId}"${title} recolor="t" rotate="t" type="frame"/>`,
    width: size.width,
    height: size.height,
  }
}

/// The picture each note shows, as its shape paints it.
///
/// A note that brings a picture of its own gets a new media part and a new
/// relationship; a note that says nothing about its picture keeps the one it
/// was already carrying; a note that says `null` loses it.
async function notePictureFills(
  pkg: MutableNotePackage,
  vmlPath: string,
  notes: readonly SheetNote[],
  carried: ReadonlyMap<string, NoteShapePicture>,
  touchedEntries: Set<string>,
): Promise<Map<string, NoteShapePicture>> {
  const fills = new Map<string, NoteShapePicture>()
  for (const note of notes) {
    const key = noteImageKey(note)
    if (note.picture === null) continue
    if (note.picture !== undefined) {
      fills.set(key, await writeNotePicture(pkg, vmlPath, note.picture, touchedEntries))
      continue
    }
    const previous = carried.get(key)
    if (previous !== undefined) fills.set(key, previous)
  }
  return fills
}

function noteShape(
  note: SheetNote,
  index: number,
  firstId: number,
  picture?: NoteShapePicture,
): string {
  // Anchor: from one column right of the cell, spanning ~3 columns / 4 rows.
  const anchor = [note.column + 1, 15, note.row, 2, note.column + 4, 15, note.row + 4, 2].join(',')
  // A shape showing a picture keeps neither the yellow fill nor the drop
  // shadow: painted behind, they would hide the picture the note exists for.
  // The picture rides on the shape as its fill, spelled the way the source
  // spelled it.
  const paint = picture
    ? ` o:insetmode="auto">${picture.fill}`
    : ' fillcolor="#ffffe1" o:insetmode="auto">' +
      '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>'
  // The frame is the picture's size: a fill is stretched to the frame it
  // paints, so a frame the size of the text box would squash the photograph.
  const size = picture
    ? `${picture.width ? `width:${picture.width};` : ''}${picture.height ? `height:${picture.height};` : ''}`
    : 'width:108pt;height:60pt;'
  return (
    `<v:shape id="_x0000_s${firstId + index}" type="#_x0000_t202"` +
    ' style="position:absolute;margin-left:80pt;margin-top:2pt;' +
    `${size}z-index:${index + 1};visibility:hidden"${paint}` +
    '<v:path o:connecttype="none"/>' +
    '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
    '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' +
    `<x:Anchor>${anchor}</x:Anchor>` +
    '<x:AutoFill>False</x:AutoFill>' +
    `<x:Row>${note.row}</x:Row><x:Column>${note.column}</x:Column></x:ClientData>` +
    '</v:shape>'
  )
}

/// Drops every Note-typed shape, keeping other legacy objects verbatim.
function stripNoteShapes(vmlXml: string): string {
  return vmlXml.replace(/<v:shape\b(?![^>]*\/>)[\s\S]*?<\/v:shape>/g, (shape) =>
    shape.includes('ObjectType="Note"') ? '' : shape,
  )
}

function ensureContentTypeOverride(contentTypes: string, partName: string): string {
  if (contentTypes.includes(`PartName="/${partName}"`)) return contentTypes
  return contentTypes.replace(
    '</Types>',
    `<Override PartName="/${partName}" ContentType="${COMMENTS_CONTENT_TYPE}"/></Types>`,
  )
}

function ensureVmlDefault(contentTypes: string): string {
  if (/<Default\b[^>]*Extension="vml"/.test(contentTypes)) return contentTypes
  return contentTypes.replace(
    '</Types>',
    '<Default Extension="vml"' +
      ' ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>',
  )
}

function appendRel(relsXml: string, id: string, type: string, target: string): string {
  return relsXml.replace(
    '</Relationships>',
    `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`,
  )
}

function removeRel(relsXml: string, type: string): string {
  return relsXml.replace(new RegExp(`<Relationship\\b[^>]*Type="${type}"[^>]*/?>`), '')
}

async function removeContentTypeOverride(
  pkg: MutableNotePackage,
  partPath: string,
  touchedEntries: Set<string>,
): Promise<void> {
  const contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
  const stripped = contentTypes.replace(
    new RegExp(
      `<Override\\b[^>]*PartName="/${partPath.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`,
    ),
    '',
  )
  if (stripped !== contentTypes) {
    pkg.write(CONTENT_TYPES_PATH, stripped)
    touchedEntries.add(CONTENT_TYPES_PATH)
  }
}

const EMPTY_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '</Relationships>'

/// CT_Worksheet order: legacyDrawing follows drawing and precedes these.
const AFTER_LEGACY_DRAWING =
  /<legacyDrawingHF\b|<picture\b|<oleObjects\b|<controls\b|<webPublishItems\b|<tableParts\b|<extLst\b/

function ensureLegacyDrawingElement(worksheetXml: string, rid: string): string {
  if (/<legacyDrawing\b/.test(worksheetXml)) return worksheetXml
  worksheetXml = ensureRelationshipNamespace(worksheetXml)
  const element = `<legacyDrawing r:id="${rid}"/>`
  const anchor = AFTER_LEGACY_DRAWING.exec(worksheetXml)
  if (anchor) {
    return worksheetXml.slice(0, anchor.index) + element + worksheetXml.slice(anchor.index)
  }
  const end = worksheetXml.lastIndexOf('</worksheet>')
  if (end === -1) throw new NoteEditError('Worksheet has no closing element.')
  return worksheetXml.slice(0, end) + element + worksheetXml.slice(end)
}

/// Replaces the worksheet's whole comment set (empty list removes it).
export async function applySheetNotes(
  pkg: MutableNotePackage,
  worksheetPath: string,
  notes: readonly SheetNote[],
  touchedEntries: Set<string>,
): Promise<void> {
  const relsPath = worksheetRelsPath(worksheetPath)
  const hasRels = await pkg.has(relsPath)
  let relsXml = hasRels ? await pkg.readText(relsPath) : EMPTY_RELS
  const existingCommentsTarget = relTarget(relsXml, COMMENTS_REL_TYPE)
  const existingCommentsPath =
    existingCommentsTarget === null ? null : resolveRelTarget(worksheetPath, existingCommentsTarget)
  const existingVmlTarget = relTarget(relsXml, VML_REL_TYPE)
  const existingVmlPath =
    existingVmlTarget === null ? null : resolveRelTarget(worksheetPath, existingVmlTarget)
  let relsChanged = false

  const existingLegacy =
    existingCommentsPath !== null && (await pkg.has(existingCommentsPath))
      ? parseLegacyComments(await pkg.readText(existingCommentsPath))
      : new Map<string, LegacyCommentEntry>()
  const threaded = await syncThreadedComments(
    pkg,
    worksheetPath,
    relsXml,
    new Map(notes.map((note) => [cellRef(note), note])),
    existingLegacy,
    touchedEntries,
  )
  if (threaded.relsXml !== relsXml) {
    relsXml = threaded.relsXml
    relsChanged = true
  }
  if (threaded.partRemoved && threaded.partPath !== null) {
    relsXml = removeRel(relsXml, THREADED_COMMENTS_REL_TYPE)
    relsChanged = true
    await removeContentTypeOverride(pkg, threaded.partPath, touchedEntries)
  }

  if (notes.length === 0) {
    if (existingCommentsPath === null) {
      if (relsChanged) {
        pkg.write(relsPath, relsXml)
        touchedEntries.add(relsPath)
      }
      return
    }
    pkg.remove(existingCommentsPath)
    touchedEntries.add(existingCommentsPath)
    relsXml = removeRel(relsXml, COMMENTS_REL_TYPE)
    await removeContentTypeOverride(pkg, existingCommentsPath, touchedEntries)
    if (existingVmlPath !== null && (await pkg.has(existingVmlPath))) {
      const remaining = stripNoteShapes(await pkg.readText(existingVmlPath))
      if (/<v:shape\b/.test(remaining)) {
        pkg.write(existingVmlPath, remaining)
      } else {
        pkg.remove(existingVmlPath)
        relsXml = removeRel(relsXml, VML_REL_TYPE)
        const worksheetXml = await pkg.readText(worksheetPath)
        const withoutLegacy = worksheetXml.replace(/<legacyDrawing\b[^>]*\/>/, '')
        if (withoutLegacy !== worksheetXml) {
          pkg.write(worksheetPath, withoutLegacy)
          touchedEntries.add(worksheetPath)
        }
      }
      touchedEntries.add(existingVmlPath)
    }
    pkg.write(relsPath, relsXml)
    touchedEntries.add(relsPath)
    return
  }

  // Comments part: rewrite in place, or allocate a fresh one.
  let commentsPath = existingCommentsPath
  if (commentsPath === null) {
    commentsPath = await nextFreePath(pkg, (index) => `xl/comments${index}.xml`)
    const rid = nextFreeRelationshipId(relsXml)
    const target = `../${commentsPath.replace(/^xl\//, '')}`
    relsXml = appendRel(relsXml, rid, COMMENTS_REL_TYPE, target)
    relsChanged = true
    pkg.add(commentsPath, buildCommentsXml(notes, threaded.legacy))
  } else {
    pkg.write(commentsPath, buildCommentsXml(notes, threaded.legacy))
  }
  touchedEntries.add(commentsPath)

  // VML part: keep foreign shapes, replace the note shapes.
  if (existingVmlPath !== null && (await pkg.has(existingVmlPath))) {
    const original = await pkg.readText(existingVmlPath)
    const carried = noteShapePictures(original)
    const vml = stripNoteShapes(original)
    const end = vml.lastIndexOf('</xml>')
    if (end === -1) throw new NoteEditError(`${existingVmlPath} is not a VML drawing.`)
    const firstId = nextShapeId(vml)
    const pictures = await notePictureFills(pkg, existingVmlPath, notes, carried, touchedEntries)
    const shapes = notes
      .map((note, index) => noteShape(note, index, firstId, pictures.get(noteImageKey(note))))
      .join('')
    pkg.write(existingVmlPath, vml.slice(0, end) + shapes + vml.slice(end))
    touchedEntries.add(existingVmlPath)
  } else {
    const vmlPath = await nextFreePath(pkg, (index) => `xl/drawings/vmlDrawing${index}.vml`)
    const rid = nextFreeRelationshipId(relsXml)
    relsXml = appendRel(relsXml, rid, VML_REL_TYPE, `../drawings/${vmlPath.split('/').pop()}`)
    relsChanged = true
    const pictures = await notePictureFills(pkg, vmlPath, notes, new Map(), touchedEntries)
    const shapes = notes
      .map((note, index) =>
        noteShape(note, index, FIRST_NOTE_SHAPE_ID, pictures.get(noteImageKey(note))),
      )
      .join('')
    pkg.add(vmlPath, `${VML_HEADER}${shapes}</xml>`)
    touchedEntries.add(vmlPath)
    const worksheetXml = await pkg.readText(worksheetPath)
    const withLegacy = ensureLegacyDrawingElement(worksheetXml, rid)
    if (withLegacy !== worksheetXml) {
      pkg.write(worksheetPath, withLegacy)
      touchedEntries.add(worksheetPath)
    }
  }

  const contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
  const updatedTypes = ensureVmlDefault(ensureContentTypeOverride(contentTypes, commentsPath))
  if (updatedTypes !== contentTypes) {
    pkg.write(CONTENT_TYPES_PATH, updatedTypes)
    touchedEntries.add(CONTENT_TYPES_PATH)
  }

  if (relsChanged) {
    if (hasRels) pkg.write(relsPath, relsXml)
    else pkg.add(relsPath, relsXml)
    touchedEntries.add(relsPath)
  }
}
