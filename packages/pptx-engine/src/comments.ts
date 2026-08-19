/**
 * Slide comments (classic comments part) read/write — archive surgery.
 *
 * Structure: ppt/commentAuthors.xml (author table, referenced by
 * presentation.xml.rels) + one ppt/comments/commentN.xml per slide (referenced by
 * slide rels). A comment is uniquely identified by (authorId, idx). All changes
 * land in archive.entries: savePptx persists automatically, and the main process's
 * snapshot-style undo covers them automatically.
 */
import type { OpenedPptx } from './index'
import { resolveTarget, type PackageArchive } from './zip'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'
import { appendRelationship, unescapeXml } from './notes'
import { removeRelationshipAndCollectOwnedTarget } from './resource-cleanup'

const XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const COMMENTS_REL = `${REL_BASE}/comments`
const AUTHORS_REL = `${REL_BASE}/commentAuthors`

const COMMENTS_CT = 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml'
const AUTHORS_CT = 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml'

const AUTHORS_PATH = 'ppt/commentAuthors.xml'

export interface SlideComment {
  authorId: number
  author: string
  initials: string
  /** ISO timestamp (the dt attribute in the part) */
  dt: string
  idx: number
  text: string
}

function setEntry(archive: PackageArchive, path: string, xml: string): void {
  archive.entries.set(path, Buffer.from(xml, 'utf8'))
}

function addContentTypeOverride(
  archive: PackageArchive,
  partPath: string,
  contentType: string,
): void {
  const ctPath = '[Content_Types].xml'
  const ct = archive.readText(ctPath)
  if (!ct || ct.includes(`PartName="/${partPath}"`)) return
  setEntry(
    archive,
    ctPath,
    ct.replace(
      '</Types>',
      `<Override PartName="/${partPath}" ContentType="${contentType}"/></Types>`,
    ),
  )
}

/** Author table: id → {name, initials}. */
function readAuthors(archive: PackageArchive): Map<number, { name: string; initials: string }> {
  const map = new Map<number, { name: string; initials: string }>()
  const xml = archive.readText(AUTHORS_PATH)
  if (!xml) return map
  for (const m of xml.matchAll(/<p:cmAuthor\b([^>]*)\/>/g)) {
    const attrs = m[1]!
    const id = Number(/\bid="(\d+)"/.exec(attrs)?.[1] ?? -1)
    if (id < 0) continue
    map.set(id, {
      name: unescapeXml(/\bname="([^"]*)"/.exec(attrs)?.[1] ?? ''),
      initials: unescapeXml(/\binitials="([^"]*)"/.exec(attrs)?.[1] ?? ''),
    })
  }
  return map
}

/** Relationship and path of a slide's comments part (null when absent). */
function commentsRelationshipForSlide(
  archive: PackageArchive,
  slidePath: string,
): { id: string; path: string } | null {
  for (const rel of archive.readRels(slidePath).values()) {
    if (rel.type === COMMENTS_REL) {
      return { id: rel.id, path: resolveTarget(slidePath, rel.target) }
    }
  }
  return null
}

/** Path of a slide's comments part (null if the slide has no comments). */
function commentsPathForSlide(archive: PackageArchive, slidePath: string): string | null {
  return commentsRelationshipForSlide(archive, slidePath)?.path ?? null
}

/** Read all comments on a slide (in order of appearance). */
export function getSlideComments(archive: PackageArchive, slidePath: string): SlideComment[] {
  const partPath = commentsPathForSlide(archive, slidePath)
  if (!partPath) return []
  const xml = archive.readText(partPath)
  if (!xml) return []
  const authors = readAuthors(archive)
  const out: SlideComment[] = []
  for (const m of xml.matchAll(/<p:cm\b([^>]*)>([\s\S]*?)<\/p:cm>/g)) {
    const attrs = m[1]!
    const authorId = Number(/\bauthorId="(\d+)"/.exec(attrs)?.[1] ?? 0)
    const a = authors.get(authorId)
    out.push({
      authorId,
      author: a?.name ?? 'Unknown author',
      initials: a?.initials ?? '',
      dt: /\bdt="([^"]*)"/.exec(attrs)?.[1] ?? '',
      idx: Number(/\bidx="(\d+)"/.exec(attrs)?.[1] ?? 0),
      text: unescapeXml(/<p:text>([\s\S]*?)<\/p:text>/.exec(m[2]!)?.[1] ?? ''),
    })
  }
  return out
}

/**
 * Ensure the author exists and return {authorId, nextIdx} (bumping lastIdx).
 * Creates the author table / content type / presentation rel if missing.
 */
function ensureAuthor(
  archive: PackageArchive,
  name: string,
  initials: string,
): { authorId: number; nextIdx: number } {
  let xml = archive.readText(AUTHORS_PATH)
  if (!xml) {
    xml = XMLDECL + `<p:cmAuthorLst xmlns:p="${NS_P}"></p:cmAuthorLst>`
    addContentTypeOverride(archive, AUTHORS_PATH, AUTHORS_CT)
    appendRelationship(archive, 'ppt/presentation.xml', AUTHORS_REL, 'commentAuthors.xml')
  }
  // Existing author with the same name: reuse its id, lastIdx + 1
  for (const m of xml.matchAll(/<p:cmAuthor\b([^>]*)\/>/g)) {
    const attrs = m[1]!
    if (unescapeXml(/\bname="([^"]*)"/.exec(attrs)?.[1] ?? '') !== name) continue
    const authorId = Number(/\bid="(\d+)"/.exec(attrs)?.[1] ?? 0)
    const lastIdx = Number(/\blastIdx="(\d+)"/.exec(attrs)?.[1] ?? 0)
    const nextIdx = lastIdx + 1
    const bumped = m[0].includes('lastIdx="')
      ? m[0].replace(/\blastIdx="\d+"/, `lastIdx="${nextIdx}"`)
      : m[0].replace('/>', ` lastIdx="${nextIdx}"/>`)
    setEntry(archive, AUTHORS_PATH, xml.replace(m[0], bumped))
    return { authorId, nextIdx }
  }
  // New author
  let maxId = -1
  for (const m of xml.matchAll(/<p:cmAuthor\b[^>]*\bid="(\d+)"/g))
    maxId = Math.max(maxId, Number(m[1]))
  const authorId = maxId + 1
  const tag =
    `<p:cmAuthor id="${authorId}" name="${escapeXmlAttr(name)}"` +
    ` initials="${escapeXmlAttr(initials)}" lastIdx="1" clrIdx="${authorId}"/>`
  setEntry(archive, AUTHORS_PATH, xml.replace('</p:cmAuthorLst>', `${tag}</p:cmAuthorLst>`))
  return { authorId, nextIdx: 1 }
}

/** Add a comment; returns the new comment (with assigned authorId/idx). */
export function addSlideComment(
  opened: OpenedPptx,
  slideIndex: number,
  opts: { author: string; initials?: string; text: string },
): SlideComment | null {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return null
  const { archive } = opened
  const initials =
    opts.initials ??
    opts.author
      .split(/\s+/)
      .map((w) => w[0] ?? '')
      .join('')
      .toUpperCase() ??
    ''
  const { authorId, nextIdx } = ensureAuthor(archive, opts.author, initials)

  let partPath = commentsPathForSlide(archive, slide.path)
  if (!partPath) {
    let maxNum = 0
    for (const path of archive.entries.keys()) {
      const m = /^ppt\/comments\/comment(\d+)\.xml$/.exec(path)
      if (m) maxNum = Math.max(maxNum, Number(m[1]))
    }
    partPath = `ppt/comments/comment${maxNum + 1}.xml`
    setEntry(archive, partPath, XMLDECL + `<p:cmLst xmlns:p="${NS_P}"></p:cmLst>`)
    addContentTypeOverride(archive, partPath, COMMENTS_CT)
    appendRelationship(archive, slide.path, COMMENTS_REL, `../${partPath.slice(4)}`)
  }
  const xml = archive.readText(partPath)
  if (!xml) return null

  const dt = new Date().toISOString()
  // Positions staggered along the top-left diagonal so comment badges don't overlap
  const count = [...xml.matchAll(/<p:cm\b/g)].length
  const pos = 10 + (count % 8) * 6
  const cm =
    `<p:cm authorId="${authorId}" dt="${dt}" idx="${nextIdx}">` +
    `<p:pos x="${pos}" y="${pos}"/>` +
    `<p:text>${escapeXmlText(opts.text)}</p:text></p:cm>`
  setEntry(archive, partPath, xml.replace('</p:cmLst>', `${cm}</p:cmLst>`))

  return { authorId, author: opts.author, initials, dt, idx: nextIdx, text: opts.text }
}

/** Delete a comment (by authorId + idx). */
export function deleteSlideComment(
  opened: OpenedPptx,
  slideIndex: number,
  ref: { authorId: number; idx: number },
): boolean {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return false
  const { archive } = opened
  const relationship = commentsRelationshipForSlide(archive, slide.path)
  if (!relationship) return false
  const partPath = relationship.path
  const xml = archive.readText(partPath)
  if (!xml) return false
  for (const m of xml.matchAll(/<p:cm\b([^>]*)>[\s\S]*?<\/p:cm>/g)) {
    const attrs = m[1]!
    if (
      Number(/\bauthorId="(\d+)"/.exec(attrs)?.[1] ?? -1) === ref.authorId &&
      Number(/\bidx="(\d+)"/.exec(attrs)?.[1] ?? -1) === ref.idx
    ) {
      const next = xml.slice(0, m.index!) + xml.slice(m.index! + m[0].length)
      setEntry(archive, partPath, next)
      if (!/<p:cm\b/.test(next)) {
        removeRelationshipAndCollectOwnedTarget(archive, slide.path, relationship.id)
      }
      return true
    }
  }
  return false
}
