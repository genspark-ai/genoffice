/**
 * Speaker notes (notesSlide) read/write — archive surgery, same approach as
 * duplicateSlide.
 *
 * - Read: slide rels → notesSlide part → body placeholder <a:t> text (\n-separated).
 * - Write: patch the existing notesSlide's body txBody; if there is no notesSlide,
 *   create one (creating a notesMaster too if needed and registering it in
 *   presentation.xml).
 * All changes land in archive.entries: savePptx persists automatically, and the
 * main process's snapshot-style undo (shallow copy of entries) covers them.
 */
import type { OpenedPptx } from './index'
import { relsPathFor, resolveTarget, type PackageArchive } from './zip'
import { escapeXmlText } from './xml-utils'

const XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NOTES_SLIDE_REL = `${REL_BASE}/notesSlide`
const NOTES_MASTER_REL = `${REL_BASE}/notesMaster`
const SLIDE_REL = `${REL_BASE}/slide`
const THEME_REL = `${REL_BASE}/theme`

const NOTES_SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
const NOTES_MASTER_CT =
  'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml'

function setEntry(archive: PackageArchive, path: string, xml: string): void {
  archive.entries.set(path, Buffer.from(xml, 'utf8'))
}

/** Unescape XML text (for reading <a:t>). */
export function unescapeXml(s: string): string {
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_m, h: string) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_m, d: string) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/** Path of the slide's notesSlide part (null if there is no notes slide). */
export function notesPathForSlide(archive: PackageArchive, slidePath: string): string | null {
  for (const rel of archive.readRels(slidePath).values()) {
    if (rel.type === NOTES_SLIDE_REL) return resolveTarget(slidePath, rel.target)
  }
  return null
}

/** Find the body placeholder sp block in the notesSlide xml. */
function findBodySp(xml: string): { xml: string; start: number; end: number } | null {
  for (const m of xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    if (/<p:ph\b[^>]*type="body"/.test(m[0])) {
      return { xml: m[0], start: m.index!, end: m.index! + m[0].length }
    }
  }
  return null
}

/** Read notes as plain text (body placeholder paragraphs joined by \n); '' if none. */
export function getSlideNotes(archive: PackageArchive, slidePath: string): string {
  const notesPath = notesPathForSlide(archive, slidePath)
  if (!notesPath) return ''
  const xml = archive.readText(notesPath)
  if (!xml) return ''
  const body = findBodySp(xml)
  if (!body) return ''
  const tx = /<p:txBody>([\s\S]*?)<\/p:txBody>/.exec(body.xml)?.[1]
  if (!tx) return ''
  const paras = [...tx.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)].map((p) =>
    [...p[1]!.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((t) => unescapeXml(t[1]!)).join(''),
  )
  // Drop trailing empty paragraphs (PowerPoint templates often carry an empty placeholder paragraph)
  while (paras.length && paras[paras.length - 1] === '') paras.pop()
  return paras.join('\n')
}

/** text (\n-separated) → notes txBody. */
function buildNotesTxBody(text: string): string {
  const lines = text.split('\n')
  const paras = lines.every((l) => l === '')
    ? '<a:p><a:endParaRPr lang="zh-CN"/></a:p>'
    : lines
        .map((line) =>
          line === ''
            ? '<a:p><a:endParaRPr lang="zh-CN"/></a:p>'
            : `<a:p><a:r><a:rPr lang="zh-CN" dirty="0"/><a:t>${escapeXmlText(line)}</a:t></a:r></a:p>`,
        )
        .join('')
  return `<p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody>`
}

const NOTES_BODY_SP_OPEN =
  '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder"/>' +
  '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
  '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>'

/**
 * Write notes (overwrite): patch the existing notesSlide's body txBody;
 * if there is no notesSlide, create the part (with a notesMaster if needed).
 */
export function setSlideNotes(opened: OpenedPptx, slideIndex: number, text: string): boolean {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return false
  const { archive } = opened
  const notesPath = notesPathForSlide(archive, slide.path) ?? createNotesSlide(opened, slide.path)
  if (!notesPath) return false
  const xml = archive.readText(notesPath)
  if (!xml) return false

  const txBody = buildNotesTxBody(text)
  const body = findBodySp(xml)
  let next: string
  if (body) {
    const patched = body.xml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, () => txBody)
    next = xml.slice(0, body.start) + patched + xml.slice(body.end)
  } else {
    next = xml.replace('</p:spTree>', `${NOTES_BODY_SP_OPEN}${txBody}</p:sp></p:spTree>`)
  }
  setEntry(archive, notesPath, next)
  return true
}

/** Add an Override to [Content_Types].xml (skipped if already present). */
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

/** Append a relationship to a part's rels, creating the rels file if missing. Returns the new rId. */
export function appendRelationship(
  archive: PackageArchive,
  partPath: string,
  type: string,
  target: string,
): string {
  const relsPath = relsPathFor(partPath)
  let xml =
    archive.readText(relsPath) ??
    XMLDECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  let maxRid = 0
  for (const m of xml.matchAll(/Id="rId(\d+)"/g)) maxRid = Math.max(maxRid, Number(m[1]))
  const rid = `rId${maxRid + 1}`
  xml = xml.replace(
    '</Relationships>',
    `<Relationship Id="${rid}" Type="${type}" Target="${target}"/></Relationships>`,
  )
  setEntry(archive, relsPath, xml)
  return rid
}

/** Ensure a notesMaster exists and return its part path. Created and registered in the presentation if missing. */
function ensureNotesMaster(archive: PackageArchive): string | null {
  for (const path of archive.entries.keys()) {
    if (/^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(path)) return path
  }
  const path = 'ppt/notesMasters/notesMaster1.xml'
  const emptyTree =
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>'
  const xml =
    XMLDECL +
    `<p:notesMaster xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">` +
    `<p:cSld>${emptyTree}</p:cSld>` +
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"' +
    ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    '</p:notesMaster>'
  setEntry(archive, path, xml)
  addContentTypeOverride(archive, path, NOTES_MASTER_CT)
  // rels: reuse an existing theme part (valid: multiple parts may point to the same theme)
  const theme = [...archive.entries.keys()].find((p) => /^ppt\/theme\/theme\d+\.xml$/.test(p))
  if (theme) appendRelationship(archive, path, THEME_REL, `../${theme.slice(4)}`)
  // Register notesMasterIdLst in presentation.xml
  const presPath = 'ppt/presentation.xml'
  const pres = archive.readText(presPath)
  if (pres && !pres.includes('<p:notesMasterIdLst>')) {
    const rid = appendRelationship(
      archive,
      presPath,
      NOTES_MASTER_REL,
      'notesMasters/notesMaster1.xml',
    )
    const lst = `<p:notesMasterIdLst><p:notesMasterId r:id="${rid}"/></p:notesMasterIdLst>`
    const next = pres.includes('</p:sldMasterIdLst>')
      ? pres.replace('</p:sldMasterIdLst>', `</p:sldMasterIdLst>${lst}`)
      : pres.replace('<p:sldIdLst>', `${lst}<p:sldIdLst>`)
    setEntry(archive, presPath, next)
  }
  return path
}

/** Create an empty notesSlide for a slide (part + rels + CT + slide rels); returns the part path. */
function createNotesSlide(opened: OpenedPptx, slidePath: string): string | null {
  const { archive } = opened
  if (!ensureNotesMaster(archive)) return null

  let maxNum = 0
  for (const path of archive.entries.keys()) {
    const m = /^ppt\/notesSlides\/notesSlide(\d+)\.xml$/.exec(path)
    if (m) maxNum = Math.max(maxNum, Number(m[1]))
  }
  const notesPath = `ppt/notesSlides/notesSlide${maxNum + 1}.xml`

  const xml =
    XMLDECL +
    `<p:notes xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">` +
    '<p:cSld><p:spTree>' +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder"/>' +
    '<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>' +
    '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>' +
    NOTES_BODY_SP_OPEN +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>' +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'
  setEntry(archive, notesPath, xml)
  addContentTypeOverride(archive, notesPath, NOTES_SLIDE_CT)

  // notesSlide rels: notesMaster + owning slide
  const master = [...archive.entries.keys()].find((p) =>
    /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(p),
  )!
  appendRelationship(archive, notesPath, NOTES_MASTER_REL, `../${master.slice(4)}`)
  appendRelationship(archive, notesPath, SLIDE_REL, `../${slidePath.slice(4)}`)

  // slide rels: point to the new notesSlide
  appendRelationship(archive, slidePath, NOTES_SLIDE_REL, `../${notesPath.slice(4)}`)
  return notesPath
}
