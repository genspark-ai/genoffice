/**
 * Audio/video / 3D model insertion — media part surgery.
 *
 * Video/audio: the modern PowerPoint dual-relationship form —
 *   <a:videoFile r:link> (2007 semantics) + p14:media r:embed (2010 semantics),
 *   both relationships pointing at the same embedded ppt/media/mediaN.ext part;
 *   blipFill is the poster frame.
 *
 * 3D models: am3d:model3d with a picture fallback, as PowerPoint writes them (model3d.ts).
 * Decks from before carry a plain picture whose cNvPr descr names the model part
 * (`aislides-3d:{partPath}`, model3dPartOf).
 */
import { deflateSync } from 'node:zlib'
import { crc32 } from './crc32'
import { glbBounds } from './glb-bounds'
import {
  GLB_CONTENT_TYPE,
  MODEL3D_REL_TYPE,
  model3dXml,
  setModel3dRotationXml,
  type Model3dRotation,
} from './model3d'
import type { EmuRect, Slide } from './types'
import { creationIdXml, escapeXmlAttr, hasDefaultFor, maxRelationshipIdNumber } from './xml-utils'
import { relsPathFor } from './zip'
import { appendRawElements, replacePictureBytes, type OpenedPptx } from './index'
import { nextCNvPrId } from './insert'

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const IMAGE_REL_TYPE = `${R_NS}/image`
const VIDEO_REL_TYPE = `${R_NS}/video`
const AUDIO_REL_TYPE = `${R_NS}/audio`
const MEDIA_REL_TYPE = 'http://schemas.microsoft.com/office/2007/relationships/media'
const P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
/** Fixed uri of the <p:ext> that hosts p14:media (ECMA/Microsoft convention) */
const MEDIA_EXT_URI = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}'

const MEDIA_MIME: Record<string, string> = {
  mp4: 'video/mp4',
  m4v: 'video/mp4',
  mov: 'video/quicktime',
  webm: 'video/webm',
  avi: 'video/x-msvideo',
  mp3: 'audio/mpeg',
  wav: 'audio/wav',
  m4a: 'audio/mp4',
  aac: 'audio/aac',
  ogg: 'audio/ogg',
}

// ── Solid-color PNG generation (poster frame fallback, no canvas dependency) ──

function pngChunk(type: string, data: Uint8Array): Buffer {
  const head = Buffer.alloc(8)
  head.writeUInt32BE(data.length, 0)
  head.write(type, 4, 'ascii')
  const crcBuf = Buffer.alloc(4)
  crcBuf.writeUInt32BE(crc32(Buffer.concat([Buffer.from(type, 'ascii'), data])), 0)
  return Buffer.concat([head, data, crcBuf])
}

/** Generate a w×h solid-color PNG (RGB, no alpha). Used as a poster frame placeholder. */
export function solidPng(w: number, h: number, rgb: [number, number, number]): Buffer {
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(w, 0)
  ihdr.writeUInt32BE(h, 4)
  ihdr[8] = 8 // bit depth
  ihdr[9] = 2 // color type RGB
  const raw = Buffer.alloc(h * (1 + w * 3))
  for (let y = 0; y < h; y++) {
    const row = y * (1 + w * 3)
    raw[row] = 0 // filter none
    for (let x = 0; x < w; x++) {
      raw[row + 1 + x * 3] = rgb[0]
      raw[row + 2 + x * 3] = rgb[1]
      raw[row + 3 + x * 3] = rgb[2]
    }
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', deflateSync(raw)),
    pngChunk('IEND', Buffer.alloc(0)),
  ])
}

// ── Shared part / rels surgery ─────────────────────────────────────────

export function ensureDefaultContentType(opened: OpenedPptx, ext: string, mime: string): void {
  const ctPath = '[Content_Types].xml'
  const ct = opened.archive.readText(ctPath)
  if (ct && !hasDefaultFor(ct, ext)) {
    const dflt = `<Default Extension="${ext}" ContentType="${mime}"/>`
    opened.archive.entries.set(
      ctPath,
      Buffer.from(
        ct.replace('</Types>', () => `${dflt}</Types>`),
        'utf8',
      ),
    )
  }
}

/** New media part: ppt/media/{prefix}N.ext, numbered max + 1 across the whole media dir. */
function newMediaPart(opened: OpenedPptx, prefix: string, ext: string, bytes: Uint8Array): string {
  let maxNum = 0
  for (const path of opened.archive.entries.keys()) {
    const m = /^ppt\/media\/[a-zA-Z]+(\d+)\./.exec(path)
    if (m) maxNum = Math.max(maxNum, Number(m[1]))
  }
  const partPath = `ppt/media/${prefix}${maxNum + 1}.${ext}`
  opened.archive.entries.set(partPath, bytes)
  return partPath
}

/** Append relationships to the slide rels; returns the assigned rIds (in order). */
function appendRels(
  opened: OpenedPptx,
  slide: Slide,
  rels: Array<{ type: string; target: string; external?: boolean }>,
): string[] {
  const relsPath = relsPathFor(slide.path)
  let xml =
    opened.archive.readText(relsPath) ??
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  let maxRid = maxRelationshipIdNumber(xml)
  const rids: string[] = []
  for (const rel of rels) {
    const rid = `rId${++maxRid}`
    rids.push(rid)
    const mode = rel.external ? ' TargetMode="External"' : ''
    xml = xml.replace(
      '</Relationships>',
      `<Relationship Id="${rid}" Type="${rel.type}" Target="${escapeXmlAttr(rel.target)}"${mode}/></Relationships>`,
    )
  }
  opened.archive.entries.set(relsPath, Buffer.from(xml, 'utf8'))
  return rids
}

// ── Audio/video ────────────────────────────────────────────────────────

export interface NewMediaOptions {
  kind: 'video' | 'audio'
  bytes: Uint8Array
  /** Lowercase extension (mp4/mov/webm/mp3/wav/m4a…) */
  ext: string
  /** Poster frame (defaults to a generated solid-color placeholder PNG) */
  poster?: { bytes: Uint8Array; ext: string }
  offset: EmuRect
  /** cNvPr name (defaults to a filename-style name) */
  name?: string
}

/** Insert audio/video: media/poster parts + dual relationships + <p:pic> fragment, append + reparse. */
export function addMedia(
  opened: OpenedPptx,
  slideIndex: number,
  opts: NewMediaOptions,
): { slide: Slide; elementId: string } | null {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return null
  const ext = opts.ext.toLowerCase()
  const mime = MEDIA_MIME[ext]
  if (!mime) return null

  // 1) media part + Content_Types
  const mediaPath = newMediaPart(opened, 'media', ext, opts.bytes)
  ensureDefaultContentType(opened, ext, mime)

  // 2) Poster frame part (solid color by default; square for audio, whose frame is PowerPoint's 64 pt icon)
  const poster = opts.poster ?? {
    bytes:
      opts.kind === 'video' ? solidPng(16, 9, [38, 38, 44]) : solidPng(16, 16, [240, 240, 244]),
    ext: 'png',
  }
  const posterExt = poster.ext.toLowerCase()
  const posterPath = newMediaPart(opened, 'image', posterExt, poster.bytes)
  ensureDefaultContentType(
    opened,
    posterExt,
    posterExt === 'jpg' || posterExt === 'jpeg' ? 'image/jpeg' : 'image/png',
  )

  // 3) Three relationships: poster image + legacy video/audio + p14 media
  const mediaTarget = `../media/${mediaPath.split('/').pop()}`
  const [ridPoster, ridLegacy, ridMedia] = appendRels(opened, slide, [
    { type: IMAGE_REL_TYPE, target: `../media/${posterPath.split('/').pop()}` },
    { type: opts.kind === 'video' ? VIDEO_REL_TYPE : AUDIO_REL_TYPE, target: mediaTarget },
    { type: MEDIA_REL_TYPE, target: mediaTarget },
  ]) as [string, string, string]

  // 4) <p:pic> fragment (videoFile/audioFile + p14:media extension hung on nvPr)
  const id = nextCNvPrId(slide)
  const name = opts.name ?? `${opts.kind === 'video' ? 'Video' : 'Audio'} ${id}`
  const fileTag = opts.kind === 'video' ? 'a:videoFile' : 'a:audioFile'
  const o = opts.offset
  const xml =
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${escapeXmlAttr(name)}">${creationIdXml()}</p:cNvPr>` +
    '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>' +
    `<p:nvPr><${fileTag} xmlns:r="${R_NS}" r:link="${ridLegacy}"/>` +
    `<p:extLst><p:ext uri="${MEDIA_EXT_URI}">` +
    `<p14:media xmlns:p14="${P14_NS}" xmlns:r="${R_NS}" r:embed="${ridMedia}"/>` +
    '</p:ext></p:extLst></p:nvPr></p:nvPicPr>' +
    `<p:blipFill><a:blip r:embed="${ridPoster}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${o.x}" y="${o.y}"/><a:ext cx="${o.cx}" cy="${o.cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'

  const r = appendRawElements(opened, slideIndex, [xml])
  return r ? { slide: r.slide, elementId: r.elementIds[r.elementIds.length - 1]! } : null
}

// ── 3D models ──────────────────────────────────────────────────────────

export interface NewModel3dOptions {
  /** glTF binary */
  bytes: Uint8Array
  ext: string
  /** The model's rendering: the Fallback picture and am3d:raster (a placeholder when omitted) */
  poster?: { bytes: Uint8Array; ext: string }
  offset: EmuRect
  name?: string
  rot?: Model3dRotation
}

/** Insert a 3D model as PowerPoint stores it (model3d.ts). Only .glb is embedded, as PowerPoint does. */
export function addModel3d(
  opened: OpenedPptx,
  slideIndex: number,
  opts: NewModel3dOptions,
): { slide: Slide; elementId: string } | null {
  const slide = opened.deck.slides[slideIndex]
  if (!slide || opts.ext.toLowerCase() !== 'glb') return null

  const modelPath = newMediaPart(opened, 'model3d', 'glb', opts.bytes)
  ensureDefaultContentType(opened, 'glb', GLB_CONTENT_TYPE)
  const poster = opts.poster ?? { bytes: solidPng(16, 16, [58, 58, 66]), ext: 'png' }
  const posterExt = poster.ext.toLowerCase()
  const posterPath = newMediaPart(opened, 'image', posterExt, poster.bytes)
  ensureDefaultContentType(
    opened,
    posterExt,
    posterExt === 'jpg' || posterExt === 'jpeg' ? 'image/jpeg' : 'image/png',
  )
  const [modelRId, posterRId] = appendRels(opened, slide, [
    { type: MODEL3D_REL_TYPE, target: `../media/${modelPath.split('/').pop()}` },
    { type: IMAGE_REL_TYPE, target: `../media/${posterPath.split('/').pop()}` },
  ]) as [string, string]

  const id = nextCNvPrId(slide)
  const xml = model3dXml({
    id,
    name: opts.name ?? `3D Model ${id}`,
    offset: opts.offset,
    modelRId,
    posterRId,
    bounds: glbBounds(opts.bytes),
    ...(opts.rot ? { rot: opts.rot } : {}),
  })
  const r = appendRawElements(opened, slideIndex, [xml])
  return r ? { slide: r.slide, elementId: r.elementIds[r.elementIds.length - 1]! } : null
}

/**
 * Turns a 3D model to `rot` and, with `poster`, replaces its rendering (the Fallback picture
 * and am3d:raster): the editor renders the model, the engine only stores what it drew.
 */
export function setModel3dView(
  opened: OpenedPptx,
  slide: Slide,
  elementId: string,
  view: { rot: Model3dRotation; poster?: { bytes: Uint8Array; ext: string } },
): boolean {
  const el = slide.elements.find((e) => e.id === elementId)
  if (el?.type !== 'picture' || !el.model3d) return false
  if (
    view.poster &&
    !replacePictureBytes(opened, slide, elementId, view.poster.bytes, view.poster.ext)
  ) {
    return false
  }
  el.anchor.originalXml = setModel3dRotationXml(el.anchor.originalXml, view.rot)
  el.model3d.rot = { ...view.rot }
  slide.structureDirty = true
  return true
}

/** The model part of a picture placeholder written before models were stored as am3d. */
export function model3dPartOf(el: { descr?: string }): string | null {
  const m = el.descr && /^aislides-3d:(.+)$/.exec(el.descr)
  return m ? m[1]! : null
}
