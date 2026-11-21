/**
 * Reduce File Size: which raster pictures of a package are stored at more
 * pixels than the largest size any drawing shows them at needs. The
 * resampling itself needs a canvas and happens in the host.
 */
import JSZip from 'jszip'
import { lazyMediaHashOf, lazyMediaUrl } from './lazy-media'
import { parseRels, resolveRelationshipTargetPath } from './parse-package'
import { loadDocxZip } from './zip-load'

const EMU_PER_INCH = 914400

export type MediaKeepReason =
  /** not PNG or JPEG */
  | 'format'
  /** shown somewhere whose size cannot be read (tiled fill, VML group, chart, …), or nowhere */
  | 'unmeasured'

/** the part of a picture a crop cuts away on each side, as fractions of its width / height */
export interface MediaCrop {
  l: number
  t: number
  r: number
  b: number
}

export interface MediaReductionTarget {
  partPath: string
  mime: 'image/png' | 'image/jpeg'
  /** pixels the largest use needs of the whole picture at the target resolution; Infinity keeps the stored size */
  widthPx: number
  heightPx: number
  /** set when every use of the part shows the same cropped part of it, so the rest can go */
  crop?: MediaCrop
  /** the part's bytes; absent when the part is a lazy-media placeholder */
  data?: Uint8Array
  /** where the host reads a lazily served part from */
  lazyUrl?: string
}

export interface MediaReductionPlan {
  targets: MediaReductionTarget[]
  kept: Array<{ partPath: string; reason: MediaKeepReason }>
}

interface Need {
  /** inches, already divided by the visible fraction of a cropped picture */
  width: number
  height: number
  unmeasured: boolean
  /** the crop of every sized use; null for a use that is not a plainly cropped picture */
  crops: Array<MediaCrop | null>
  /** used somewhere a crop cannot be taken out of (a fallback, VML) */
  mirrored: boolean
}

interface Usage {
  cx: number
  cy: number
  cropX: number
  cropY: number
  tiled: boolean
  crop: MediaCrop | null
  /** a picture's own fill, not stretched past its frame */
  plain: boolean
}

interface Frame {
  name: string
  drawing?: { cx: number; cy: number }
  vml?: { width: number; height: number } | null
  usages?: Usage[]
}

const TAG = /<(\/?)([A-Za-z_][\w.-]*(?::[\w.-]+)?)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g
const ATTR = /([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g
const VML_SHAPES = new Set(['v:shape', 'v:rect', 'v:roundrect', 'v:oval', 'v:image'])
const VML_UNIT_INCHES: Record<string, number> = {
  in: 1,
  pt: 1 / 72,
  pc: 1 / 6,
  cm: 1 / 2.54,
  mm: 1 / 25.4,
  px: 1 / 96,
}

function attrsOfTag(text: string): Map<string, string> {
  const out = new Map<string, string>()
  for (const m of text.matchAll(ATTR)) out.set(m[1], m[2] ?? m[3] ?? '')
  return out
}

function vmlSize(style: string | undefined): { width: number; height: number } | null {
  const read = (prop: string): number => {
    const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([\\d.]+)(in|pt|pc|cm|mm|px)\\s*(?:;|$)`).exec(
      style ?? '',
    )
    return m ? Number(m[1]) * VML_UNIT_INCHES[m[2]] : 0
  }
  const width = read('width')
  const height = read('height')
  return width > 0 && height > 0 ? { width, height } : null
}

/** share of the picture a crop rectangle leaves visible along one axis (1 = all of it) */
function visibleShare(a: string | undefined, b: string | undefined): number {
  const cut = (Number(a) || 0) + (Number(b) || 0)
  return Math.min(1, 1 - cut / 100000)
}

/**
 * Every use of the image relationships of one part, as the size it needs.
 * A use this cannot size marks the media part unmeasured; a use inside
 * mc:Fallback mirrors its mc:Choice and constrains nothing.
 */
function collectNeeds(
  xml: string,
  mediaByRel: Map<string, string>,
  needs: Map<string, Need>,
): void {
  const needOf = (part: string): Need => {
    let need = needs.get(part)
    if (!need) {
      need = { width: 0, height: 0, unmeasured: false, crops: [], mirrored: false }
      needs.set(part, need)
    }
    return need
  }
  const stack: Frame[] = []
  const pendingUsages: Array<{ part: string; usage: Usage }> = []
  for (const m of xml.matchAll(TAG)) {
    const [, closing, name, attrText, selfClosing] = m
    if (closing) {
      for (let i = stack.length - 1; i >= 0; i--) {
        if (stack[i].name !== name) continue
        stack.length = i
        break
      }
      continue
    }
    const attrs = attrsOfTag(attrText)
    const parent = stack[stack.length - 1]
    const frame: Frame = { name }
    if (name === 'wp:inline' || name === 'wp:anchor') frame.drawing = { cx: 0, cy: 0 }
    else if (name === 'wp:extent' && parent?.drawing) {
      parent.drawing.cx = Number(attrs.get('cx')) || 0
      parent.drawing.cy = Number(attrs.get('cy')) || 0
    } else if (VML_SHAPES.has(name)) frame.vml = vmlSize(attrs.get('style'))
    else if (name === 'a:srcRect' && parent?.usages) {
      const side = (edge: string): number => (Number(attrs.get(edge)) || 0) / 100000
      for (const usage of parent.usages) {
        usage.cropX = visibleShare(attrs.get('l'), attrs.get('r'))
        usage.cropY = visibleShare(attrs.get('t'), attrs.get('b'))
        usage.crop = { l: side('l'), t: side('t'), r: side('r'), b: side('b') }
      }
    } else if (name === 'a:tile' && parent?.usages) {
      for (const usage of parent.usages) usage.tiled = true
    } else if (name === 'a:fillRect') {
      // negative insets stretch the picture past the shape it fills
      const fill = stack[stack.length - 2]?.usages
      const grow = (a: string | undefined, b: string | undefined): number =>
        Math.max(1, 1 - ((Number(a) || 0) + (Number(b) || 0)) / 100000)
      const inset = ['l', 't', 'r', 'b'].some((a) => Number(attrs.get(a)) || 0)
      for (const usage of fill ?? []) {
        usage.cropX /= grow(attrs.get('l'), attrs.get('r'))
        usage.cropY /= grow(attrs.get('t'), attrs.get('b'))
        if (inset) usage.plain = false
      }
    }

    const inFallback = stack.some((f) => f.name === 'mc:Fallback')
    for (const [attr, value] of attrs) {
      const part = mediaByRel.get(value)
      if (!part || attr.startsWith('xmlns')) continue
      if (inFallback) {
        needOf(part).mirrored = true
        continue
      }
      const drawing = [...stack].reverse().find((f) => f.drawing)?.drawing
      if (name === 'a:blip' && /:embed$/.test(attr) && drawing && parent) {
        const usage: Usage = {
          cx: drawing.cx,
          cy: drawing.cy,
          cropX: 1,
          cropY: 1,
          tiled: false,
          crop: null,
          plain: parent.name === 'pic:blipFill',
        }
        ;(parent.usages ??= []).push(usage)
        pendingUsages.push({ part, usage })
        continue
      }
      const shape = [...stack].reverse().find((f) => f.vml !== undefined)?.vml
      const cropped = ['croptop', 'cropbottom', 'cropleft', 'cropright'].some((a) => attrs.has(a))
      const grouped = stack.some((f) => f.name === 'v:group')
      const need = needOf(part)
      need.mirrored = true
      if (name === 'v:imagedata' && shape && !cropped && !grouped) {
        need.width = Math.max(need.width, shape.width)
        need.height = Math.max(need.height, shape.height)
      } else {
        need.unmeasured = true
      }
    }
    if (!selfClosing) stack.push(frame)
  }
  // crop and tile siblings follow their a:blip, so the sizes settle only now
  for (const { part, usage } of pendingUsages) {
    const need = needOf(part)
    if (usage.tiled || usage.cx <= 0 || usage.cy <= 0 || usage.cropX < 0.01 || usage.cropY < 0.01) {
      need.unmeasured = true
      continue
    }
    need.width = Math.max(need.width, usage.cx / EMU_PER_INCH / usage.cropX)
    need.height = Math.max(need.height, usage.cy / EMU_PER_INCH / usage.cropY)
    need.crops.push(usage.plain ? usage.crop : null)
  }
}

/** the one crop all uses of a part share, when cutting the picture down to it changes nothing on the page */
function sharedCrop(need: Need): MediaCrop | undefined {
  const [first, ...rest] = need.crops
  if (!first || need.mirrored) return undefined
  const sides = [first.l, first.t, first.r, first.b]
  if (sides.some((v) => v < 0) || !sides.some((v) => v > 0)) return undefined
  if (first.l + first.r >= 0.99 || first.t + first.b >= 0.99) return undefined
  const same = (c: MediaCrop | null): boolean =>
    !!c && c.l === first.l && c.t === first.t && c.r === first.r && c.b === first.b
  return rest.every(same) ? first : undefined
}

function rasterMime(partPath: string): MediaReductionTarget['mime'] | null {
  const ext = partPath.split('.').pop()?.toLowerCase()
  if (ext === 'png') return 'image/png'
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg'
  return null
}

/**
 * For every picture part of the package, the pixel size its largest use
 * needs at `ppi`, or why it is left alone. All parts with relationships are
 * read (body, headers, footers, notes, comments, and anything else), so a
 * picture shared by several drawings is sized by the largest of them.
 * A null `ppi` keeps every picture at the resolution it is stored at.
 */
export async function planMediaReduction(
  bytes: Uint8Array,
  ppi: number | null,
  options: { deleteCropped?: boolean } = {},
): Promise<MediaReductionPlan> {
  const zip = await loadDocxZip(bytes)
  const needs = new Map<string, Need>()
  for (const relsPath of Object.keys(zip.files)) {
    const m = /^((?:.*\/)?)_rels\/([^/]+)\.rels$/.exec(relsPath)
    if (!m) continue
    const sourcePath = `${m[1]}${m[2]}`
    const mediaByRel = new Map<string, string>()
    for (const [id, rel] of await parseRels(zip, relsPath)) {
      if (rel.targetMode === 'External' || !/\/image$/.test(rel.type)) continue
      const part = resolveRelationshipTargetPath(sourcePath, rel.target)
      if (part && zip.file(part)) mediaByRel.set(id, part)
    }
    if (mediaByRel.size === 0) continue
    const source = zip.file(sourcePath)
    if (!source) continue
    collectNeeds(await source.async('string'), mediaByRel, needs)
  }

  const plan: MediaReductionPlan = { targets: [], kept: [] }
  for (const [partPath, need] of [...needs].sort(([a], [b]) => a.localeCompare(b))) {
    const mime = rasterMime(partPath)
    if (!mime) {
      plan.kept.push({ partPath, reason: 'format' })
      continue
    }
    if (need.unmeasured || need.width <= 0 || need.height <= 0) {
      plan.kept.push({ partPath, reason: 'unmeasured' })
      continue
    }
    const data = await zip.file(partPath)!.async('uint8array')
    const lazyHash = lazyMediaHashOf(data)
    const crop = options.deleteCropped ? sharedCrop(need) : undefined
    plan.targets.push({
      partPath,
      mime,
      widthPx: ppi === null ? Infinity : Math.ceil(need.width * ppi),
      heightPx: ppi === null ? Infinity : Math.ceil(need.height * ppi),
      ...(crop ? { crop } : {}),
      ...(lazyHash ? { lazyUrl: lazyMediaUrl(lazyHash, partPath) } : { data }),
    })
  }
  return plan
}

export interface RasterInfo {
  mime: 'image/png' | 'image/jpeg'
  width: number
  height: number
  /** a canvas round trip keeps what the picture shows: one frame, upright, 1 or 3 colour channels */
  resamplable: boolean
}

function pngInfo(bytes: Uint8Array): RasterInfo | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (bytes.length < 33 || view.getUint32(12) !== 0x49484452) return null
  let animated = false
  for (let at = 8; at + 12 <= bytes.length;) {
    const type = view.getUint32(at + 4)
    if (type === 0x6163544c) animated = true
    if (type === 0x49444154 || type === 0x49454e44) break
    at += 12 + view.getUint32(at)
  }
  return {
    mime: 'image/png',
    width: view.getUint32(16),
    height: view.getUint32(20),
    resamplable: !animated,
  }
}

/** null when the APP1 segment is not Exif (XMP is APP1 too) and so says nothing */
function exifOrientation(view: DataView, at: number, length: number): number | null {
  if (length < 14 || view.getUint32(at) !== 0x45786966) return null
  const tiff = at + 6
  const little = view.getUint16(tiff) === 0x4949
  const ifd = tiff + view.getUint32(tiff + 4, little)
  if (ifd + 2 > at + length) return 1
  const count = view.getUint16(ifd, little)
  for (let i = 0; i < count; i++) {
    const entry = ifd + 2 + i * 12
    if (entry + 12 > at + length) break
    if (view.getUint16(entry, little) === 0x0112) return view.getUint16(entry + 8, little)
  }
  return 1
}

function jpegInfo(bytes: Uint8Array): RasterInfo | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let orientation = 1
  for (let at = 2; at + 4 <= bytes.length;) {
    if (view.getUint8(at) !== 0xff) return null
    const marker = view.getUint8(at + 1)
    if (marker === 0xff) {
      at += 1
      continue
    }
    const length = view.getUint16(at + 2)
    if (marker === 0xe1) orientation = exifOrientation(view, at + 4, length - 2) ?? orientation
    const frame = marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)
    if (frame) {
      if (at + 10 > bytes.length) return null
      const channels = view.getUint8(at + 9)
      return {
        mime: 'image/jpeg',
        width: view.getUint16(at + 7),
        height: view.getUint16(at + 5),
        resamplable: orientation === 1 && (channels === 1 || channels === 3),
      }
    }
    if (marker === 0xda) return null
    at += 2 + length
  }
  return null
}

/** what the bytes themselves say they are: an extension can lie */
export function rasterInfo(bytes: Uint8Array): RasterInfo | null {
  try {
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47)
      return pngInfo(bytes)
    if (bytes[0] === 0xff && bytes[1] === 0xd8) return jpegInfo(bytes)
  } catch {
    /* truncated header */
  }
  return null
}

/**
 * The pixel size to resample a picture to so that it still covers `need` in
 * both directions, or null when it is not stored any larger than that.
 */
export function reducedPixelSize(
  stored: { width: number; height: number },
  need: { widthPx: number; heightPx: number },
): { width: number; height: number } | null {
  if (stored.width <= 0 || stored.height <= 0) return null
  const scale = Math.max(need.widthPx / stored.width, need.heightPx / stored.height)
  if (!(scale < 1)) return null
  const width = Math.max(1, Math.ceil(stored.width * scale))
  const height = Math.max(1, Math.ceil(stored.height * scale))
  return width < stored.width || height < stored.height ? { width, height } : null
}

/**
 * The pixels of a stored picture a crop leaves visible, and the size the
 * visible part is needed at (never larger than it is stored).
 */
export function croppedPixelRect(
  stored: { width: number; height: number },
  target: Pick<MediaReductionTarget, 'widthPx' | 'heightPx'> & { crop: MediaCrop },
): { sx: number; sy: number; sw: number; sh: number; width: number; height: number } {
  const { crop } = target
  const sx = Math.min(stored.width - 1, Math.max(0, Math.round(crop.l * stored.width)))
  const sy = Math.min(stored.height - 1, Math.max(0, Math.round(crop.t * stored.height)))
  const sw = Math.min(
    stored.width - sx,
    Math.max(1, Math.round((1 - crop.l - crop.r) * stored.width)),
  )
  const sh = Math.min(
    stored.height - sy,
    Math.max(1, Math.round((1 - crop.t - crop.b) * stored.height)),
  )
  const smaller = reducedPixelSize(
    { width: sw, height: sh },
    {
      widthPx: target.widthPx * (1 - crop.l - crop.r),
      heightPx: target.heightPx * (1 - crop.t - crop.b),
    },
  )
  return { sx, sy, sw, sh, width: smaller?.width ?? sw, height: smaller?.height ?? sh }
}

const BLIP_FILL = /<pic:blipFill\b[^>]*>[\s\S]*?<\/pic:blipFill>/g
const SRC_RECT = /<a:srcRect\b[^>]*?(?:\/>|>\s*<\/a:srcRect>)/g

/** takes the crop off every picture showing one of `uncropped`, whose bytes now hold only the visible part */
async function dropCrops(zip: JSZip, uncropped: ReadonlySet<string>): Promise<void> {
  for (const relsPath of Object.keys(zip.files)) {
    const m = /^((?:.*\/)?)_rels\/([^/]+)\.rels$/.exec(relsPath)
    const source = m && zip.file(`${m[1]}${m[2]}`)
    if (!m || !source) continue
    const ids = new Set<string>()
    for (const [id, rel] of await parseRels(zip, relsPath)) {
      if (rel.targetMode === 'External') continue
      const part = resolveRelationshipTargetPath(`${m[1]}${m[2]}`, rel.target)
      if (part && uncropped.has(part)) ids.add(id)
    }
    if (ids.size === 0) continue
    const xml = await source.async('string')
    const next = xml.replace(BLIP_FILL, (fill) => {
      const id = /<a:blip\b[^>]*?\br:embed="([^"]*)"/.exec(fill)?.[1]
      return id && ids.has(id) ? fill.replace(SRC_RECT, '') : fill
    })
    if (next !== xml) zip.file(source.name, next, { date: source.date })
  }
}

/**
 * The package with the bytes of some parts replaced under the same names.
 * The parts in `uncropped` were cut down to what their crop showed, so the
 * crop comes off the pictures that use them.
 */
export async function replaceDocxParts(
  bytes: Uint8Array,
  parts: Record<string, Uint8Array>,
  uncropped: Iterable<string> = [],
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  for (const [name, data] of Object.entries(parts)) {
    const entry = zip.file(name)
    if (!entry) throw new Error(`no such part: ${name}`)
    zip.file(name, data, { date: entry.date })
  }
  const cropless = new Set(uncropped)
  if (cropless.size > 0) await dropCrops(zip, cropless)
  for (const [name, entry] of Object.entries(zip.files)) if (entry.dir) delete zip.files[name]
  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}
