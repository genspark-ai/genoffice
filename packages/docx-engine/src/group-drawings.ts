/**
 * Group and Ungroup on the saved markup: several anchored drawings become the
 * children of one wpg:wgp anchor, and a wpg:wgp anchor becomes one anchor per
 * child. Both work on the block's own bytes, so everything a drawing carries
 * (geometry, fills, text, picture relationships) goes along unread.
 */

const WPG_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

/** one drawing of a group: where it is and the box it takes inside the group, EMU */
export interface GroupMember {
  /** index of the body block holding the drawing */
  block: number
  /** which of the block's anchored drawings, in document order */
  drawing: number
  xEmu: number
  yEmu: number
  widthEmu: number
  heightEmu: number
}

export interface DrawingGroup {
  /** in document order: the first one's anchor becomes the group's */
  members: GroupMember[]
}

interface Holder {
  /** the w:drawing, or the mc:AlternateContent wrapping it */
  start: number
  end: number
  drawing: string
  /** false when the holder wraps the drawing's run too: what replaces it needs a run of its own */
  inRun: boolean
}

/** a drawing put where `holder` was, inside a run either way */
const inHolder = (holder: Holder, drawings: string): string =>
  holder.inRun ? drawings : `<w:r>${drawings}</w:r>`

/** The block's anchored drawings that are not inside another drawing or a fallback. */
function anchoredDrawings(xml: string): Holder[] {
  const out: Holder[] = []
  const tag = /<(\/?)(w:drawing|mc:AlternateContent|mc:Fallback)\b[^>]*>/g
  let depth = 0
  let fallback = 0
  let drawingStart = -1
  let wrapStart = -1
  let wrapDepth = 0
  let pending: { start: number; end: number } | null = null
  for (let m = tag.exec(xml); m; m = tag.exec(xml)) {
    const closing = m[1] === '/'
    const name = m[2]
    if (name === 'mc:Fallback') {
      if (depth === 0) fallback += closing ? -1 : 1
      continue
    }
    if (name === 'mc:AlternateContent') {
      if (depth > 0) continue
      if (!closing) {
        if (wrapDepth === 0) wrapStart = m.index
        wrapDepth++
      } else if (--wrapDepth === 0) {
        if (pending) {
          out.push({
            start: wrapStart,
            end: m.index + m[0].length,
            drawing: xml.slice(pending.start, pending.end),
            inRun: !/<w:r[\s>]/.test(xml.slice(wrapStart, pending.start)),
          })
        }
        pending = null
      }
      continue
    }
    if (fallback > 0) continue
    if (!closing) {
      if (depth++ === 0) drawingStart = m.index
    } else if (--depth === 0) {
      const end = m.index + m[0].length
      if (wrapDepth > 0) pending ??= { start: drawingStart, end }
      else
        out.push({ start: drawingStart, end, drawing: xml.slice(drawingStart, end), inRun: true })
    }
  }
  return out.filter((h) => /^<w:drawing>\s*<wp:anchor\b/.test(h.drawing))
}

/** the namespace declarations written on the tags of `head`, later ones winning */
function inlineNamespaces(head: string): Map<string, string> {
  const out = new Map<string, string>()
  for (const m of head.matchAll(/\sxmlns:([\w.-]+)\s*=\s*"([^"]*)"/g)) out.set(m[1], m[2])
  return out
}

const namespaceAttrs = (spaces: Map<string, string>): string =>
  Array.from(spaces, ([prefix, uri]) => ` xmlns:${prefix}="${uri}"`).join('')

interface DrawingParts {
  /** from <wp:anchor to just before <a:graphic */
  head: string
  /** from after </a:graphic> to the end of </wp:anchor> */
  tail: string
  /** what a:graphicData holds */
  content: string
  spaces: Map<string, string>
}

function drawingParts(drawing: string): DrawingParts | null {
  const anchorStart = drawing.indexOf('<wp:anchor')
  // not a:graphicFrameLocks, which Word writes ahead of it
  const graphicStart =
    anchorStart + drawing.slice(Math.max(anchorStart, 0)).search(/<a:graphic[\s>]/)
  const graphicEnd = drawing.lastIndexOf('</a:graphic>')
  const anchorEnd = drawing.lastIndexOf('</wp:anchor>')
  if (
    anchorStart < 0 ||
    graphicStart < anchorStart ||
    graphicEnd < graphicStart ||
    anchorEnd < graphicEnd
  )
    return null
  const dataOpen = /<a:graphicData\b[^>]*>/.exec(drawing.slice(graphicStart))
  const dataEnd = drawing.lastIndexOf('</a:graphicData>')
  if (!dataOpen || dataEnd < 0) return null
  const contentStart = graphicStart + dataOpen.index + dataOpen[0].length
  return {
    head: drawing.slice(anchorStart, graphicStart),
    tail: drawing.slice(graphicEnd + '</a:graphic>'.length, anchorEnd + '</wp:anchor>'.length),
    content: drawing.slice(contentStart, dataEnd).trim(),
    spaces: inlineNamespaces(drawing.slice(anchorStart, contentStart)),
  }
}

/** the shape properties element of a wps:wsp or pic:pic: [start of its content, end of its content] */
function shapePropsRange(child: string): [number, number] | null {
  const open = /<(wps|pic):spPr\b[^>]*?(\/?)>/.exec(child)
  if (!open || open[2] === '/') return null
  const start = open.index + open[0].length
  const end = child.indexOf(`</${open[1]}:spPr>`, start)
  return end < 0 ? null : [start, end]
}

const attrOf = (tag: string, name: string): number | null => {
  const m = new RegExp(`\\b${name}\\s*=\\s*"(-?\\d+)"`).exec(tag)
  return m ? parseInt(m[1], 10) : null
}

/** A child with its a:xfrm set to `box`; a turned shape keeps its own size around the box's middle. */
function placedChild(
  child: string,
  box: { x: number; y: number; cx: number; cy: number },
): string | null {
  const range = shapePropsRange(child)
  if (!range) return null
  const props = child.slice(range[0], range[1])
  const xfrm = /<a:xfrm\b([^>]*)>([\s\S]*?)<\/a:xfrm>/.exec(props)
  const round = Math.round
  if (!xfrm) {
    const added = `<a:xfrm><a:off x="${round(box.x)}" y="${round(box.y)}"/><a:ext cx="${round(box.cx)}" cy="${round(box.cy)}"/></a:xfrm>`
    return child.slice(0, range[0]) + added + child.slice(range[0])
  }
  const ext = /<a:ext\b[^>]*\/>/.exec(xfrm[2])?.[0] ?? ''
  const size = { cx: attrOf(ext, 'cx') ?? box.cx, cy: attrOf(ext, 'cy') ?? box.cy }
  // a rotation turns the shape about its middle: the measured box is the turned one
  const turned = /\brot\s*=\s*"-?[1-9]/.test(xfrm[1])
  const x = turned ? box.x + (box.cx - size.cx) / 2 : box.x
  const y = turned ? box.y + (box.cy - size.cy) / 2 : box.y
  const next =
    `<a:xfrm${xfrm[1]}><a:off x="${round(x)}" y="${round(y)}"/>` +
    `<a:ext cx="${round(size.cx)}" cy="${round(size.cy)}"/></a:xfrm>`
  const at = range[0] + xfrm.index
  return child.slice(0, at) + next + child.slice(at + xfrm[0].length)
}

/** Shifts a posOffset of the anchor head by `delta`; null when that axis is placed by alignment. */
function shiftedPosition(head: string, axis: 'H' | 'V', delta: number): string | null {
  const re = new RegExp(`(<wp:position${axis}\\b[^>]*>\\s*<wp:posOffset>)(-?\\d+)(</wp:posOffset>)`)
  if (!re.test(head)) return null
  return head.replace(re, (_, open: string, value: string, close: string) => {
    return `${open}${parseInt(value, 10) + Math.round(delta)}${close}`
  })
}

const withExtent = (head: string, cx: number, cy: number): string =>
  head.replace(/<wp:extent\b[^>]*\/>/, `<wp:extent cx="${Math.round(cx)}" cy="${Math.round(cy)}"/>`)

/** one replacement inside a block, by offsets into the block as it was */
interface Edit {
  block: number
  start: number
  end: number
  text: string
}

/** The range to cut for a drawing holder: the run around it too, when nothing else is in that run. */
function removalRange(xml: string, holder: Holder): [number, number] {
  const before = xml.slice(0, holder.start)
  const after = xml.slice(holder.end)
  // the nearest run only: an earlier run's properties must not be read as this one's
  const open = Array.from(before.matchAll(/<w:r(?:\s[^>]*)?>/g)).at(-1)
  const close = /^\s*<\/w:r>/.exec(after)
  const lead = open ? before.slice(open.index + open[0].length) : ''
  if (open && close && /^\s*(?:<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>)?\s*$/.test(lead))
    return [open.index, holder.end + close[0].length]
  return [holder.start, holder.end]
}

/** The blocks with the edits made; every edit addresses its block as it was before any of them. */
function withEdits(blocks: readonly string[], edits: Edit[]): string[] {
  const out = [...blocks]
  const byBlock = new Map<number, Edit[]>()
  for (const edit of edits) byBlock.set(edit.block, [...(byBlock.get(edit.block) ?? []), edit])
  for (const [block, list] of byBlock) {
    let xml = out[block]
    // back to front, so earlier offsets stay good
    for (const edit of list.sort((x, y) => y.start - x.start))
      xml = xml.slice(0, edit.start) + edit.text + xml.slice(edit.end)
    out[block] = xml
  }
  return out
}

/** the highest wp:docPr / cNvPr id in the markup, so new ones can follow it */
export function maxDrawingId(xml: string): number {
  let max = 0
  for (const m of xml.matchAll(/<(?:wp:docPr|\w+:cNvPr)\b[^>]*?\bid\s*=\s*"(\d+)"/g))
    max = Math.max(max, parseInt(m[1], 10))
  return max
}

/**
 * Merges the group's member drawings into one wpg:wgp anchor in the first
 * member's place; the other members leave their paragraphs (which stay).
 * Returns the blocks unchanged-in-length, or null when a member is missing or
 * is not a plain shape or picture.
 */
export function groupAnchoredDrawings(
  blocks: readonly string[],
  group: DrawingGroup,
  nextId: number,
): string[] | null {
  const edits = groupEdits(blocks, group, nextId)
  return edits && withEdits(blocks, edits)
}

/** the edits that merge one group, against the blocks as they are; null when it cannot be merged */
function groupEdits(blocks: readonly string[], group: DrawingGroup, nextId: number): Edit[] | null {
  if (group.members.length < 2) return null
  const holders = new Map<number, Holder[]>()
  const holderOf = (member: GroupMember): Holder | null => {
    const xml = blocks[member.block]
    if (xml === undefined) return null
    let list = holders.get(member.block)
    if (!list) holders.set(member.block, (list = anchoredDrawings(xml)))
    return list[member.drawing] ?? null
  }
  const spaces = new Map<string, string>()
  const children: string[] = []
  let first: DrawingParts | null = null
  let id = nextId
  for (const member of group.members) {
    const holder = holderOf(member)
    const parts = holder ? drawingParts(holder.drawing) : null
    if (!parts || !/^<(?:wps:wsp|pic:pic)\b/.test(parts.content)) return null
    first ??= parts
    for (const [prefix, uri] of parts.spaces) spaces.set(prefix, uri)
    let child = placedChild(parts.content, {
      x: member.xEmu,
      y: member.yEmu,
      cx: member.widthEmu,
      cy: member.heightEmu,
    })
    if (!child) return null
    // a grouped shape names itself: on its own the anchor's wp:docPr did
    if (child.startsWith('<wps:wsp') && !/^<wps:wsp\b[^>]*>\s*<wps:cNvPr\b/.test(child)) {
      const name = /<wp:docPr\b[^>]*?\bname\s*=\s*"([^"]*)"/.exec(parts.head)?.[1] ?? 'Shape'
      child = child.replace(
        /^<wps:wsp\b[^>]*>/,
        (open) => `${open}<wps:cNvPr id="${id++}" name="${name}"/>`,
      )
    }
    children.push(child)
  }
  if (!first) return null
  const lead = group.members[0]
  const cx = Math.max(...group.members.map((m) => m.xEmu + m.widthEmu))
  const cy = Math.max(...group.members.map((m) => m.yEmu + m.heightEmu))
  // the anchor described the first member's own corner; the group's is up and left of it
  let head = withExtent(first.head, cx, cy)
  head = shiftedPosition(head, 'H', -lead.xEmu) ?? head
  head = shiftedPosition(head, 'V', -lead.yEmu) ?? head
  head = head.replace(/(<wp:docPr\b[^>]*?\bname\s*=\s*")[^"]*(")/, `$1Group ${id}$2`)
  spaces.set('wpg', WPG_NS)
  spaces.set('wps', spaces.get('wps') ?? WPS_NS)
  spaces.set('pic', spaces.get('pic') ?? PIC_NS)
  spaces.delete('a')
  const size = `cx="${Math.round(cx)}" cy="${Math.round(cy)}"`
  const drawing =
    `<w:drawing>${head}<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="${WPG_NS}">` +
    `<wpg:wgp${namespaceAttrs(spaces)}><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/>` +
    `<a:ext ${size}/><a:chOff x="0" y="0"/><a:chExt ${size}/></a:xfrm></wpg:grpSpPr>` +
    `${children.join('')}</wpg:wgp></a:graphicData></a:graphic>` +
    // a size relative to the page belonged to the first member alone
    `${first.tail.replace(/<wp14:sizeRel[HV]\b[\s\S]*?<\/wp14:sizeRel[HV]>/g, '')}</w:drawing>`

  return group.members.map((member, index) => {
    const holder = holderOf(member)!
    if (index === 0) {
      const text = inHolder(holder, drawing)
      return { block: member.block, start: holder.start, end: holder.end, text }
    }
    const [start, end] = removalRange(blocks[member.block], holder)
    return { block: member.block, start, end, text: '' }
  })
}

function elementAt(xml: string, from: number): { name: string; end: number } | null {
  const open = /^<((?:wps:wsp|pic:pic|wpg:grpSp))\b[^>]*?(\/?)>/.exec(xml.slice(from))
  if (!open) return null
  const name = open[1]
  if (open[2] === '/') return { name, end: from + open[0].length }
  const tag = new RegExp(`<(/?)${name}\\b[^>]*?(/?)>`, 'g')
  tag.lastIndex = from
  let depth = 0
  for (let m = tag.exec(xml); m; m = tag.exec(xml)) {
    if (m[2] === '/') continue
    depth += m[1] === '/' ? -1 : 1
    if (depth === 0) return { name, end: m.index + m[0].length }
  }
  return null
}

export interface GroupXfrm {
  x: number
  y: number
  cx: number
  cy: number
  /** 60000ths of a degree, clockwise */
  rot: number
  flipH: boolean
  flipV: boolean
}

const TURN = 21600000

function xfrmOf(attrs: string, inner: string): GroupXfrm {
  const tagOf = (name: string) => new RegExp(`<a:${name}\\b[^>]*/>`).exec(inner)?.[0] ?? ''
  const on = (name: string) => new RegExp(`\\b${name}\\s*=\\s*"(?:1|true)"`).test(attrs)
  return {
    x: attrOf(tagOf('off'), 'x') ?? 0,
    y: attrOf(tagOf('off'), 'y') ?? 0,
    cx: attrOf(tagOf('ext'), 'cx') ?? 0,
    cy: attrOf(tagOf('ext'), 'cy') ?? 0,
    rot: attrOf(attrs, 'rot') ?? 0,
    flipH: on('flipH'),
    flipV: on('flipV'),
  }
}

/**
 * Where a child stands once its group is gone: scaled into the group's box,
 * mirrored, then turned about the box's middle, keeping its own size about
 * its middle. A mirror reverses the child's own turn.
 */
export function placedOutsideGroup(
  child: GroupXfrm,
  group: GroupXfrm,
  chOff: { x: number; y: number },
  chExt: { cx: number; cy: number },
): GroupXfrm {
  const sx = group.cx > 0 && chExt.cx > 0 ? group.cx / chExt.cx : 1
  const sy = group.cy > 0 && chExt.cy > 0 ? group.cy / chExt.cy : 1
  const cx = child.cx * sx
  const cy = child.cy * sy
  let mx = (child.x - chOff.x) * sx + cx / 2
  let my = (child.y - chOff.y) * sy + cy / 2
  if (group.flipH) mx = group.cx - mx
  if (group.flipV) my = group.cy - my
  const angle = (group.rot / TURN) * 2 * Math.PI
  const dx = mx - group.cx / 2
  const dy = my - group.cy / 2
  mx = group.cx / 2 + dx * Math.cos(angle) - dy * Math.sin(angle)
  my = group.cy / 2 + dx * Math.sin(angle) + dy * Math.cos(angle)
  const mirrored = group.flipH !== group.flipV
  const rot = (((group.rot + (mirrored ? -child.rot : child.rot)) % TURN) + TURN) % TURN
  return {
    x: group.x + mx - cx / 2,
    y: group.y + my - cy / 2,
    cx,
    cy,
    rot,
    flipH: child.flipH !== group.flipH,
    flipV: child.flipV !== group.flipV,
  }
}

function turnedAttrs(attrs: string, to: GroupXfrm): string {
  const rest = attrs.replace(/\s+(?:rot|flipH|flipV)\s*=\s*"[^"]*"/g, '')
  return (
    rest +
    (to.rot ? ` rot="${Math.round(to.rot)}"` : '') +
    (to.flipH ? ' flipH="1"' : '') +
    (to.flipV ? ' flipV="1"' : '')
  )
}

function childPropsRange(child: string): [number, number] | null {
  if (!child.startsWith('<wpg:grpSp')) return shapePropsRange(child)
  const open = /<wpg:grpSpPr\b[^>]*?(\/?)>/.exec(child)
  if (!open || open[1] === '/') return null
  const start = open.index + open[0].length
  const end = child.indexOf('</wpg:grpSpPr>', start)
  return end < 0 ? null : [start, end]
}

export interface UngroupedMember {
  drawing: string
  kind: 'shape' | 'picture' | 'group'
  /** the anchor's offsets, EMU */
  xEmu: number
  yEmu: number
  /** how many of the group's text boxes come before this child's own */
  textBoxesBefore: number
}

/** w:txbxContent elements not inside another one */
function topTextBoxes(xml: string): number {
  let count = 0
  let depth = 0
  for (const m of xml.matchAll(/<(\/?)w:txbxContent\b[^>]*?(\/?)>/g)) {
    if (m[2] === '/') {
      if (depth === 0) count++
    } else if (m[1] === '/') depth--
    else if (depth++ === 0) count++
  }
  return count
}

interface Ungrouped {
  holder: Holder
  members: UngroupedMember[]
}

/**
 * The children of the block's `drawing`-th anchored drawing, a group, each as
 * an anchor of its own where it stood; an inner group stays a group. Null
 * when it is not a group, holds anything but shapes, pictures and groups, or
 * is placed by alignment (no offsets to write the children's places from).
 */
function ungrouped(xml: string, drawing: number, nextId: number): Ungrouped | null {
  const holder = anchoredDrawings(xml)[drawing]
  const parts = holder ? drawingParts(holder.drawing) : null
  if (!holder || !parts) return null
  const open = /^<wpg:wgp\b[^>]*>/.exec(parts.content)
  const propsEnd = parts.content.indexOf('</wpg:grpSpPr>')
  if (!open || propsEnd < 0 || !parts.content.endsWith('</wpg:wgp>')) return null
  const groupProps = parts.content.slice(open[0].length, propsEnd)
  const xfrm = /<a:xfrm\b([^>]*)>([\s\S]*?)<\/a:xfrm>/.exec(groupProps)
  if (!xfrm) return null
  const group = xfrmOf(xfrm[1], xfrm[2])
  const tagOf = (name: string) => new RegExp(`<a:${name}\\b[^>]*/>`).exec(xfrm[2])?.[0] ?? ''
  const chOff = { x: attrOf(tagOf('chOff'), 'x') ?? 0, y: attrOf(tagOf('chOff'), 'y') ?? 0 }
  const chExt = {
    cx: attrOf(tagOf('chExt'), 'cx') ?? group.cx,
    cy: attrOf(tagOf('chExt'), 'cy') ?? group.cy,
  }
  const offsetOf = (axis: 'H' | 'V'): number | null => {
    const m = new RegExp(
      `<wp:position${axis}\\b[^>]*>\\s*<wp:posOffset>(-?\\d+)</wp:posOffset>`,
    ).exec(parts.head)
    return m ? parseInt(m[1], 10) : null
  }
  const anchorX = offsetOf('H')
  const anchorY = offsetOf('V')
  if (anchorX === null || anchorY === null) return null

  const spaces = new Map([...parts.spaces, ...inlineNamespaces(open[0])])
  spaces.delete('a')
  const body = parts.content.slice(propsEnd + '</wpg:grpSpPr>'.length, -'</wpg:wgp>'.length)
  const members: UngroupedMember[] = []
  let id = nextId
  let at = 0
  for (;;) {
    at += /^\s*/.exec(body.slice(at))![0].length
    if (at >= body.length) break
    const element = elementAt(body, at)
    if (!element) return null
    const textBoxesBefore = topTextBoxes(body.slice(0, at))
    let child = body.slice(at, element.end)
    at = element.end
    const range = childPropsRange(child)
    const own = range
      ? /<a:xfrm\b([^>]*)>([\s\S]*?)<\/a:xfrm>/.exec(child.slice(range[0], range[1]))
      : null
    if (!range || !own) return null
    const placed = placedOutsideGroup(xfrmOf(own[1], own[2]), group, chOff, chExt)
    const inner = element.name === 'wpg:grpSp'
    const size = `<a:ext cx="${Math.round(placed.cx)}" cy="${Math.round(placed.cy)}"/>`
    // an inner group keeps the space its own children are measured in
    const childSpace = inner
      ? (/<a:chOff\b[^>]*\/>/.exec(own[2])?.[0] ?? '') +
        (/<a:chExt\b[^>]*\/>/.exec(own[2])?.[0] ?? '')
      : ''
    const alone = `<a:xfrm${turnedAttrs(own[1], placed)}><a:off x="0" y="0"/>${size}${childSpace}</a:xfrm>`
    const xfrmAt = range[0] + own.index
    child = child.slice(0, xfrmAt) + alone + child.slice(xfrmAt + own[0].length)
    const name = /<\w+:cNvPr\b[^>]*?\bname\s*=\s*"([^"]*)"/.exec(child)?.[1]
    if (inner) {
      // on its own the group is named by its anchor
      child = child
        .replace(/^<wpg:grpSp\b/, '<wpg:wgp')
        .replace(/<\/wpg:grpSp>$/, '</wpg:wgp>')
        .replace(/^(<wpg:wgp\b[^>]*>)\s*<wpg:cNvPr\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/wpg:cNvPr>)/, '$1')
    }
    const xEmu = anchorX + Math.round(placed.x)
    const yEmu = anchorY + Math.round(placed.y)
    let head = shiftedPosition(parts.head, 'H', xEmu - anchorX)!
    head = shiftedPosition(head, 'V', yEmu - anchorY)!
    head = withExtent(head, placed.cx, placed.cy).replace(
      /<wp:docPr\b[^>]*?(\/?)>/,
      (_, selfClosing: string) =>
        `<wp:docPr id="${id++}" name="${name ?? `Shape ${id}`}"${selfClosing ? '/' : ''}>`,
    )
    const uri = inner ? WPG_NS : element.name === 'pic:pic' ? PIC_NS : WPS_NS
    members.push({
      drawing:
        `<w:drawing>${head}<a:graphic xmlns:a="${A_NS}"${namespaceAttrs(spaces)}>` +
        `<a:graphicData uri="${uri}">${child}</a:graphicData></a:graphic>${parts.tail}</w:drawing>`,
      kind: inner ? 'group' : element.name === 'pic:pic' ? 'picture' : 'shape',
      xEmu,
      yEmu,
      textBoxesBefore,
    })
  }
  return members.length > 0 ? { holder, members } : null
}

/** The group split into one anchor per child in the same run; null when it cannot be taken apart. */
export function ungroupAnchoredDrawing(
  xml: string,
  drawing: number,
  nextId: number,
): string | null {
  const split = ungrouped(xml, drawing, nextId)
  if (!split) return null
  const { holder, members } = split
  return (
    xml.slice(0, holder.start) +
    inHolder(holder, members.map((m) => m.drawing).join('')) +
    xml.slice(holder.end)
  )
}

export interface SplitGroupParagraph {
  members: UngroupedMember[]
  /**
   * One per child: the first is the paragraph as it was with the first child
   * in the group's place; the others hold only their drawing and are carried
   * (see `regroupDrawings`).
   */
  paragraphs: string[]
}

export function splitGroupParagraph(
  xml: string,
  drawing: number,
  nextId: number,
): SplitGroupParagraph | null {
  const split = ungrouped(xml, drawing, nextId)
  if (!split || !/^\s*<w:p[\s>]/.test(xml)) return null
  const { holder, members } = split
  return {
    members,
    paragraphs: [
      xml.slice(0, holder.start) + inHolder(holder, members[0].drawing) + xml.slice(holder.end),
      ...members.slice(1).map((m) => `<w:p><w:r>${m.drawing}</w:r></w:p>`),
    ],
  }
}

function paragraphContent(xml: string): { start: number; end: number } | null {
  const open = /^\s*<w:p(?:\s[^>]*)?>/.exec(xml)
  const end = xml.lastIndexOf('</w:p>')
  if (!open || end < 0 || xml.slice(end + '</w:p>'.length).trim() !== '') return null
  let start = open[0].length
  const props = /^\s*(?:<w:pPr\/>|<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>)/.exec(xml.slice(start))
  if (props) start += props[0].length
  return { start, end }
}

/**
 * Moves what each carried paragraph holds to the end of the paragraph before
 * it and drops the carried one: the file gets no paragraph the page did not
 * show. One with no paragraph before it stays.
 */
function mergeCarried(blocks: string[], carried: readonly number[]): string[] {
  const out = [...blocks]
  // front to back, so a run of them lands in order
  for (const block of [...carried].sort((a, b) => a - b)) {
    const from = paragraphContent(out[block] ?? '')
    if (!from) continue
    let before = block - 1
    while (before >= 0 && out[before] === '') before--
    const into = paragraphContent(out[before] ?? '')
    if (!into) continue
    out[before] =
      out[before].slice(0, into.end) +
      out[block].slice(from.start, from.end) +
      out[before].slice(into.end)
    out[block] = ''
  }
  return out
}

/**
 * Every group merged, each addressing its drawings as the blocks hold them
 * now, then the `carried` blocks folded into the paragraphs before them. A
 * group that cannot be made, or touches a drawing another already took, is
 * left out.
 */
export function regroupDrawings(
  blocks: readonly string[],
  groups: readonly DrawingGroup[],
  carried: readonly number[],
  nextId: number,
): string[] {
  const edits: Edit[] = []
  const clear = (candidates: Edit[]) =>
    candidates.every((c) =>
      edits.every((e) => e.block !== c.block || c.end <= e.start || c.start >= e.end),
    )
  let id = nextId
  for (const group of groups) {
    const merged = groupEdits(blocks, group, id)
    if (!merged || !clear(merged)) continue
    edits.push(...merged)
    id += group.members.length + 1
  }
  return mergeCarried(withEdits(blocks, edits), carried)
}
