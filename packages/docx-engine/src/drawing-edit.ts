// Modified for ForaOffice: source-addressed DrawingML edits in parent coordinates.
import type { TextboxDisplay } from './types'
import { patchShapeProperties, xmlSegments, type ShapeStylePatch } from './generate'
import { attrsOf, childrenOf, nameOf, xmlParser, type XNode } from './xml-utils'
import { vmlGroupScale, type VmlGroupScale } from './parse-vml'

const EMU_PER_PX = 9525
const numAttr = (xml: string, key: string) =>
  Number(new RegExp(`\\b${key}="(-?[\\d.]+)"`).exec(xml)?.[1] ?? 0)
const setAttr = (tag: string, key: string, value: string | number): string => {
  const re = new RegExp(`\\b${key}="[^"]*"`)
  return re.test(tag)
    ? tag.replace(re, `${key}="${value}"`)
    : tag.replace(/\s*\/?>$/, (end) => ` ${key}="${value}"${end}`)
}

/** Only changes to source-addressed boxes; untouched shapes keep their exact bytes. */
export function patchDrawingShapes(
  xml: string,
  before: readonly TextboxDisplay[],
  after: readonly TextboxDisplay[],
): string {
  const edits: Array<{ start: number; end: number; xml: string }> = []
  const fallbacks = xmlSegments(xml, 'mc:Fallback', 0, xml.length)
  const drawings = xmlSegments(xml, 'w:drawing', 0, xml.length).filter(
    (s) => !fallbacks.some((f) => s.start > f.start && s.end < f.end),
  )
  after.forEach((box, i) => {
    const old = before[i]
    const source = old?.drawingSource
    if (!source) return
    const style: ShapeStylePatch = {}
    if (box.fill !== old.fill) style.fillHex = box.fill ?? null
    if (box.borderColor !== old.borderColor) style.borderHex = box.borderColor ?? null
    if ((box.rotDeg ?? 0) !== (old.rotDeg ?? 0)) style.rotDeg = box.rotDeg ?? null
    if (box.borderWidthPx !== old.borderWidthPx) style.borderWidthPx = box.borderWidthPx ?? null
    const width = box.widthPx !== old.widthPx && box.widthPx !== undefined ? box.widthPx : undefined
    const height =
      box.heightPx !== old.heightPx && box.heightPx !== undefined ? box.heightPx : undefined
    const dx = (box.offsetXEmu ?? 0) - (old.offsetXEmu ?? 0)
    const dy = (box.offsetYEmu ?? 0) - (old.offsetYEmu ?? 0)
    if (width === undefined && height === undefined && !dx && !dy && !Object.keys(style).length)
      return
    const drawing = drawings[source.drawingIndex]
    if (!drawing) throw new Error('Drawing edit target is missing')
    const shapes = xmlSegments(xml, 'wps:wsp', drawing.start, drawing.end)
    const shape = old.shapeId
      ? shapes.find(
          (s) =>
            /<wps:cNvPr\b[^>]*\bid="([^"]+)"/.exec(xml.slice(s.start, s.end))?.[1] === old.shapeId,
        )
      : shapes[source.shapeIndex]
    if (!shape) throw new Error('Drawing child edit target is missing')
    let next = xml.slice(shape.start, shape.end)
    next = next.replace(/<wps:spPr(?:\s[^>]*)?>[\s\S]*?<\/wps:spPr>/, (spPr) =>
      spPr.replace(/<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/, (xfrm) => {
        let out = xfrm
        if (width !== undefined || height !== undefined)
          out = out.replace(/<a:ext\b[^>]*>/, (tag) => {
            let value = tag
            if (width !== undefined)
              value = setAttr(
                value,
                'cx',
                Math.max(0, Math.round((width * EMU_PER_PX) / source.scaleX)),
              )
            if (height !== undefined)
              value = setAttr(
                value,
                'cy',
                Math.max(0, Math.round((height * EMU_PER_PX) / source.scaleY)),
              )
            return value
          })
        if (dx || dy)
          out = out.replace(/<a:off\b[^>]*>/, (tag) =>
            setAttr(
              setAttr(tag, 'x', Math.round(numAttr(tag, 'x') + dx / source.scaleX)),
              'y',
              Math.round(numAttr(tag, 'y') + dy / source.scaleY),
            ),
          )
        return out
      }),
    )
    next = patchShapeProperties(next, style)
    if (height !== undefined)
      next = next.replace(/<a:spAutoFit\s*\/>|<a:spAutoFit\s*>\s*<\/a:spAutoFit>/, '<a:noAutofit/>')
    edits.push({ ...shape, xml: next })
    // An ungrouped shape owns the outer drawing extent as well.
    if (
      source.scaleX === 1 &&
      source.scaleY === 1 &&
      !xml.slice(drawing.start, drawing.end).includes('<wpg:')
    ) {
      const extent = /<wp:extent\b[^>]*>/.exec(xml.slice(drawing.start, drawing.end))
      if (extent && (width !== undefined || height !== undefined)) {
        let tag = extent[0]
        if (width !== undefined) tag = setAttr(tag, 'cx', Math.round(width * EMU_PER_PX))
        if (height !== undefined) tag = setAttr(tag, 'cy', Math.round(height * EMU_PER_PX))
        edits.push({
          start: drawing.start + extent.index,
          end: drawing.start + extent.index + extent[0].length,
          xml: tag,
        })
      }
    }
    const fallback = fallbackShape(xml, source.name, drawing)
    if (fallback)
      edits.push({
        start: fallback.start,
        end: fallback.end,
        xml: patchFallback(
          xml.slice(fallback.start, fallback.end),
          fallback.scale,
          width,
          height,
          dx,
          dy,
          style,
        ),
      })
  })
  for (const edit of edits.sort((a, b) => b.start - a.start))
    xml = xml.slice(0, edit.start) + edit.xml + xml.slice(edit.end)
  return xml
}

function fallbackShape(
  xml: string,
  name: string | undefined,
  drawing: { start: number; end: number },
): { start: number; end: number; scale: VmlGroupScale | null } | null {
  if (!name || !xml.includes('<mc:Fallback')) return null
  const alternative = xmlSegments(xml, 'mc:AlternateContent', 0, xml.length).find(
    (s) => s.start < drawing.start && s.end > drawing.end,
  )
  if (!alternative) return null
  let scale: VmlGroupScale | null = null
  let found = false
  const walk = (nodes: XNode[], parent: VmlGroupScale | null) => {
    for (const node of nodes) {
      const kind = nameOf(node)
      const current = kind === 'v:group' ? vmlGroupScale(node, parent) : parent
      if (
        attrsOf(node)['id'] === name &&
        ['v:shape', 'v:line', 'v:rect', 'v:oval', 'v:roundrect'].includes(kind ?? '')
      ) {
        scale = current
        found = true
        return
      }
      walk(childrenOf(node), current)
      if (found) return
    }
  }
  for (const fallback of xmlSegments(xml, 'mc:Fallback', alternative.start, alternative.end)) {
    walk(xmlParser.parse(xml.slice(fallback.start, fallback.end)) as XNode[], null)
    if (!found) continue
    for (const kind of ['v:shape', 'v:line', 'v:rect', 'v:oval', 'v:roundrect']) {
      const seg = xmlSegments(xml, kind, fallback.start, fallback.end).find(
        (s) =>
          /\bid="([^"]+)"/.exec(xml.slice(s.start, xml.indexOf('>', s.start) + 1))?.[1] === name,
      )
      if (seg) return { ...seg, scale }
    }
  }
  return null
}

function patchFallback(
  xml: string,
  scale: VmlGroupScale | null,
  width: number | undefined,
  height: number | undefined,
  dx: number,
  dy: number,
  patch: ShapeStylePatch,
): string {
  return xml.replace(/^<v:[^>]+>/, (open) => {
    let tag = open
    const sx = scale?.sx ?? 96 / 72
    const sy = scale?.sy ?? 96 / 72
    const style = /\bstyle="([^"]*)"/.exec(tag)?.[1] ?? ''
    let next = style
    const prop = (key: string, value: number) => {
      const text = `${Math.round(value * 1000) / 1000}${scale ? '' : 'pt'}`
      const re = new RegExp(`(^|;)${key}:[^;]*`)
      next = re.test(next) ? next.replace(re, `$1${key}:${text}`) : `${next};${key}:${text}`
    }
    if (tag.startsWith('<v:line')) {
      const from = /\bfrom="(-?[\d.]+),(-?[\d.]+)"/.exec(tag)
      const to = /\bto="(-?[\d.]+),(-?[\d.]+)"/.exec(tag)
      if (from && to) {
        const x = Number(from[1]) + dx / EMU_PER_PX / sx
        const y = Number(from[2]) + dy / EMU_PER_PX / sy
        const ex =
          width === undefined
            ? Number(to[1]) - Number(from[1])
            : ((Number(to[1]) < Number(from[1]) ? -1 : 1) * width) / sx
        const ey =
          height === undefined
            ? Number(to[2]) - Number(from[2])
            : ((Number(to[2]) < Number(from[2]) ? -1 : 1) * height) / sy
        tag = setAttr(setAttr(tag, 'from', `${x},${y}`), 'to', `${x + ex},${y + ey}`)
      }
    } else {
      if (width !== undefined) prop('width', width / sx)
      if (height !== undefined) prop('height', height / sy)
      for (const [key, delta, unit] of [
        [scale ? 'left' : 'margin-left', dx, sx],
        [scale ? 'top' : 'margin-top', dy, sy],
      ] as const) {
        if (delta)
          prop(
            key,
            Number(new RegExp(`(?:^|;)${key}:(-?[\\d.]+)`).exec(style)?.[1] ?? 0) +
              delta / EMU_PER_PX / unit,
          )
      }
    }
    if (patch.rotDeg !== undefined) {
      next = next.replace(/(?:^|;)rotation:[^;]*/g, '')
      if (patch.rotDeg) next += `;rotation:${patch.rotDeg}`
    }
    if (next !== style) tag = setAttr(tag, 'style', next)
    if (patch.fillHex !== undefined)
      tag = setAttr(
        setAttr(tag, 'filled', patch.fillHex ? 't' : 'f'),
        'fillcolor',
        `#${patch.fillHex ?? 'FFFFFF'}`,
      )
    if (patch.borderHex !== undefined)
      tag = setAttr(
        setAttr(tag, 'stroked', patch.borderHex ? 't' : 'f'),
        'strokecolor',
        `#${patch.borderHex ?? '000000'}`,
      )
    if (patch.borderWidthPx !== undefined)
      tag = setAttr(tag, 'strokeweight', `${((patch.borderWidthPx ?? 1) * 72) / 96}pt`)
    return tag
  })
}
