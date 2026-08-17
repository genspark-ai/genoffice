/**
 * Static node rendering — draws a RenderNode as Konva shapes (no interaction logic).
 *
 * Reused in three places so canvas / thumbnails / group children / master decoration
 * layer stay visually identical:
 * - SlideCanvas: shape content inside the interactive Group + group children + decoration layer
 * - SlideThumb: whole page statically scaled down
 */
import React from 'react'
import { Group, Rect, Ellipse, Text, Line, Arrow, Image as KImage, Path } from 'react-konva'
import type {
  RenderNode,
  ShapeRenderNode,
  PictureRenderNode,
  ChipRenderNode,
  TableRenderNode,
  ChartRenderNode,
  GroupRenderNode,
  ArrowEndRender,
} from '@genoffice/pptx-render'
import {
  featheredImage,
  fillToKonva,
  processedImage,
  processedImageKey,
  flatColorImage,
  isDegenerateImage,
  strokeToKonva,
  shadowToKonva,
  cropToKonva,
  presetToShapeKind,
  shapeGlyphs,
  layoutGlyphs,
  normalizeColor,
  boxPivotProps,
  centerFillProps,
} from './konva-adapter'
import { ChartBody } from './ChartBody'
import { needsTextFrameHitArea } from './text-hit-area'

export interface NodeBodyProps {
  node: RenderNode
  images: Map<string, HTMLImageElement>
  /** Currently edited in the DOM overlay: skip drawing the shape's text (avoids ghosting under the edit overlay) */
  hideText?: boolean
  /** Table cell being edited (model coordinates): skip drawing that cell's text */
  hideCellText?: { row: number; col: number }
}

/**
 * Node content (local coordinates relative to the node box top-left); positioning/rotation/flip
 * are handled by the outer container.
 */
export const NodeBody = React.memo(function NodeBody({
  node,
  images,
  hideText,
  hideCellText,
}: NodeBodyProps) {
  const { box } = node

  if (node.type === 'group') {
    const g = node as GroupRenderNode
    // ext/chExt scaling is already baked into children geometry (build-slide); no whole-group stretch here
    return (
      <Group>
        {g.children.map((c) => (
          <StaticNode key={c.id} node={c} images={images} />
        ))}
      </Group>
    )
  }

  if (node.type === 'picture') {
    const pic = node as PictureRenderNode
    const rawImg = pic.dataUrl ? images.get(pic.dataUrl) : undefined
    // clrChange/duotone recolor the pixels; the derived key keeps processed variants out of raw cache slots
    const srcKey = processedImageKey(pic.dataUrl ?? '', pic.clrChange, pic.duotone)
    const procImg =
      rawImg && (pic.clrChange || pic.duotone)
        ? (processedImage(
            rawImg,
            pic.dataUrl ?? '',
            pic.clrChange,
            pic.duotone,
          ) as HTMLImageElement)
        : rawImg
    // ≤2×2 pictures stretch to one flat color in PowerPoint (crop is meaningless there)
    const tiny = !!procImg && isDegenerateImage(procImg)
    const img = procImg && tiny ? (flatColorImage(procImg, srcKey) as HTMLImageElement) : procImg
    const clip = pic.clip
    const image = img ? (
      <KImage
        image={
          pic.softEdgePx && img.width
            ? featheredImage(img, srcKey, pic.softEdgePx * (img.width / Math.max(box.w, 1)))
            : img
        }
        width={box.w}
        height={box.h}
        {...(tiny ? {} : cropToKonva(pic, img))}
        {...(pic.opacity != null ? { opacity: pic.opacity } : {})}
        {...(clip ? {} : { ...strokeToKonva(pic.stroke), ...shadowToKonva(pic.shadow, pic.glow) })}
      />
    ) : (
      <Rect width={box.w} height={box.h} fill="#eef" stroke="#99f" dash={[4, 4]} />
    )
    if (clip && img) {
      // picture styles shape clip: the image is clipped into the geometry, stroke follows the geometry outline;
      // shadow/glow is cast by an opaque backing shape (the image exactly covers it, so no color shows through)
      const shadowProps = shadowToKonva(pic.shadow, pic.glow)
      const strokeProps = strokeToKonva(pic.stroke)
      const outline = (extra: Record<string, unknown>) =>
        clip.pathData ? (
          <Path data={clip.pathData} {...extra} />
        ) : clip.polygonPoints ? (
          <Line points={clip.polygonPoints} closed {...extra} />
        ) : (
          <Rect width={box.w} height={box.h} cornerRadius={clip.cornerRadiusPx ?? 0} {...extra} />
        )
      return (
        <>
          {'shadowColor' in shadowProps && outline({ fill: '#ffffff', ...shadowProps })}
          <Group
            clipFunc={(ctx) => {
              if (clip.pathData) return [new Path2D(clip.pathData)] as [Path2D]
              if (clip.polygonPoints) {
                const pts = clip.polygonPoints
                ctx.moveTo(pts[0]!, pts[1]!)
                for (let i = 2; i + 1 < pts.length; i += 2) ctx.lineTo(pts[i]!, pts[i + 1]!)
                ctx.closePath()
                return
              }
              const r = Math.min(clip.cornerRadiusPx ?? 0, box.w / 2, box.h / 2)
              ctx.moveTo(r, 0)
              ctx.arcTo(box.w, 0, box.w, box.h, r)
              ctx.arcTo(box.w, box.h, 0, box.h, r)
              ctx.arcTo(0, box.h, 0, 0, r)
              ctx.arcTo(0, 0, box.w, 0, r)
              ctx.closePath()
            }}
          >
            {/* spPr fill: PowerPoint always paints it behind the image, even a translucent one */}
            {pic.fill && outline({ ...fillToKonva(pic.fill, box.w, box.h, images) })}
            {/* Backdrop only for fully opaque previews: with partial opacity the two alphas would stack */}
            {pic.bgColor && (pic.opacity ?? 1) >= 1 && outline({ fill: pic.bgColor })}
            {image}
          </Group>
          {'stroke' in strokeProps && outline({ ...strokeProps, fillEnabled: false })}
          {pic.media && <MediaBadge kind={pic.media} w={box.w} h={box.h} />}
        </>
      )
    }
    return (
      <>
        {/* spPr fill: PowerPoint always paints it behind the image, even a translucent one */}
        {pic.fill && (
          <Rect width={box.w} height={box.h} {...fillToKonva(pic.fill, box.w, box.h, images)} />
        )}
        {/* Backdrop only for fully opaque previews: with partial opacity the two alphas would stack */}
        {pic.bgColor && img && (pic.opacity ?? 1) >= 1 && (
          <Rect width={box.w} height={box.h} fill={pic.bgColor} />
        )}
        {image}
        {pic.media && <MediaBadge kind={pic.media} w={box.w} h={box.h} />}
      </>
    )
  }

  if (node.type === 'table') {
    const table = node as TableRenderNode
    const tblW = table.gridX[table.gridX.length - 1] ?? node.box.w
    const tblH = table.gridY[table.gridY.length - 1] ?? node.box.h
    return (
      <>
        {table.bgFill && (
          <Rect
            x={0}
            y={0}
            width={tblW}
            height={tblH}
            {...fillToKonva(table.bgFill, tblW, tblH, images)}
          />
        )}
        {table.cells.map((cell, i) => {
          const cellGlyphs =
            hideCellText && cell.row === hideCellText.row && cell.col === hideCellText.col
              ? []
              : layoutGlyphs(cell.text)
          return (
            <React.Fragment key={i}>
              <Rect
                x={cell.x}
                y={cell.y}
                width={cell.w}
                height={cell.h}
                {...fillToKonva(cell.fill, cell.w, cell.h, images)}
              />
              {cell.borders?.t && (
                <Line
                  points={[cell.x, cell.y, cell.x + cell.w, cell.y]}
                  {...strokeToKonva(cell.borders.t)}
                />
              )}
              {cell.borders?.b && (
                <Line
                  points={[cell.x, cell.y + cell.h, cell.x + cell.w, cell.y + cell.h]}
                  {...strokeToKonva(cell.borders.b)}
                />
              )}
              {cell.borders?.l && (
                <Line
                  points={[cell.x, cell.y, cell.x, cell.y + cell.h]}
                  {...strokeToKonva(cell.borders.l)}
                />
              )}
              {cell.borders?.r && (
                <Line
                  points={[cell.x + cell.w, cell.y, cell.x + cell.w, cell.y + cell.h]}
                  {...strokeToKonva(cell.borders.r)}
                />
              )}
              {/* Highlight backgrounds first, so a run's highlight never covers a neighbor's glyphs */}
              {cellGlyphs.map(
                (g, j) =>
                  g.highlight && (
                    <Rect
                      key={`hl${j}`}
                      x={cell.x + (cell.text?.insets.l ?? 0) + g.highlight.x}
                      y={cell.y + (cell.text?.insets.t ?? 0) + g.highlight.y}
                      width={g.highlight.w}
                      height={g.highlight.h}
                      fill={g.highlight.color}
                      listening={false}
                    />
                  ),
              )}
              {cellGlyphs.map((g, j) => (
                <Text
                  key={j}
                  x={cell.x + (cell.text?.insets.l ?? 0) + g.x}
                  y={cell.y + (cell.text?.insets.t ?? 0) + g.y}
                  text={g.text}
                  fontSize={g.fontSize}
                  fontFamily={g.fontFamily}
                  fontStyle={g.fontStyle}
                  textDecoration={g.textDecoration}
                  rotation={g.rotation ?? 0}
                  letterSpacing={g.letterSpacing ?? 0}
                  fill={g.fill}
                  direction={g.direction ?? 'inherit'}
                  {...(g.stroke
                    ? { stroke: g.stroke, strokeWidth: g.strokeWidth, fillAfterStrokeEnabled: true }
                    : {})}
                  {...(g.shadowEnabled
                    ? {
                        shadowColor: g.shadowColor,
                        shadowBlur: g.shadowBlur,
                        shadowOffsetX: g.shadowOffsetX,
                        shadowOffsetY: g.shadowOffsetY,
                      }
                    : {})}
                />
              ))}
            </React.Fragment>
          )
        })}
      </>
    )
  }

  if (node.type === 'chart') {
    return <ChartBody chart={node as ChartRenderNode} images={images} />
  }

  if (node.type === 'placeholder-chip') {
    const chip = node as ChipRenderNode
    return (
      <>
        <Rect
          width={box.w}
          height={box.h}
          fill="#f5f5f7"
          stroke="#c7c7cc"
          dash={[6, 4]}
          cornerRadius={4}
        />
        <Text
          text={`⧉ ${chip.label}`}
          width={box.w}
          height={box.h}
          align="center"
          verticalAlign="middle"
          fontSize={Math.min(16, box.h / 2)}
          fill="#8e8e93"
        />
      </>
    )
  }

  // shape / text
  const shape = node as ShapeRenderNode
  const fillProps = fillToKonva(shape.fill, box.w, box.h, images)
  const strokeProps = strokeToKonva(shape.stroke)
  const shadowProps = shadowToKonva(shape.shadow, shape.glow)
  const glyphs = hideText ? [] : shapeGlyphs(shape)

  // Connector/straight line: polyline (flip already baked into points), with optional arrow endpoints
  if (shape.line) {
    const color = strokeProps.stroke ?? normalizeColor('#000000')
    const sw = strokeProps.strokeWidth ?? 1
    const { headEnd, tailEnd, bezier } = shape.line

    // Curved connector: draw the bezier with an SVG Path
    if (bezier && bezier.length) {
      const pts = shape.line.points
      const startX = pts[0] ?? 0
      const startY = pts[1] ?? 0
      let d = `M ${startX} ${startY}`
      // Each segment: [cp1x,cp1y,cp2x,cp2y,ex,ey]
      for (let i = 0; i < bezier.length; i += 6) {
        d += ` C ${bezier[i]} ${bezier[i + 1]}, ${bezier[i + 2]} ${bezier[i + 3]}, ${bezier[i + 4]} ${bezier[i + 5]}`
      }
      const dashProps = strokeProps.dash ? { dash: strokeProps.dash } : {}
      return (
        <>
          <Path
            data={d}
            stroke={color}
            strokeWidth={sw}
            hitStrokeWidth={Math.max(sw, 12)}
            fill="transparent"
            lineCap="round"
            lineJoin="round"
            {...dashProps}
            {...shadowProps}
          />
          {headEnd && <ArrowHead end={headEnd} pts={shape.line.points} atStart color={color} />}
          {tailEnd && (
            <ArrowHead end={tailEnd} pts={shape.line.points} atStart={false} color={color} />
          )}
        </>
      )
    }

    // Straight/bent lines (line/straightConnector/bentConnector)
    // Determine whether custom arrowheads are needed (non-triangle types; triangle uses Konva Arrow built-in)
    const hasCustomHead = headEnd && headEnd.type !== 'arrow' && headEnd.type !== 'triangle'
    const hasCustomTail = tailEnd && tailEnd.type !== 'arrow' && tailEnd.type !== 'triangle'
    const useKonvaArrow = (headEnd || tailEnd) && !hasCustomHead && !hasCustomTail

    return (
      <>
        {useKonvaArrow ? (
          <Arrow
            points={shape.line.points}
            stroke={color}
            strokeWidth={sw}
            hitStrokeWidth={Math.max(sw, 12)}
            {...(strokeProps.dash ? { dash: strokeProps.dash } : {})}
            fill={color}
            pointerAtBeginning={!!headEnd}
            pointerAtEnding={!!tailEnd}
            pointerLength={
              tailEnd
                ? Math.max(tailEnd.lengthPx, 6)
                : headEnd
                  ? Math.max(headEnd.lengthPx, 6)
                  : sw * 3.5
            }
            pointerWidth={
              tailEnd
                ? Math.max(tailEnd.widthPx, 5)
                : headEnd
                  ? Math.max(headEnd.widthPx, 5)
                  : sw * 3
            }
            lineCap="round"
            lineJoin="round"
            {...shadowProps}
          />
        ) : (
          <Line
            points={shape.line.points}
            stroke={color}
            strokeWidth={sw}
            hitStrokeWidth={Math.max(sw, 12)}
            {...(strokeProps.dash ? { dash: strokeProps.dash } : {})}
            lineCap="round"
            lineJoin="round"
            {...shadowProps}
          />
        )}
        {hasCustomHead && headEnd && (
          <ArrowHead end={headEnd} pts={shape.line.points} atStart color={color} />
        )}
        {hasCustomTail && tailEnd && (
          <ArrowHead end={tailEnd} pts={shape.line.points} atStart={false} color={color} />
        )}
      </>
    )
  }

  let geom: React.ReactNode
  if (shape.pathData || shape.fillPathData || shape.strokePathData) {
    // custGeom / arc-type presets: SVG path (local px; flip/rot handled by the outer container)
    geom = (
      <>
        {shape.fillPathData && <Path data={shape.fillPathData} {...fillProps} {...shadowProps} />}
        {shape.pathData && (
          <Path
            data={shape.pathData}
            {...fillProps}
            {...strokeProps}
            {...shadowProps}
            lineJoin="round"
          />
        )}
        {shape.strokePathData && (
          <Path data={shape.strokePathData} {...strokeProps} fillEnabled={false} lineJoin="round" />
        )}
      </>
    )
  } else if (shape.polygonPoints) {
    geom = (
      <Line
        points={shape.polygonPoints}
        closed
        {...fillProps}
        {...strokeProps}
        {...shadowProps}
        lineJoin="round"
      />
    )
  } else if (presetToShapeKind(shape.presetGeometry) === 'ellipse') {
    geom = (
      <Ellipse
        x={box.w / 2}
        y={box.h / 2}
        radiusX={box.w / 2}
        radiusY={box.h / 2}
        {...centerFillProps(fillProps, box.w, box.h)}
        {...strokeProps}
        {...shadowProps}
      />
    )
  } else {
    const rounded =
      presetToShapeKind(shape.presetGeometry) === 'roundRect' || shape.cornerRadiusPx != null
    geom = (
      <Rect
        width={box.w}
        height={box.h}
        cornerRadius={rounded ? (shape.cornerRadiusPx ?? Math.min(box.w, box.h) * 0.167) : 0}
        {...fillProps}
        {...strokeProps}
        {...shadowProps}
      />
    )
  }

  // fillOverlay: PowerPoint blends the overlay against the shape's own fill in isolation.
  // Approximation: an opaque white underlay + base fill + overlay with multiply — where the
  // base fill is absent/translucent the multiply hits white (= plain overlay color) instead
  // of bleeding in whatever lies under the shape.
  let overlayUnder: React.ReactNode = null
  let overlayGeom: React.ReactNode = null
  if (shape.fillOverlay) {
    const oProps = fillToKonva(shape.fillOverlay, box.w, box.h, images)
    const sameGeom = (props: Record<string, unknown>): React.ReactNode => {
      const common = { listening: false, ...props }
      if (shape.fillPathData || shape.pathData)
        return <Path data={(shape.fillPathData ?? shape.pathData)!} {...common} />
      if (shape.polygonPoints) return <Line points={shape.polygonPoints} closed {...common} />
      if (presetToShapeKind(shape.presetGeometry) === 'ellipse')
        return (
          <Ellipse
            x={box.w / 2}
            y={box.h / 2}
            radiusX={box.w / 2}
            radiusY={box.h / 2}
            {...common}
          />
        )
      return (
        <Rect width={box.w} height={box.h} cornerRadius={shape.cornerRadiusPx ?? 0} {...common} />
      )
    }
    overlayUnder = sameGeom({ fill: '#ffffff' })
    const isEllipse =
      !shape.fillPathData &&
      !shape.pathData &&
      !shape.polygonPoints &&
      presetToShapeKind(shape.presetGeometry) === 'ellipse'
    overlayGeom = sameGeom({
      ...(isEllipse ? centerFillProps(oProps, box.w, box.h) : oProps),
      globalCompositeOperation: 'multiply' as const,
    })
  }

  return (
    <>
      {/* Text is drawn as individual glyph runs, so line spacing and insets otherwise
          have no hit area. Cover every text-bearing shape (including round/custom
          geometry) so clicks inside its text frame cannot reach a picture underneath. */}
      {needsTextFrameHitArea(shape) && <Rect width={box.w} height={box.h} fill="transparent" />}
      {overlayUnder}
      {geom}
      {overlayGeom}
      {/* Text highlight backgrounds: all rects first so a run's highlight never covers a neighbor's glyphs */}
      {glyphs.map(
        (g, i) =>
          g.highlight && (
            <Rect
              key={`hl${i}`}
              x={(shape.text?.insets.l ?? 0) + g.highlight.x}
              y={(shape.text?.insets.t ?? 0) + g.highlight.y}
              width={g.highlight.w}
              height={g.highlight.h}
              fill={g.highlight.color}
              listening={false}
            />
          ),
      )}
      {glyphs.map((g, i) => (
        <Text
          key={i}
          x={(shape.text?.insets.l ?? 0) + g.x}
          y={(shape.text?.insets.t ?? 0) + g.y}
          text={g.text}
          fontSize={g.fontSize}
          fontFamily={g.fontFamily}
          fontStyle={g.fontStyle}
          textDecoration={g.textDecoration}
          rotation={g.rotation ?? 0}
          letterSpacing={g.letterSpacing ?? 0}
          fill={g.fill}
          direction={g.direction ?? 'inherit'}
          {...(g.stroke
            ? { stroke: g.stroke, strokeWidth: g.strokeWidth, fillAfterStrokeEnabled: true }
            : {})}
          {...(g.shadowEnabled
            ? {
                shadowColor: g.shadowColor,
                shadowBlur: g.shadowBlur,
                shadowOffsetX: g.shadowOffsetX,
                shadowOffsetY: g.shadowOffsetY,
              }
            : {})}
        />
      ))}
    </>
  )
})

/**
 * ArrowHead — draws a custom arrowhead (stealth/diamond/oval/arrow) at a polyline's start or end.
 * The arrowhead always faces the segment direction (angle taken from the tangent at the end point).
 */
function ArrowHead({
  end,
  pts,
  atStart,
  color,
}: {
  end: ArrowEndRender
  pts: number[]
  atStart: boolean
  color: string
}) {
  const nPts = pts.length / 2
  if (nPts < 2) return null

  // Compute the arrowhead angle
  let angle: number
  let tipX: number
  let tipY: number
  if (atStart) {
    // Looking from point 2 toward point 1 (arrow points at pts[0,1])
    const dx = pts[0]! - pts[2]!
    const dy = pts[1]! - pts[3]!
    angle = Math.atan2(dy, dx) * (180 / Math.PI)
    tipX = pts[0]!
    tipY = pts[1]!
  } else {
    // Looking from the second-to-last point toward the last
    const lx = pts[(nPts - 1) * 2]!
    const ly = pts[(nPts - 1) * 2 + 1]!
    const px = pts[(nPts - 2) * 2]!
    const py = pts[(nPts - 2) * 2 + 1]!
    const dx = lx - px
    const dy = ly - py
    angle = Math.atan2(dy, dx) * (180 / Math.PI)
    tipX = lx
    tipY = ly
  }

  const w = end.widthPx
  const l = end.lengthPx
  const halfW = w / 2

  if (end.type === 'oval') {
    return (
      <Ellipse
        x={tipX}
        y={tipY}
        radiusX={l / 2}
        radiusY={halfW}
        rotation={angle}
        fill={color}
        listening={false}
      />
    )
  }

  if (end.type === 'diamond') {
    // Diamond: tip as the front vertex, extending backward by l
    const pts2 = [0, 0, -l / 2, -halfW, -l, 0, -l / 2, halfW]
    return (
      <Line
        points={pts2}
        closed
        fill={color}
        stroke={color}
        strokeWidth={0}
        x={tipX}
        y={tipY}
        rotation={angle}
        listening={false}
      />
    )
  }

  if (end.type === 'stealth') {
    // Stealth arrowhead: concave spear shape
    const pts2 = [0, 0, -l, -halfW, -l * 0.65, 0, -l, halfW]
    return (
      <Line
        points={pts2}
        closed
        fill={color}
        stroke={color}
        strokeWidth={0}
        x={tipX}
        y={tipY}
        rotation={angle}
        listening={false}
      />
    )
  }

  // 'arrow': open V-shaped line arrowhead
  const pts2 = [-l, -halfW, 0, 0, -l, halfW]
  return (
    <Line
      points={pts2}
      stroke={color}
      strokeWidth={Math.max(1, w * 0.12)}
      x={tipX}
      y={tipY}
      rotation={angle}
      listening={false}
      lineCap="round"
      lineJoin="round"
    />
  )
}

/** Play/speaker badge on audio/video poster frames (centered translucent circle + white glyph). */
function MediaBadge({ kind, w, h }: { kind: 'video' | 'audio'; w: number; h: number }) {
  const r = Math.max(10, Math.min(32, Math.min(w, h) * 0.18))
  const cx = w / 2
  const cy = h / 2
  return (
    <Group listening={false}>
      <Ellipse x={cx} y={cy} radiusX={r} radiusY={r} fill="rgba(0,0,0,0.55)" />
      {kind === 'video' ? (
        <Line
          points={[cx - r * 0.3, cy - r * 0.5, cx - r * 0.3, cy + r * 0.5, cx + r * 0.55, cy]}
          closed
          fill="#ffffff"
          lineJoin="round"
        />
      ) : (
        <Text
          text="♪"
          x={cx - r}
          y={cy - r}
          width={r * 2}
          height={r * 2}
          align="center"
          verticalAlign="middle"
          fontSize={r * 1.2}
          fill="#ffffff"
        />
      )}
    </Group>
  )
}

/**
 * Statically positioned node: container (position/rotation/flip) + NodeBody.
 * Rotation and flip pivot on the box center (boxPivotProps), matching OOXML.
 */
export const StaticNode = React.memo(function StaticNode({ node, images }: NodeBodyProps) {
  const { box } = node
  return (
    <Group {...boxPivotProps(box)} listening={false}>
      <NodeBody node={node} images={images} />
    </Group>
  )
})
