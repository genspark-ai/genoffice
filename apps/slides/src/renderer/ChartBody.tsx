/**
 * Chart drawing primitives (ChartRenderNode → Konva shapes).
 * Shared by SlideCanvas and SlideThumb; coordinates are already relative to the chart box
 * top-left, the outer Group handles positioning.
 */
import React from 'react'
import { Rect, Text, Line, Circle, Arc, Path } from 'react-konva'
import type { ChartRenderNode } from '@genoffice/pptx-render'
import { fillToKonva, smoothTension } from './konva-adapter'

export function ChartBody({
  chart,
  images,
}: {
  chart: ChartRenderNode
  images?: Map<string, HTMLImageElement>
}) {
  return (
    <>
      {chart.bgFill && (
        <Rect
          x={0}
          y={0}
          width={chart.box.w}
          height={chart.box.h}
          {...fillToKonva(chart.bgFill, chart.box.w, chart.box.h, images)}
        />
      )}
      {chart.plotRect && (
        <Rect
          x={chart.plotRect.x}
          y={chart.plotRect.y}
          width={chart.plotRect.w}
          height={chart.plotRect.h}
          {...(chart.plotRect.fill
            ? fillToKonva(chart.plotRect.fill, chart.plotRect.w, chart.plotRect.h, images)
            : {})}
          {...(chart.plotRect.borderColor
            ? {
                stroke: chart.plotRect.borderColor,
                strokeWidth: chart.plotRect.borderWidthPx ?? 1,
              }
            : {})}
        />
      )}
      {(chart.wedges ?? []).map((wd, i) => (
        <Arc
          key={`w${i}`}
          x={wd.cx}
          y={wd.cy}
          innerRadius={wd.innerR}
          outerRadius={wd.outerR}
          angle={wd.sweepDeg}
          rotation={wd.startDeg}
          fill={wd.color}
          stroke="#ffffff"
          strokeWidth={1}
        />
      ))}
      {chart.gridLines.map((g, i) => (
        <Line
          key={`g${i}`}
          points={[g.x1, g.y1, g.x2, g.y2]}
          stroke={g.color}
          strokeWidth={g.widthPx ?? 1}
          {...(g.dash ? { dash: g.dash } : {})}
        />
      ))}
      {chart.axisLines.map((a, i) => (
        <Line
          key={`a${i}`}
          points={[a.x1, a.y1, a.x2, a.y2]}
          stroke={a.color}
          strokeWidth={a.widthPx}
        />
      ))}
      {(chart.paths ?? []).map((p, i) => (
        <Path
          key={`pp${i}`}
          data={p.d}
          y={p.dy ?? 0}
          fill={p.fill}
          {...(p.stroke ? { stroke: p.stroke, strokeWidth: 1 } : {})}
        />
      ))}
      {chart.bars.map((b, i) => (
        <Rect key={`b${i}`} x={b.x} y={b.y} width={b.w} height={b.h} fill={b.color} />
      ))}
      {chart.polylines.map((p, i) => (
        <Line
          key={`p${i}`}
          points={p.points}
          stroke={p.color}
          strokeWidth={p.widthPx}
          lineCap="round"
          lineJoin="round"
          tension={smoothTension(p.smooth)}
          {...(p.closed ? { closed: true } : {})}
          {...(p.fill ? { fill: p.fill } : {})}
          {...(p.dash ? { dash: p.dash } : {})}
        />
      ))}
      {chart.markers.map((m, i) => (
        <Circle key={`m${i}`} x={m.x} y={m.y} radius={m.r} fill={m.color} />
      ))}
      {chart.swatches.map((s, i) => (
        <Rect
          key={`s${i}`}
          x={s.x}
          y={s.y}
          width={s.w}
          height={s.h}
          fill={s.color}
          cornerRadius={1}
        />
      ))}
      {chart.labels.map((l, i) => (
        <Text
          key={`t${i}`}
          x={l.x}
          y={l.y}
          text={l.text}
          fontSize={l.fontSizePx}
          fontFamily="Arial"
          fill={l.color}
          fontStyle={l.bold ? 'bold' : 'normal'}
          {...(l.rotationDeg ? { rotation: l.rotationDeg } : {})}
        />
      ))}
      {chart.border && (
        <Rect
          x={chart.border.widthPx / 2}
          y={chart.border.widthPx / 2}
          width={chart.box.w - chart.border.widthPx}
          height={chart.box.h - chart.border.widthPx}
          stroke={chart.border.color}
          strokeWidth={chart.border.widthPx}
        />
      )}
    </>
  )
}
