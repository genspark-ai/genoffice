/**
 * Show tween for the Morph transition — a Konva approximation of PowerPoint "Morph".
 *
 * Pairing rules: elements of the previous and target pages match by stable keys — cNvPr id first
 * (same-id elements produced by duplicating a page), then cNvPr name (same-named shapes); each key
 * is consumed one-to-one. Paired elements tween position/size/rotation (around the target box
 * center), new elements on the target page fade in, elements unique to the previous page fade out.
 * The decoration layer (master/layout) renders statically per the target page — in the common
 * same-layout case it's identical anyway; with different layouts we accept an instant switch.
 */
import React, { useEffect, useMemo, useRef, useState } from 'react'
import { Stage, Layer, Rect, Group } from 'react-konva'
import type { RenderNode, RenderSlide } from '@genoffice/pptx-render'
import type { ShapeKey } from '../../shared/ipc'
import { fillToKonva } from '../konva-adapter'
import { StaticNode } from '../NodeBody'

/** Tween duration (ms); PowerPoint Morph defaults to longer, we use a snappier 500ms. */
const MORPH_MS = 500

function easeInOut(t: number): number {
  return t < 0.5 ? 2 * t * t : 1 - ((2 - 2 * t) * (2 - 2 * t)) / 2
}

function lerp(a: number, b: number, t: number): number {
  return a + (b - a) * t
}

/** Tween source for each content node on the target page (no from = new, fades in). */
interface MorphPlan {
  /** Target node sourceId → paired node on the previous page */
  fromOf: Map<string, RenderNode>
  /** Nodes unique to the previous page (fade out) */
  leaving: RenderNode[]
}

function buildPlan(
  from: RenderSlide,
  to: RenderSlide,
  fromKeys: ShapeKey[],
  toKeys: ShapeKey[],
): MorphPlan {
  const keyOf = (keys: ShapeKey[]) => new Map(keys.map((k) => [k.sourceId, k]))
  const fk = keyOf(fromKeys)
  const tk = keyOf(toKeys)
  // Previous page content node index: one table each for cNvPr id and name (same keys queue up, first come first paired)
  const bySpid = new Map<number, RenderNode[]>()
  const byName = new Map<string, RenderNode[]>()
  for (const n of from.nodes) {
    if (n.decoration) continue
    const k = fk.get(n.sourceId)
    if (!k) continue
    if (k.spid != null) {
      const arr = bySpid.get(k.spid) ?? []
      arr.push(n)
      bySpid.set(k.spid, arr)
    }
    if (k.name) {
      const arr = byName.get(k.name) ?? []
      arr.push(n)
      byName.set(k.name, arr)
    }
  }
  const used = new Set<RenderNode>()
  const take = (arr: RenderNode[] | undefined): RenderNode | null => {
    for (const n of arr ?? []) {
      if (!used.has(n)) {
        used.add(n)
        return n
      }
    }
    return null
  }
  const fromOf = new Map<string, RenderNode>()
  for (const n of to.nodes) {
    if (n.decoration) continue
    const k = tk.get(n.sourceId)
    if (!k) continue
    const match =
      (k.spid != null ? take(bySpid.get(k.spid)) : null) ??
      (k.name ? take(byName.get(k.name)) : null)
    if (match) fromOf.set(n.sourceId, match)
  }
  const leaving = from.nodes.filter((n) => !n.decoration && !used.has(n))
  return { fromOf, leaving }
}

export function MorphStage({
  from,
  to,
  fromKeys,
  toKeys,
  images,
  width,
  onDone,
}: {
  from: RenderSlide
  to: RenderSlide
  fromKeys: ShapeKey[]
  toKeys: ShapeKey[]
  images: Map<string, HTMLImageElement>
  width: number
  /** Tween finished (guaranteed to fire only once) */
  onDone: () => void
}) {
  const [t, setT] = useState(0)
  const onDoneRef = useRef(onDone)
  onDoneRef.current = onDone

  useEffect(() => {
    let raf = 0
    const t0 = performance.now()
    const tick = () => {
      const u = (performance.now() - t0) / MORPH_MS
      if (u >= 1) {
        setT(1)
        onDoneRef.current()
      } else {
        setT(u)
        raf = requestAnimationFrame(tick)
      }
    }
    raf = requestAnimationFrame(tick)
    return () => cancelAnimationFrame(raf)
  }, [])

  const plan = useMemo(() => buildPlan(from, to, fromKeys, toKeys), [from, to, fromKeys, toKeys])

  const scale = width / to.widthPx
  const h = to.heightPx * scale
  const bg = fillToKonva(to.background, to.widthPx, to.heightPx, images)
  const q = easeInOut(Math.min(1, Math.max(0, t)))

  return (
    <Stage width={width} height={h} listening={false} style={{ pointerEvents: 'none' }}>
      <Layer scaleX={scale} scaleY={scale} listening={false}>
        <Rect
          x={0}
          y={0}
          width={to.widthPx}
          height={to.heightPx}
          {...(Object.keys(bg).length ? bg : { fill: '#ffffff' })}
        />
        {/* Elements unique to the previous page: fade out in place */}
        {plan.leaving.map((n) => (
          <Group key={`out-${n.id}`} opacity={1 - q} listening={false}>
            <StaticNode node={n} images={images} />
          </Group>
        ))}
        {/* Target page elements: paired → geometry tween; new → fade in; decoration → static */}
        {to.nodes.map((n) => {
          if (n.decoration) return <StaticNode key={n.id} node={n} images={images} />
          const f = plan.fromOf.get(n.sourceId)
          if (!f) {
            return (
              <Group key={n.id} opacity={q} listening={false}>
                <StaticNode node={n} images={images} />
              </Group>
            )
          }
          // Tween around the target box center: position along the line between centers, size by width/height ratios, rotation by angle difference
          const b0 = f.box
          const b1 = n.box
          return (
            <Group
              key={n.id}
              x={lerp(b0.centerX, b1.centerX, q)}
              y={lerp(b0.centerY, b1.centerY, q)}
              offsetX={b1.centerX}
              offsetY={b1.centerY}
              scaleX={lerp(b0.w / Math.max(1, b1.w), 1, q)}
              scaleY={lerp(b0.h / Math.max(1, b1.h), 1, q)}
              rotation={lerp(b0.rotationDeg - b1.rotationDeg, 0, q)}
              listening={false}
            >
              <StaticNode node={n} images={images} />
            </Group>
          )
        })}
      </Layer>
    </Stage>
  )
}
