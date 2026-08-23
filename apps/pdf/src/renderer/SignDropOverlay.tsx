import { useState } from 'react'
import type { ReactElement } from 'react'
import { cssRgb } from './DrawLayer'
import type { SignatureData } from './SignatureDialog'

/** Scale factor applied when placing a signature: 1/3 of the displayed page width,
    capped at 1/6 of its height so tall images stay signature-sized */
export const signPlaceK = (sig: SignatureData, dispW: number, dispH: number): number =>
  Math.min(dispW / 3 / sig.width, dispH / 6 / sig.height)

/** Inserted images land at up to half the page, never above natural size
    (0.75 ≈ px→pt, so a screen-resolution image keeps its printed size) */
export const imagePlaceK = (sig: SignatureData, dispW: number, dispH: number): number =>
  Math.min(dispW / 2 / sig.width, dispH / 2 / sig.height, 0.75)
export const staticFormFillPlaceK = (): number => 1

/** Click-to-place overlay: a translucent ghost of the pending signature follows the
    cursor at its actual landing size, and clicking drops it centered on that point */
export function SignDropOverlay({
  sig,
  dispW,
  dispH,
  scale,
  color,
  title,
  onPlace,
  placeK = signPlaceK,
}: {
  sig: SignatureData
  /** Displayed page size at scale=1 (view coords) */
  dispW: number
  dispH: number
  scale: number
  color: [number, number, number]
  title: string
  onPlace: (vx: number, vy: number) => void
  /** Landing-size rule; defaults to signature sizing (image insert passes its own) */
  placeK?: (sig: SignatureData, dispW: number, dispH: number) => number
}): ReactElement {
  const [pt, setPt] = useState<[number, number] | null>(null)
  const k = placeK(sig, dispW, dispH) * scale
  const w = sig.width * k
  const h = sig.height * k
  return (
    <div
      className="pdf-sign-drop"
      data-tip={title}
      onPointerMove={(e) => {
        const box = e.currentTarget.getBoundingClientRect()
        setPt([e.clientX - box.left, e.clientY - box.top])
      }}
      onPointerLeave={() => setPt(null)}
      onClick={(e) => {
        const box = e.currentTarget.getBoundingClientRect()
        onPlace((e.clientX - box.left) / scale, (e.clientY - box.top) / scale)
      }}
    >
      {pt && (
        <div
          className="pdf-sign-ghost"
          style={{
            left: Math.min(Math.max(pt[0] - w / 2, 0), Math.max(dispW * scale - w, 0)),
            top: Math.min(Math.max(pt[1] - h / 2, 0), Math.max(dispH * scale - h, 0)),
            width: w,
            height: h,
          }}
        >
          {sig.kind === 'image' ? (
            <img src={`data:image/png;base64,${sig.image}`} alt="" draggable={false} />
          ) : (
            <svg viewBox={`0 0 ${sig.width} ${sig.height}`} preserveAspectRatio="none">
              {sig.paths.map((p, i) => {
                const pts: string[] = []
                for (let j = 0; j < p.length; j += 2) pts.push(`${p[j]},${p[j + 1]}`)
                return (
                  <polyline
                    key={i}
                    points={pts.join(' ')}
                    fill="none"
                    stroke={cssRgb(color)}
                    strokeWidth={2.4}
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                )
              })}
            </svg>
          )}
        </div>
      )}
    </div>
  )
}
