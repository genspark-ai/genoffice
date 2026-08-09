import { useEffect, useMemo, useRef, useState } from 'react'
import type { ChangeEvent, PointerEvent as ReactPointerEvent, ReactElement } from 'react'
import { cssRgb } from './DrawLayer'
import type { TFunc } from './i18n/locale'

const PAD_W = 420
const PAD_H = 150
/** Uploaded signature images are downscaled to this max dimension to bound file size */
const MAX_IMG = 1600

/** Typed-signature font choices; families fall back across macOS/Windows system fonts.
    Each card previews the typed name in its font, so no localized labels are needed. */
const SIGN_FONTS: { id: string; family: string; italic?: boolean }[] = [
  {
    id: 'script',
    family: `"Snell Roundhand", "Segoe Script", "Brush Script MT", cursive`,
    italic: true,
  },
  { id: 'xingkai', family: `"Xingkai SC", "STXingkai", "KaiTi", "Kaiti SC", cursive` },
  { id: 'kaiti', family: `"Kaiti SC", "STKaiti", "KaiTi", "DFKai-SB", serif` },
  { id: 'songti', family: `"Songti SC", "SimSun", "Times New Roman", serif` },
]

/** Font size used when rasterizing a typed signature; large so the stamp stays crisp when scaled */
const TYPE_FONT_PX = 200

/** Signature strokes: pad pixel coords, scaled proportionally and y-flipped when placed on the page */
export interface SignatureStrokes {
  paths: number[][]
  width: number
  height: number
}

/** Confirmed signature awaiting placement: hand strokes (Ink) or a bitmap (Stamp) */
export type SignatureData =
  | ({ kind: 'strokes' } & SignatureStrokes)
  | {
      kind: 'image'
      /** base64 PNG, without the data: prefix */
      image: string
      width: number
      height: number
    }

/** Decode + downscale an image file onto a canvas (null if it can't be decoded) */
export async function fileToCanvas(
  file: File,
  maxDim: number = MAX_IMG,
): Promise<HTMLCanvasElement | null> {
  const url = URL.createObjectURL(file)
  try {
    const img = new Image()
    await new Promise<void>((resolve, reject) => {
      img.onload = () => resolve()
      img.onerror = () => reject(new Error('decode'))
      img.src = url
    })
    const k = Math.min(1, maxDim / Math.max(img.naturalWidth, img.naturalHeight, 1))
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.round(img.naturalWidth * k))
    canvas.height = Math.max(1, Math.round(img.naturalHeight * k))
    canvas.getContext('2d')?.drawImage(img, 0, 0, canvas.width, canvas.height)
    return canvas
  } catch {
    return null
  } finally {
    URL.revokeObjectURL(url)
  }
}

/** Black & white: dark pixels → solid black, light pixels → transparent (with a soft
    ramp in between for antialiasing). Turns a photographed signature into a clean stamp. */
function toBlackAndWhite(src: HTMLCanvasElement): HTMLCanvasElement {
  const out = document.createElement('canvas')
  out.width = src.width
  out.height = src.height
  const ctx = out.getContext('2d', { willReadFrequently: true })
  const srcCtx = src.getContext('2d', { willReadFrequently: true })
  if (!ctx || !srcCtx) return src
  const img = srcCtx.getImageData(0, 0, src.width, src.height)
  const d = img.data
  for (let i = 0; i < d.length; i += 4) {
    const lum = 0.299 * d[i]! + 0.587 * d[i + 1]! + 0.114 * d[i + 2]!
    const a = lum > 200 ? 0 : lum < 120 ? 255 : Math.round(((200 - lum) / 80) * 255)
    d[i] = 0
    d[i + 1] = 0
    d[i + 2] = 0
    d[i + 3] = Math.round((a * d[i + 3]!) / 255)
  }
  ctx.putImageData(img, 0, 0)
  return out
}

/** Rasterize typed text with the chosen font onto a tightly-cropped transparent canvas.
    The result goes through the image (Stamp) pipeline, so what you see in the font
    preview card is exactly what lands on the page — no stroke approximation. */
function textToImage(
  text: string,
  f: (typeof SIGN_FONTS)[number],
  color: [number, number, number],
): Extract<SignatureData, { kind: 'image' }> | null {
  const canvas = document.createElement('canvas')
  const ctx = canvas.getContext('2d')
  if (!ctx) return null
  const font = `${f.italic ? 'italic ' : ''}${TYPE_FONT_PX}px ${f.family}`
  ctx.font = font
  const m = ctx.measureText(text)
  // Tight bounds from actual glyph extents; padding covers italic overhang and antialiasing
  const pad = Math.ceil(TYPE_FONT_PX * 0.08)
  const left = Math.ceil(m.actualBoundingBoxLeft) + pad
  const ascent = Math.ceil(m.actualBoundingBoxAscent) + pad
  canvas.width = left + Math.ceil(m.actualBoundingBoxRight) + pad
  canvas.height = ascent + Math.ceil(m.actualBoundingBoxDescent) + pad
  if (canvas.width < 2 || canvas.height < 2) return null
  // Resizing the canvas resets the context state, so the font must be set again
  ctx.font = font
  ctx.fillStyle = cssRgb(color)
  ctx.fillText(text, left, ascent)
  const base64 = canvas.toDataURL('image/png').split(',')[1]
  return base64
    ? { kind: 'image', image: base64, width: canvas.width, height: canvas.height }
    : null
}

/** Signature dialog: draw on a pad, type text, or upload an image; confirming enters placement mode */
export function SignatureDialog({
  color,
  t,
  onCancel,
  onConfirm,
}: {
  color: [number, number, number]
  t: TFunc
  onCancel: () => void
  onConfirm: (sig: SignatureData) => void
}): ReactElement {
  const [mode, setMode] = useState<'draw' | 'type' | 'image'>('draw')
  const [typed, setTyped] = useState('')
  const [fontIdx, setFontIdx] = useState(0)
  const [imgCanvas, setImgCanvas] = useState<HTMLCanvasElement | null>(null)
  const [bw, setBw] = useState(false)
  const fileRef = useRef<HTMLInputElement>(null)
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const pathsRef = useRef<number[][]>([])
  const curRef = useRef<number[] | null>(null)

  /** Processed image + preview data URL, recomputed when the source or the B&W toggle changes */
  const processedImg = useMemo(() => {
    if (!imgCanvas) return null
    const canvas = bw ? toBlackAndWhite(imgCanvas) : imgCanvas
    return { canvas, url: canvas.toDataURL('image/png') }
  }, [imgCanvas, bw])

  const pickImage = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    e.target.value = ''
    if (!file) return
    const canvas = await fileToCanvas(file)
    if (canvas) setImgCanvas(canvas)
  }

  const redraw = () => {
    const canvas = canvasRef.current
    const ctx = canvas?.getContext('2d')
    if (!canvas || !ctx) return
    ctx.clearRect(0, 0, canvas.width, canvas.height)
    ctx.strokeStyle = cssRgb(color)
    ctx.lineWidth = 2.4
    ctx.lineCap = 'round'
    ctx.lineJoin = 'round'
    for (const p of pathsRef.current) {
      ctx.beginPath()
      ctx.moveTo(p[0]!, p[1]!)
      for (let i = 2; i < p.length; i += 2) ctx.lineTo(p[i]!, p[i + 1]!)
      ctx.stroke()
    }
  }

  useEffect(redraw, [color, mode])

  const pos = (e: ReactPointerEvent): [number, number] => {
    const box = e.currentTarget.getBoundingClientRect()
    return [e.clientX - box.left, e.clientY - box.top]
  }

  const down = (e: ReactPointerEvent) => {
    e.currentTarget.setPointerCapture(e.pointerId)
    curRef.current = pos(e)
    pathsRef.current = [...pathsRef.current, curRef.current]
  }
  const move = (e: ReactPointerEvent) => {
    if (!curRef.current) return
    const [x, y] = pos(e)
    curRef.current.push(x, y)
    redraw()
  }
  const up = () => {
    curRef.current = null
  }

  const clear = () => {
    pathsRef.current = []
    redraw()
  }

  const confirm = () => {
    if (mode === 'draw') {
      const paths = pathsRef.current.filter((p) => p.length >= 4)
      if (paths.length > 0) onConfirm({ kind: 'strokes', paths, width: PAD_W, height: PAD_H })
      return
    }
    if (mode === 'image') {
      if (!processedImg) return
      const base64 = processedImg.canvas.toDataURL('image/png').split(',')[1]
      if (base64) {
        onConfirm({
          kind: 'image',
          image: base64,
          width: processedImg.canvas.width,
          height: processedImg.canvas.height,
        })
      }
      return
    }
    const text = typed.trim()
    if (!text) return
    const sig = textToImage(text, SIGN_FONTS[fontIdx] ?? SIGN_FONTS[0]!, color)
    if (sig) onConfirm(sig)
  }

  const canConfirm =
    mode === 'draw' || (mode === 'image' ? processedImg !== null : typed.trim().length > 0)

  return (
    <div className="pdf-modal-mask" onClick={onCancel}>
      <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
        <div className="pdf-modal-title">{t('signTitle')}</div>
        <div className="pdf-sign-tabs">
          <button
            className={`pdf-sign-tab${mode === 'draw' ? ' active' : ''}`}
            onClick={() => setMode('draw')}
          >
            {t('signDraw')}
          </button>
          <button
            className={`pdf-sign-tab${mode === 'type' ? ' active' : ''}`}
            onClick={() => setMode('type')}
          >
            {t('signType')}
          </button>
          <button
            className={`pdf-sign-tab${mode === 'image' ? ' active' : ''}`}
            onClick={() => setMode('image')}
          >
            {t('signImage')}
          </button>
        </div>
        {mode === 'draw' && (
          <canvas
            ref={canvasRef}
            className="pdf-sign-pad"
            width={PAD_W}
            height={PAD_H}
            onPointerDown={down}
            onPointerMove={move}
            onPointerUp={up}
            onPointerCancel={up}
          />
        )}
        {mode === 'type' && (
          <>
            <input
              className="pdf-modal-input"
              value={typed}
              placeholder={t('signTypePlaceholder')}
              autoFocus
              onChange={(e) => setTyped(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && confirm()}
            />
            <div className="pdf-sign-fontgrid">
              {SIGN_FONTS.map((f, i) => (
                <button
                  key={f.id}
                  className={`pdf-sign-fontcard${i === fontIdx ? ' active' : ''}`}
                  style={{
                    color: cssRgb(color),
                    fontFamily: f.family,
                    fontStyle: f.italic ? 'italic' : 'normal',
                  }}
                  onClick={() => setFontIdx(i)}
                >
                  {typed || t('signTypePlaceholder')}
                </button>
              ))}
            </div>
          </>
        )}
        {mode === 'image' && (
          <>
            <input
              ref={fileRef}
              type="file"
              accept="image/*"
              style={{ display: 'none' }}
              onChange={(e) => void pickImage(e)}
            />
            {processedImg ? (
              <div className="pdf-sign-imgbox" onClick={() => fileRef.current?.click()}>
                <img className="pdf-sign-img" src={processedImg.url} alt="" />
              </div>
            ) : (
              <button
                className="pdf-sign-imgbox pdf-sign-imgbox-empty"
                onClick={() => fileRef.current?.click()}
              >
                <span className="pdf-sign-imgplus">+</span>
                {t('signAddImage')}
              </button>
            )}
            <label className={`pdf-sign-bw${imgCanvas ? '' : ' disabled'}`}>
              <input
                type="checkbox"
                checked={bw}
                disabled={!imgCanvas}
                onChange={(e) => setBw(e.target.checked)}
              />
              {t('signBw')}
            </label>
          </>
        )}
        <div className="pdf-modal-actions">
          {mode === 'draw' && (
            <button className="pdf-modal-btn" onClick={clear}>
              {t('signClear')}
            </button>
          )}
          {mode === 'image' && imgCanvas && (
            <button
              className="pdf-modal-btn"
              onClick={() => {
                setImgCanvas(null)
                setBw(false)
              }}
            >
              {t('signClear')}
            </button>
          )}
          <span style={{ flex: 1 }} />
          <button className="pdf-modal-btn" onClick={onCancel}>
            {t('cancel')}
          </button>
          <button className="pdf-modal-btn primary" disabled={!canConfirm} onClick={confirm}>
            {t('signPlace')}
          </button>
        </div>
        <div className="pdf-sign-hint">{t('signHint')}</div>
      </div>
    </div>
  )
}
