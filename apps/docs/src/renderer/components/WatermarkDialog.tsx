/**
 * Word's Insert Watermark dialog: No watermark / Picture (scale, washout) /
 * Text (face, size, bold, italic, transparency, color, layout), with a page
 * preview drawn through the same helpers the page preview and save use.
 */
import { useEffect, useState } from 'react'
import {
  pictureWatermarkPreviewImage,
  textWatermarkPreviewImage,
  type HfImage,
  type PictureWatermarkSpec,
  type Watermark,
  type WatermarkSpec,
} from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import { fontFamiliesFor } from '../font-list'
import { hfWashoutFilter, wordArtSvgMarkup } from '../editor/hf-dom'
import { imageSizeOf } from './ribbon-tabs'
import { useModalKeys } from './modal-keys'

type Mode = 'none' | 'picture' | 'text'

const SCALES = [500, 200, 150, 100, 50]
const SIZES = [36, 40, 44, 48, 54, 60, 66, 72, 80, 90, 96, 105, 120, 144]
const DEFAULT_COLOR = 'C0C0C0'
const DEFAULT_FONT = 'Calibri'
const PICKABLE_MIME = new Set(['image/png', 'image/jpeg', 'image/gif'])

export interface WatermarkDialogCurrent {
  text: string | null
  style: Omit<WatermarkSpec, 'text'> | null
  picture: PictureWatermarkSpec | null
  /** the saved document's watermark shape (parsed header image), used when nothing is pending */
  docImage?: HfImage | null
}

/** section geometry in points; null falls back to Letter with 1in margins */
export interface WatermarkPageBox {
  pageWidthPt: number
  pageHeightPt: number
  marginWidthPt: number
  marginHeightPt: number
}

const LETTER: WatermarkPageBox = {
  pageWidthPt: 612,
  pageHeightPt: 792,
  marginWidthPt: 468,
  marginHeightPt: 648,
}

let inkCtx: CanvasRenderingContext2D | null | undefined

/** glyph ink box of the string at `sizePt`, i.e. the shape box Word gives an explicit Size */
function inkBoxPt(spec: WatermarkSpec, sizePt: number): { widthPt: number; heightPt: number } {
  if (inkCtx === undefined) inkCtx = document.createElement('canvas').getContext('2d')
  const fallback = { widthPt: sizePt * 0.6 * spec.text.length, heightPt: sizePt * 0.72 }
  if (!inkCtx) return fallback
  inkCtx.font = `${spec.italic ? 'italic' : 'normal'} ${spec.bold ? 'bold' : 'normal'} ${sizePt}px "${spec.fontFamily ?? DEFAULT_FONT}", sans-serif`
  const m = inkCtx.measureText(spec.text)
  const w = (m.actualBoundingBoxLeft ?? 0) + (m.actualBoundingBoxRight ?? m.width)
  const h = (m.actualBoundingBoxAscent ?? 0) + (m.actualBoundingBoxDescent ?? 0)
  if (!(w > 0) || !(h > 0)) return fallback
  const r = (v: number) => Math.round(v * 100) / 100
  return { widthPt: r(w), heightPt: r(h) }
}

const AUTO_BOX_HEIGHTS = [247.45, 131.95]

/** the Size a saved WordArt box was measured at; 'auto' for Word's auto boxes */
function sizeOfBox(spec: WatermarkSpec, heightPt: number): number | 'auto' {
  if (AUTO_BOX_HEIGHTS.some((h) => Math.abs(h - heightPt) < 0.5)) return 'auto'
  const perPt = inkBoxPt(spec, 100).heightPt / 100
  return perPt > 0 ? Math.max(1, Math.round(heightPt / perPt)) : 'auto'
}

/** parsed header shape → the dialog's text settings */
function textFromDocImage(img: HfImage): WatermarkSpec | null {
  const wa = img.wordArt
  if (!wa) return null
  const spec: WatermarkSpec = {
    text: wa.text,
    fontFamily: wa.fontFamily ?? DEFAULT_FONT,
    colorHex: wa.colorHex,
    opacity: wa.opacity,
    diagonal: (img.rotationDeg ?? 0) !== 0,
    ...(wa.bold ? { bold: true } : {}),
    ...(wa.italic ? { italic: true } : {}),
  }
  const heightPt = (img.heightPx ?? 0) * 0.75
  const size = heightPt > 0 ? sizeOfBox(spec, heightPt) : 'auto'
  if (size !== 'auto') {
    spec.sizePt = size
    spec.boxPt = inkBoxPt(spec, size)
  }
  return spec
}

/** parsed picture frame → picture settings; the natural size needs a decode, so async */
async function pictureFromDocImage(
  img: HfImage,
  box: WatermarkPageBox,
): Promise<PictureWatermarkSpec | null> {
  const m = /^data:(image\/(?:png|jpeg|gif));base64,(.*)$/s.exec(img.dataUrl)
  if (!m) return null
  const natural = await imageSizeOf(img.dataUrl)
  const spec: PictureWatermarkSpec = {
    image: {
      base64: m[2]!,
      mime: m[1] as PictureWatermarkSpec['image']['mime'],
      widthPx: natural.width,
      heightPx: natural.height,
    },
    washout: Boolean(img.washout),
  }
  const naturalWPt = (natural.width * 72) / 96
  const naturalHPt = (natural.height * 72) / 96
  const autoFactor = Math.min(box.marginWidthPt / naturalWPt, box.marginHeightPt / naturalHPt)
  const factor = ((img.widthPx ?? 0) * 0.75) / naturalWPt
  if (factor > 0 && Math.abs(factor - autoFactor) > 0.01) spec.scale = Math.round(factor * 100)
  return spec
}

export function WatermarkDialog({
  current,
  presets,
  pageBox,
  onApply,
  onClose,
}: {
  current: WatermarkDialogCurrent
  presets: string[]
  pageBox: WatermarkPageBox | null
  onApply: (spec: Watermark | null) => void
  onClose: () => void
}) {
  const { t, lang } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const box = pageBox ?? LETTER
  // pending edits win; otherwise the saved shape seeds the fields (text style is
  // not kept on the file context, only its text is)
  const docText = !current.picture && !current.style && current.docImage?.wordArt
  const docPicture = !current.picture && !current.text && current.docImage?.watermark
  const cur: Partial<WatermarkSpec> =
    current.style ?? (docText ? (textFromDocImage(current.docImage!) ?? {}) : {})

  const [mode, setMode] = useState<Mode>(
    current.picture || docPicture ? 'picture' : current.text ? 'text' : 'none',
  )
  const [image, setImage] = useState<PictureWatermarkSpec['image'] | null>(
    current.picture?.image ?? null,
  )
  const [scale, setScale] = useState<number | 'auto'>(current.picture?.scale ?? 'auto')
  const [washout, setWashout] = useState(
    current.picture
      ? current.picture.washout !== false
      : !docPicture || Boolean(current.docImage?.washout),
  )
  useEffect(() => {
    if (!docPicture) return
    let alive = true
    void pictureFromDocImage(current.docImage!, box).then((spec) => {
      if (!alive || !spec) return
      setImage(spec.image)
      setScale(spec.scale ?? 'auto')
      setWashout(spec.washout !== false)
    })
    return () => {
      alive = false
    }
    // seeds once from the shape the dialog opened on
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])
  const [text, setText] = useState(current.text ?? presets[0] ?? 'DRAFT')
  const [font, setFont] = useState(cur.fontFamily ?? DEFAULT_FONT)
  const [size, setSize] = useState<number | 'auto'>(cur.sizePt ?? 'auto')
  const [bold, setBold] = useState(cur.bold ?? false)
  const [italic, setItalic] = useState(cur.italic ?? false)
  const [transparency, setTransparency] = useState(Math.round((1 - (cur.opacity ?? 0.5)) * 100))
  const [color, setColor] = useState((cur.colorHex ?? DEFAULT_COLOR).replace(/^#/, ''))
  const [diagonal, setDiagonal] = useState(cur.diagonal !== false)
  const [pickError, setPickError] = useState(false)

  const fonts = fontFamiliesFor(lang)
  const fontOptions = fonts.includes(font) ? fonts : [font, ...fonts]

  const textSpec = (): WatermarkSpec | null => {
    const trimmed = text.trim()
    if (!trimmed) return null
    const spec: WatermarkSpec = {
      text: trimmed,
      fontFamily: font,
      colorHex: color.toUpperCase(),
      opacity: Math.round((1 - transparency / 100) * 100) / 100,
      diagonal,
      ...(bold ? { bold: true } : {}),
      ...(italic ? { italic: true } : {}),
    }
    if (size !== 'auto') {
      spec.sizePt = size
      spec.boxPt = inkBoxPt(spec, size)
    }
    return spec
  }
  const pictureSpec = (): PictureWatermarkSpec | null =>
    image ? { image, ...(scale === 'auto' ? {} : { scale }), washout } : null

  const resolved: Watermark | null =
    mode === 'text' ? textSpec() : mode === 'picture' ? pictureSpec() : null
  const canApply = mode === 'none' || resolved !== null

  const pick = async () => {
    const picked = await window.desktop.pickImage()
    if (!picked) return
    if (!PICKABLE_MIME.has(picked.mime)) {
      setPickError(true)
      return
    }
    try {
      const natural = await imageSizeOf(`data:${picked.mime};base64,${picked.base64}`)
      setImage({
        base64: picked.base64,
        mime: picked.mime as PictureWatermarkSpec['image']['mime'],
        widthPx: natural.width,
        heightPx: natural.height,
      })
      setPickError(false)
    } catch {
      setPickError(true)
    }
  }

  const submit = () => {
    if (!canApply) return
    onApply(resolved)
    onClose()
  }

  const previewImage: HfImage | null =
    resolved === null
      ? null
      : 'image' in resolved
        ? pictureWatermarkPreviewImage(resolved, {
            widthPt: box.marginWidthPt,
            heightPt: box.marginHeightPt,
          })
        : textWatermarkPreviewImage(resolved)

  const textOn = mode === 'text'
  const picOn = mode === 'picture'

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        className="modal watermark-dialog"
        role="dialog"
        aria-label={t('ribbonWatermarkDialogTitle')}
      >
        <h2>{t('ribbonWatermarkDialogTitle')}</h2>
        <div className="wm-body">
          <div className="wm-form">
            <label className="wm-mode">
              <input type="radio" checked={mode === 'none'} onChange={() => setMode('none')} />
              {t('ribbonWatermarkNone')}
            </label>

            <label className="wm-mode">
              <input type="radio" checked={picOn} onChange={() => setMode('picture')} />
              {t('ribbonWatermarkPicture')}
            </label>
            <div className="wm-fields" aria-disabled={!picOn}>
              <div className="wm-row">
                <button className="btn-ghost wm-pick" disabled={!picOn} onClick={() => void pick()}>
                  {t('ribbonWatermarkSelectPicture')}
                </button>
                <span className="wm-pick-name">
                  {image
                    ? `${image.widthPx} × ${image.heightPx} px`
                    : t('ribbonWatermarkNoPicture')}
                </span>
              </div>
              {pickError && <div className="modal-error">{t('ribbonImageLoadFail')}</div>}
              <div className="wm-row">
                <span className="wm-label">{t('ribbonWatermarkScale')}</span>
                <select
                  disabled={!picOn}
                  value={scale}
                  onChange={(e) =>
                    setScale(e.target.value === 'auto' ? 'auto' : Number(e.target.value))
                  }
                >
                  <option value="auto">{t('ribbonWatermarkAuto')}</option>
                  {(scale === 'auto' || SCALES.includes(scale)
                    ? SCALES
                    : [...SCALES, scale].sort((a, b) => b - a)
                  ).map((s) => (
                    <option key={s} value={s}>
                      {s}%
                    </option>
                  ))}
                </select>
                <label className="wm-check">
                  <input
                    type="checkbox"
                    disabled={!picOn}
                    checked={washout}
                    onChange={(e) => setWashout(e.target.checked)}
                  />
                  {t('ribbonWatermarkWashout')}
                </label>
              </div>
            </div>

            <label className="wm-mode">
              <input type="radio" checked={textOn} onChange={() => setMode('text')} />
              {t('ribbonWatermarkText')}
            </label>
            <div className="wm-fields" aria-disabled={!textOn}>
              <div className="wm-row">
                <input
                  type="text"
                  list="wm-text-presets"
                  disabled={!textOn}
                  value={text}
                  placeholder={t('ribbonWatermarkPh')}
                  onChange={(e) => setText(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && submit()}
                  autoFocus={textOn}
                />
                <datalist id="wm-text-presets">
                  {presets.map((p) => (
                    <option key={p} value={p} />
                  ))}
                </datalist>
              </div>
              <div className="wm-row">
                <span className="wm-label">{t('ribbonWatermarkFont')}</span>
                <select disabled={!textOn} value={font} onChange={(e) => setFont(e.target.value)}>
                  {fontOptions.map((f) => (
                    <option key={f} value={f} style={{ fontFamily: f }}>
                      {f}
                    </option>
                  ))}
                </select>
              </div>
              <div className="wm-row">
                <span className="wm-label">{t('ribbonWatermarkSize')}</span>
                <select
                  className="wm-size"
                  disabled={!textOn}
                  value={size}
                  onChange={(e) =>
                    setSize(e.target.value === 'auto' ? 'auto' : Number(e.target.value))
                  }
                >
                  <option value="auto">{t('ribbonWatermarkAuto')}</option>
                  {(size === 'auto' || SIZES.includes(size)
                    ? SIZES
                    : [...SIZES, size].sort((a, b) => a - b)
                  ).map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
                <label className="wm-check">
                  <input
                    type="checkbox"
                    disabled={!textOn}
                    checked={bold}
                    onChange={(e) => setBold(e.target.checked)}
                  />
                  {t('appFontBold')}
                </label>
                <label className="wm-check">
                  <input
                    type="checkbox"
                    disabled={!textOn}
                    checked={italic}
                    onChange={(e) => setItalic(e.target.checked)}
                  />
                  {t('appFontItalic')}
                </label>
              </div>
              <div className="wm-row">
                <span className="wm-label">{t('ribbonWatermarkTransparency')}</span>
                <input
                  type="range"
                  min={0}
                  max={100}
                  disabled={!textOn}
                  value={transparency}
                  onChange={(e) => setTransparency(Number(e.target.value))}
                  aria-label={t('ribbonWatermarkTransparency')}
                />
                <input
                  type="number"
                  className="wm-pct"
                  min={0}
                  max={100}
                  disabled={!textOn}
                  value={transparency}
                  onChange={(e) =>
                    setTransparency(Math.max(0, Math.min(100, Number(e.target.value) || 0)))
                  }
                />
                <span className="wm-unit">%</span>
              </div>
              <div className="wm-row">
                <span className="wm-label">{t('ribbonWatermarkFontColor')}</span>
                <input
                  type="color"
                  className="font-color-input wm-color"
                  disabled={!textOn}
                  value={`#${color}`}
                  onChange={(e) => setColor(e.target.value.slice(1))}
                  aria-label={t('ribbonWatermarkFontColor')}
                />
              </div>
              <div className="wm-row">
                <span className="wm-label">{t('ribbonWatermarkOrientation')}</span>
                <label className="wm-check">
                  <input
                    type="radio"
                    disabled={!textOn}
                    checked={!diagonal}
                    onChange={() => setDiagonal(false)}
                  />
                  {t('ribbonWatermarkHorizontal')}
                </label>
                <label className="wm-check">
                  <input
                    type="radio"
                    disabled={!textOn}
                    checked={diagonal}
                    onChange={() => setDiagonal(true)}
                  />
                  {t('ribbonWatermarkDiagonal')}
                </label>
              </div>
            </div>
          </div>
          <WatermarkPagePreview box={box} image={previewImage} />
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" disabled={!canApply} onClick={submit}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

const PREVIEW_HEIGHT_PX = 340
const PREVIEW_LINES = 36

/** scaled page: grey body lines inside the margin box, the watermark centered on it */
function WatermarkPagePreview({ box, image }: { box: WatermarkPageBox; image: HfImage | null }) {
  const k = PREVIEW_HEIGHT_PX / box.pageHeightPt
  const pageW = box.pageWidthPt * k
  const marginW = box.marginWidthPt * k
  const marginH = box.marginHeightPt * k
  const [svg, setSvg] = useState('')
  // text ink is measured against document fonts, which may still be loading
  useEffect(() => {
    if (!image?.wordArt) {
      setSvg('')
      return
    }
    setSvg(wordArtSvgMarkup(image))
    let alive = true
    void document.fonts?.ready.then(() => alive && setSvg(wordArtSvgMarkup(image)))
    return () => {
      alive = false
    }
  }, [image])
  const wmW = (image?.widthPx ?? 0) * 0.75 * k
  const wmH = (image?.heightPx ?? 0) * 0.75 * k
  const wmStyle: React.CSSProperties = {
    width: wmW,
    height: wmH,
    left: (marginW - wmW) / 2,
    top: (marginH - wmH) / 2,
    ...(image?.rotationDeg ? { transform: `rotate(${image.rotationDeg}deg)` } : {}),
    ...(image?.washout ? { filter: hfWashoutFilter(image.washout) } : {}),
  }
  return (
    <div className="wm-preview" style={{ width: pageW, height: PREVIEW_HEIGHT_PX }}>
      <div
        className="wm-preview-margin"
        style={{
          width: marginW,
          height: marginH,
          left: (pageW - marginW) / 2,
          top: (PREVIEW_HEIGHT_PX - marginH) / 2,
        }}
      >
        {image &&
          (image.wordArt ? (
            <span
              className="wm-preview-mark"
              style={wmStyle}
              dangerouslySetInnerHTML={{ __html: svg }}
            />
          ) : (
            <img className="wm-preview-mark" alt="" src={image.dataUrl} style={wmStyle} />
          ))}
        <div className="wm-preview-lines">
          {Array.from({ length: PREVIEW_LINES }, (_, i) => (
            <span key={i} style={i % 4 === 3 ? { width: '55%' } : undefined} />
          ))}
        </div>
      </div>
    </div>
  )
}
