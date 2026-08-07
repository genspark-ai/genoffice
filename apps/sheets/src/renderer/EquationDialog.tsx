import { useMemo, useRef, useState } from 'react'

import { latexToOmml, ommlToMathML } from '@genoffice/docx-engine/math'
import { useI18n } from './i18n/locale'

/// Excel's Insert → Equation: LaTeX input with a live MathML preview (docs'
/// LaTeX→OMML pipeline, rendered natively by Chromium). Inserting rasterizes
/// the MathML to a transparent PNG so it rides the same picture pipeline
/// (journal + xl/media) as any inserted image.

const PRESETS = [
  'A = \\pi r^2',
  'a^2 + b^2 = c^2',
  'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}',
  '(x+a)^n = \\sum_{k=0}^{n} \\binom{n}{k} x^k a^{n-k}',
  'e^x = 1 + \\frac{x}{1!} + \\frac{x^2}{2!} + \\frac{x^3}{3!} + \\cdots',
  'f(x) = a_0 + \\sum_{n=1}^{\\infty} \\left( a_n \\cos \\frac{n\\pi x}{L} + b_n \\sin \\frac{n\\pi x}{L} \\right)',
]
const MATH_FONT = "'STIX Two Math', 'Cambria Math', 'Latin Modern Math', serif"
const FONT_PX = 30
/// Rasterize at 2× and anchor at the measured CSS size, so the PNG stays
/// sharp on retina displays.
const RASTER_SCALE = 2
/// Breathing room around the measured MathML box; glyphs (radicals, large
/// operators) can paint slightly outside their layout bounds.
const PADDING_PX = 4

function mathmlOf(latex: string): { mathml: string } | { error: string } | null {
  if (!latex.trim()) return null
  try {
    const omml = latexToOmml(latex)
    const mathml = ommlToMathML(`<m:oMath>${omml}</m:oMath>`)
    return mathml ? { mathml } : { error: latex }
  } catch (error) {
    return { error: error instanceof Error ? error.message : String(error) }
  }
}

/// MathML paints inside <foreignObject> even in the restricted SVG-in-<img>
/// mode (no external resources — the math font stack is system fonts).
/// The foreignObject subtree is XML-parsed (unlike the HTML live preview,
/// where the parser infers the namespace), so <math> needs an explicit
/// MathML xmlns or Chromium renders it as an unknown element.
function rasterizeMathml(mathml: string, width: number, height: number): Promise<string> {
  const namespaced = mathml.replace(
    /<math(?=[\s>])/g,
    '<math xmlns="http://www.w3.org/1998/Math/MathML"',
  )
  const style =
    `font-family:${MATH_FONT};font-size:${FONT_PX}px;color:#000;` +
    `display:inline-block;padding:${PADDING_PX}px`
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width * RASTER_SCALE}" ` +
    `height="${height * RASTER_SCALE}" viewBox="0 0 ${width} ${height}">` +
    `<foreignObject width="${width}" height="${height}">` +
    `<div xmlns="http://www.w3.org/1999/xhtml" style="${style}">${namespaced}</div>` +
    `</foreignObject></svg>`
  return new Promise((resolve, reject) => {
    const image = new Image()
    image.onload = () => {
      const canvas = document.createElement('canvas')
      canvas.width = width * RASTER_SCALE
      canvas.height = height * RASTER_SCALE
      const context = canvas.getContext('2d')
      if (!context) {
        reject(new Error('canvas 2d context unavailable'))
        return
      }
      context.drawImage(image, 0, 0, canvas.width, canvas.height)
      resolve(canvas.toDataURL('image/png'))
    }
    image.onerror = () => reject(new Error('equation rasterization failed'))
    image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`
  })
}

export function EquationDialog({
  onInsert,
  onClose,
}: {
  readonly onInsert: (dataUrl: string, width: number, height: number) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [latex, setLatex] = useState('')
  const [busy, setBusy] = useState(false)
  const previewRef = useRef<HTMLSpanElement>(null)
  const preview = useMemo(() => mathmlOf(latex), [latex])
  const canInsert = preview !== null && 'mathml' in preview

  const insert = (): void => {
    const node = previewRef.current
    if (busy || !canInsert || !node) return
    // The preview span uses the exact style the SVG raster does, so its
    // measured box is the raster's layout box.
    const rect = node.getBoundingClientRect()
    const width = Math.ceil(rect.width) + PADDING_PX * 2
    const height = Math.ceil(rect.height) + PADDING_PX * 2
    setBusy(true)
    rasterizeMathml(preview.mathml, width, height)
      .then((dataUrl) => {
        onInsert(dataUrl, width, height)
        onClose()
      })
      .catch(() => setBusy(false))
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog equation-dialog"
        role="dialog"
        aria-label={t('dlgEquationTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgEquationTitle')}</header>
        <section className="dialog-body">
          <div className="equation-presets">
            {PRESETS.map((preset) => {
              const rendered = mathmlOf(preset)
              return (
                <button key={preset} type="button" title={preset} onClick={() => setLatex(preset)}>
                  {rendered !== null && 'mathml' in rendered && (
                    <span dangerouslySetInnerHTML={{ __html: rendered.mathml }} />
                  )}
                </button>
              )
            })}
          </div>
          <input
            className="equation-latex"
            value={latex}
            placeholder={t('dlgEquationPlaceholder')}
            autoFocus
            onChange={(event) => setLatex(event.target.value)}
            onKeyDown={(event) => event.key === 'Enter' && insert()}
          />
          <div className="equation-preview">
            {preview === null && <span className="dialog-note">{t('dlgEquationHint')}</span>}
            {preview !== null && 'error' in preview && (
              <span className="equation-error">{preview.error}</span>
            )}
            {preview !== null && 'mathml' in preview && (
              <span
                ref={previewRef}
                style={{
                  fontFamily: MATH_FONT,
                  fontSize: FONT_PX,
                  display: 'inline-block',
                }}
                dangerouslySetInnerHTML={{ __html: preview.mathml }}
              />
            )}
          </div>
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary" disabled={!canInsert || busy} onClick={insert}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
