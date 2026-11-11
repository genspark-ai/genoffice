import { openPptContainer } from './container'
import { parsePptPresentation } from './ppt'
import { writePptxPresentation } from './pptx-writer'

export { LegacyPptError, LEGACY_PPT_ERROR_CODES, type LegacyPptErrorCode } from './errors'
export type { PptPresentation, PptSlide, PptSlideText } from './ppt'
export type { PptShape, PptImage, PptMaster, EscherRect } from './escher'
export type { TextCharFormat, TextParagraphFormat, TextStyle } from './text-style'
export { parsePptPresentation } from './ppt'

export interface PptConversionResult {
  /** the .pptx bytes, ready to write atomically like any other deck */
  bytes: Uint8Array
  slideCount: number
  /** what stage 1 did not carry over, so callers (and later stages) can say so */
  warnings: string[]
}

/**
 * Legacy .ppt → .pptx. Stage 1 made every text sequence editable; stage 3
 * carries the Escher anchors (shape position/size, type and flips),
 * StyleTextProp character/paragraph runs, embedded PNG/JPEG pictures and
 * the main master's shapes. Text with no anchor still uses the stage-1
 * stacked layout, so a deck without a drawing keeps its old output.
 */
export async function convertPptToPptx(bytes: Uint8Array): Promise<PptConversionResult> {
  const container = openPptContainer(bytes)
  const presentation = parsePptPresentation(
    container.document,
    container.currentUser,
    container.pictures,
  )
  const out = await writePptxPresentation(presentation)
  const warnings: string[] = []
  const unreadable = presentation.images.filter((image) => image === null).length
  if (unreadable > 0) warnings.push(`${unreadable} picture(s) use a format stage 3 does not carry`)
  warnings.push('shape rotation and metafile pictures are not carried')
  return { bytes: out, slideCount: presentation.slides.length, warnings }
}
