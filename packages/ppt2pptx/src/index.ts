import { openPptContainer } from './container'
import { parsePptPresentation } from './ppt'
import { writePptxPresentation } from './pptx-writer'

export { LegacyPptError, LEGACY_PPT_ERROR_CODES, type LegacyPptErrorCode } from './errors'
export type {
  PptPresentation,
  PptSlide,
  PptSlideText,
  PptAnyShape,
  PptShape,
  PptPicture,
  PptFill,
} from './ppt'
export { parsePptPresentation } from './ppt'

export interface PptConversionResult {
  /** the .pptx bytes, ready to write atomically like any other deck */
  bytes: Uint8Array
  slideCount: number
  /** what stage 1 did not carry over, so callers (and later stages) can say so */
  warnings: string[]
}

/**
 * Legacy .ppt → .pptx. Every slide's drawing lands as real shapes at their
 * Escher anchors — autoshapes with fill and line, freeforms, pictures from the
 * Pictures stream, text with its StyleTextProp runs over the master's text
 * styles — plus the slide background and the master's decorations. Notes,
 * animations, hyperlinks, shadows and OLE objects beyond their preview are
 * not carried.
 */
export async function convertPptToPptx(bytes: Uint8Array): Promise<PptConversionResult> {
  const container = openPptContainer(bytes)
  const presentation = parsePptPresentation(
    container.document,
    container.currentUser,
    container.pictures,
  )
  const { bytes: out, warnings } = await writePptxPresentation(presentation)
  return { bytes: out, slideCount: presentation.slides.length, warnings }
}
