import { openPptContainer } from './cfb'
import { parsePptPresentation } from './ppt'
import { writePptxPresentation } from './pptx-writer'

export { LegacyPptError, LEGACY_PPT_ERROR_CODES, type LegacyPptErrorCode } from './errors'
export type { PptPresentation, PptSlide, PptSlideText } from './ppt'
export { parsePptPresentation } from './ppt'

export interface PptConversionResult {
  /** the .pptx bytes, ready to write atomically like any other deck */
  bytes: Uint8Array
  slideCount: number
  /** what stage 1 did not carry over, so callers (and later stages) can say so */
  warnings: string[]
}

/**
 * Legacy .ppt → .pptx, stage 1: editable text. Every slide's text sequences
 * (titles, bodies, notes-flagged others) become real text boxes on a real
 * slide, in the deck's order and slide size. Geometry (the Escher anchors),
 * formatting (StyleTextProp), images and masters are later stages; stage 1
 * stacks each slide's text boxes down the surface. The foundation the CLI
 * and the Slides open flow wire up in the following stages.
 */
export async function convertPptToPptx(bytes: Uint8Array): Promise<PptConversionResult> {
  const container = openPptContainer(bytes)
  const presentation = parsePptPresentation(container.document, container.currentUser)
  const out = await writePptxPresentation(presentation)
  const warnings: string[] = []
  warnings.push('shapes land as stacked text boxes; the original geometry is a later stage')
  warnings.push('text formatting (runs, indents, autofit) is a later stage')
  return { bytes: out, slideCount: presentation.slides.length, warnings }
}
