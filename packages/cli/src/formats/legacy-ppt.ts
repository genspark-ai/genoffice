import { convertPptToPptx, LegacyPptError } from '@genoffice/ppt2pptx'
import { CliError, EXIT } from '../result'

/**
 * .ppt → .pptx through the in-process legacy importer (stage 1: editable
 * text boxes in the original slide order and size; geometry, formatting and
 * images are later stages — the warnings ride in detail for scripts).
 */
export async function convertLegacyPpt(input: Uint8Array): Promise<{
  bytes: Uint8Array
  detail: Record<string, unknown>
}> {
  try {
    const result = await convertPptToPptx(input)
    return {
      bytes: result.bytes,
      detail: { slides: result.slideCount, warnings: result.warnings },
    }
  } catch (error) {
    // stable importer codes (encrypted, bad_container, ...) stay in detail for scripts
    if (error instanceof LegacyPptError)
      throw new CliError(EXIT.conversion, error.message, { reason: error.code })
    throw error
  }
}
