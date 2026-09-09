// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
/**
 * Generation layer: intent tree (IR) -> native OOXML via the `docx` library.
 *
 * Layout vocabulary used (deliberately no floating text boxes):
 *   card    -> single-cell borderless table with shading / left border
 *   kpirow  -> single-row N-column borderless table
 *   code    -> shaded paragraphs with a monospace font
 */
import { Packer } from 'docx'
import {
  addPageBackgroundFloat,
  createDocument,
  partitionIr,
  renderSection,
} from './generate/page-settings'
import { createRenderContext } from './generate/render-context'
import { Generator } from './generate/renderer'

async function generateDocx(ir, images) {
  const parts = partitionIr(ir)
  const context = createRenderContext(parts.docSettings)
  const generator = new Generator(images, context)
  addPageBackgroundFloat(generator, parts.pageBgNode)
  const rendered = renderSection(generator, parts)
  const document = createDocument(context, parts, rendered, generator)
  return Packer.toBuffer(document)
}

export { generateDocx }
