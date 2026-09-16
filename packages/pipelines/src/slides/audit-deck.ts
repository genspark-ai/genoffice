import { elementDurableId, openPptx, slideDurableId } from '@genoffice/pptx-engine'
import { buildRenderSlide, EMU_PER_PX_96, type FontMetricsProvider } from '@genoffice/pptx-render'
import { auditSlideFindings, type AuditFinding } from './layout-audit'

export interface DeckAuditPage {
  /** 0-based slide index */
  slide: number
  id: string
  issues: string[]
  /** the same issues typed, with a `setTransform` suggestion where geometry alone fixes one */
  findings: AuditFinding[]
}

/**
 * Lays every slide out at native pixel size and runs the geometry audit; ids in
 * the findings are the durable ids ops target. Without metrics the layout uses
 * heuristic glyph widths, so overflow figures are approximate.
 */
export async function auditDeck(
  bytes: Uint8Array,
  opts: { metrics?: FontMetricsProvider; only?: number } = {},
): Promise<DeckAuditPage[]> {
  const opened = await openPptx(bytes)
  const fitWidthPx = opened.deck.size.cx / EMU_PER_PX_96
  const out: DeckAuditPage[] = []
  for (const [index, slide] of opened.deck.slides.entries()) {
    if (opts.only !== undefined && opts.only !== index) continue
    const ids = new Map(slide.elements.map((el) => [el.id, elementDurableId(el) ?? el.id]))
    const rendered = buildRenderSlide(slide, opened.deck.size, {
      fitWidthPx,
      ...(opts.metrics ? { metrics: opts.metrics } : {}),
      slideNo: index + 1,
    })
    const slideId = slideDurableId(slide)
    const findings = auditSlideFindings(rendered, (id) => ids.get(id) ?? id).map((f) =>
      f.suggest
        ? { ...f, suggest: { ...f.suggest, target: { slide: slideId, el: f.suggest.target.el } } }
        : f,
    )
    out.push({ slide: index, id: slideId, issues: findings.map((f) => f.message), findings })
  }
  return out
}
