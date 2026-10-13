import { placeholderSource } from './redact'
import type { GlyphRun, RenderNode, ShapeRenderNode, TextLine } from '@genoffice/pptx-render'

/**
 * The model-facing view of a slide: every withheld span becomes its marker and
 * the words it covers disappear. This is the **only** place a deck's text is
 * turned into something a model reads, so it is also the only place that has to
 * know about withholding at all.
 *
 * ## The one rule
 *
 * `forDisplay` defaults to **false**, which means the redacted view. A new call
 * site that forgets to pass anything therefore gets the safe answer; the leak
 * requires deliberately asking for the raw text. The two places that genuinely
 * need the words — the reader's own find-and-replace, and the node description
 * a card shows on their screen — pass `true` explicitly and are named as such at
 * the call site.
 *
 * ## Why the render tree and not the deck model
 *
 * Text here is already laid out: `text.lines[].runs[]` are the wrapped lines the
 * canvas draws, not the paragraphs in the file. That matters for a withheld span
 * long enough to wrap — the mark lands on both lines and the model reads
 * `{{label` + `}}` across a newline. The words are still hidden and the write
 * guard refuses a split marker, so the cost is a refused edit rather than a leak.
 * Reading the unwrapped deck model instead would need a new accessor on
 * `DeckAccess`, which only exposes the render tree, and a fallback between the
 * two sources would be fail-open. A visible false refusal is the better trade.
 */

/** A run as the model may see it. */
export function runTextForModel(run: GlyphRun, forDisplay = false): string {
  if (!forDisplay && run.redact) return placeholderSource(run.redact)
  return run.text
}

/** One laid-out line, joined. */
export function lineTextForModel(line: TextLine, forDisplay = false): string {
  return line.runs.map((r) => runTextForModel(r, forDisplay)).join('')
}

const isTextNode = (n: RenderNode): n is ShapeRenderNode => n.type === 'shape' || n.type === 'text'

/** The `lines` of a shape, or of one table cell. */
function linesOf(node: { text?: { lines?: TextLine[] } }): TextLine[] {
  return node.text?.lines ?? []
}

/**
 * A shape's text, tab-separated cells for a table — the same shape the existing
 * `nodeText` produced, with withheld runs swapped for their markers.
 */
export function textForModel(node: RenderNode, forDisplay = false): string {
  if (node.type === 'picture') {
    // a withheld picture reads as its marker so the model can see something is
    // there and write around it. The reader's own view stays empty: a picture
    // has no text, and inventing one for their card would be a lie.
    return !forDisplay && node.redact ? placeholderSource(node.redact) : ''
  }
  if (isTextNode(node)) {
    return linesOf(node)
      .map((line) => lineTextForModel(line, forDisplay))
      .join('\n')
  }
  if (node.type === 'table') {
    // row by row so the model can read table content; a cell's runs get the
    // same treatment as a shape's
    const byRow = new Map<number, string[]>()
    for (const cell of node.cells) {
      const text = linesOf(cell)
        .map((line) => lineTextForModel(line, forDisplay))
        .join(' ')
      const row = byRow.get(cell.y) ?? []
      row.push(text)
      byRow.set(cell.y, row)
    }
    return [...byRow.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([, r]) => r.join('\t'))
      .join('\n')
  }
  return ''
}

/**
 * A one-line description of a media node for the model's element inventory.
 *
 * A withheld picture, video or audio shape reads as its marker, so the model can
 * see that something is there and write around it — which is the whole contract
 * the prompt asks of it. What it must not get is anything it could fetch: the
 * render tree deliberately carries only `media`'s *kind* (`PictureRenderNode.media`
 * is `'video' | 'audio'`, never the target), so an external `r:link` URL does not
 * reach this layer at all.
 */
export function mediaLineForModel(node: RenderNode): string | null {
  if (node.type !== 'picture') return null
  return node.redact ? placeholderSource(node.redact) : null
}

/** A withheld media node's own line, for callers that list media separately. */

/**
 * The distinct labels withheld on a slide, for the model's instruction.
 *
 * Distinct, because a run is laid out as many glyph tokens and every one of them
 * carries the mark — counting them would report one span as four, and the
 * prompt would list the same placeholder over and over. Two spans that happen
 * to share a label also collapse into one entry, which is what the instruction
 * wants anyway: it is a list of names to treat as opaque, not a census.
 */
export function redactLabelsOf(slide: { nodes: RenderNode[] }): string[] {
  const seen = new Set<string>()
  const visit = (nodes: RenderNode[]) => {
    for (const n of nodes) {
      if (n.type === 'group') {
        visit(n.children)
        continue
      }
      if (n.type === 'picture' && n.redact) {
        seen.add(n.redact)
        continue
      }
      if (!isTextNode(n)) continue
      for (const line of linesOf(n)) {
        for (const r of line.runs) if (r.redact) seen.add(r.redact)
      }
    }
  }
  visit(slide.nodes)
  return [...seen]
}

/**
 * How many distinct labels this slide withholds. The guard's "is there anything
 * to check at all" — a deck with no withheld spans costs nothing to pass through.
 */
export function redactionCount(slide: { nodes: RenderNode[] }): number {
  return redactLabelsOf(slide).length
}
