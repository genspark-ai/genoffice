import type { Mermaid } from 'mermaid'
import type { DiagramResult } from './diagrams'

export const MERMAID_LANGUAGE = 'mermaid'

export const MERMAID_TEMPLATE = [
  '```mermaid',
  'flowchart LR',
  '    A[Start] --> B{Decision}',
  '    B -->|Yes| C[Done]',
  '    B -->|No| A',
  '```',
].join('\n')

let loading: Promise<Mermaid> | null = null

/** Loaded on first use so documents without diagrams never pay for the bundle */
export function loadMermaid(): Promise<Mermaid> {
  loading ??= import('mermaid').then(({ default: mermaid }) => {
    mermaid.initialize({
      startOnLoad: false,
      securityLevel: 'strict',
      // diagrams are document content: one fixed theme in both UI themes
      theme: 'default',
      // SVG text labels instead of <foreignObject> so the diagram can be drawn
      // onto a canvas (docx export) without tainting it
      htmlLabels: false,
      flowchart: { htmlLabels: false },
      suppressErrorRendering: true,
    })
    return mermaid
  })
  return loading
}

let renderSeq = 0

export async function renderMermaid(source: string): Promise<DiagramResult> {
  const mermaid = await loadMermaid()
  try {
    const { svg } = await mermaid.render(`md-mermaid-${++renderSeq}`, source)
    return { ok: true, svg }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}
