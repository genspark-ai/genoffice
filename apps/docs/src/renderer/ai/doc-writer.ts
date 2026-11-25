// Modified for ForaOffice from GenOffice v0.12.852.
import { HTML_RULES } from '../../shared/html-rules'

/**
 * Long-form writing runs as its own tool-less request whose reply body IS the
 * restricted-HTML fragment. Text deltas stream straight into the document as
 * a draft, so the user watches the text arrive and no gateway sees a silent
 * connection (tool arguments are buffered server-side until the JSON is
 * complete; a document-sized argument exceeds gateway idle cutoffs).
 */
const CONTEXT_CAP = 8000

export interface DocWriteSpec {
  plan: string
  title?: string | undefined
  context?: string | undefined
  /** the user's own request, verbatim */
  instruction: string
}

export interface DocWriteResult {
  ok: boolean
  /** the fragment to land (complete, or the adopted partial) */
  html?: string
  /** the stream ended early: transport drop, stop, or output cap */
  truncated?: boolean
  error?: string
}

/** Panel-owned writer: streams the fragment, reports progress, settles partial output with the user. */
export interface DocWriterObserver {
  onProgress(html: string): void
  onStopped(): Promise<void>
}

export interface AiDocWriter {
  write(
    spec: DocWriteSpec,
    observer: DocWriterObserver,
    signal?: AbortSignal,
  ): Promise<DocWriteResult>
}

const WRITER_SYSTEM = [
  'You are the document writer of GenOffice Docs, a Word-style editor. You receive a plan (title, sections, key points, tone, length) and reference material, and you write the complete content.',
  '',
  '## Writing rules',
  '- Write the requested content in full: every section in the plan, real prose, no placeholders, no "[to be added]" notes.',
  '- Real data only: use the facts, figures, names and quotes from the reference material and the plan; never invent facts or numbers beyond them.',
  '- Follow the length the plan asks for; when it gives none, write what the content genuinely needs.',
  '',
  '## Output',
  'Reply with the restricted HTML fragment only: it starts with the first block tag and ends with the last closing tag. No markdown code fences, no commentary before or after, no <html>/<body> wrapper.',
  HTML_RULES,
].join('\n')

export function buildDocWriterRequest(
  spec: DocWriteSpec,
  langDirective: string,
): { system: string; user: string } {
  const context = spec.context?.trim()
    ? `\n\n## Reference material (all names, figures and facts come from here)\n${spec.context.slice(0, CONTEXT_CAP)}`
    : ''
  const instruction = spec.instruction.trim()
    ? `\n\n## The user's request\n${spec.instruction.slice(0, 2000)}`
    : ''
  return {
    system: WRITER_SYSTEM + langDirective,
    user: `## Plan\n${spec.title ? `Title: ${spec.title}\n` : ''}${spec.plan.slice(0, 6000)}${context}${instruction}\n\nWrite the content now.`,
  }
}

/** Strip a fence or chatter before the first block tag and a closing fence at the end. */
export function extractFragment(raw: string): string {
  let text = raw.replace(/\r\n/g, '\n')
  const start = text.search(/<(?:h[1-6]|p|ul|ol|table|pre|blockquote|formula)[\s>]/i)
  if (start > 0) text = text.slice(start)
  else if (start < 0) text = text.replace(/^\s*```[a-z]*\n?/i, '')
  return text.replace(/\n?```\s*$/, '').trimEnd()
}

/** rough top-level block count of a fragment, for the progress chip */
export function countFragmentBlocks(html: string): number {
  return (html.match(/<(?:h[1-6]|p|ul|ol|table|pre|blockquote|formula)[\s>]/gi) ?? []).length
}
