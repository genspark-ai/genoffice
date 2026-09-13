import { streamText, type AgentMessage, type AgentTransport } from '@genoffice/agent-core'
import type { PlanMode } from './tools'

/**
 * The brief (directions, sections, reference material) is drafted by its own
 * tool-less request whose reply body IS the JSON. A brief-sized tool argument is
 * buffered server-side and arrives in bursts with minute-long silences, which
 * the Genspark gateway cuts as an idle connection (observed: 4.6k chars, 64 s
 * without a byte). Text deltas stream continuously, so the same content is safe.
 */
export const BRIEF_MAX_CHARS = 40_000
const TRANSCRIPT_CAP = 24_000
const TURN_CAP = 6000
const PAGE_CAP = 6000
const NOTES_CAP = 1200

export interface BriefPlanSpec {
  mode: PlanMode
  /** what the agent wants honored: the user's choices, a redo note, direction hints */
  notes?: string
  /** the open page for restyle / extract; the writer reads its head and styles */
  page?: string
}

export type BriefPlanResult =
  { ok: true; raw: Record<string, unknown> } | { ok: false; error: string }

const SYSTEM = [
  'You are the brief writer of GenOffice HTML, a design-first editor for single-file HTML pages. From the conversation transcript you receive, draft the brief the user will confirm before the system writes the page.',
  '',
  '## The brief',
  '- core_hook: one sentence with tension, ideally a number or contrast.',
  '- style: the recommended direction; alternatives: 1–2 other directions that differ meaningfully in mood, palette and type pairing (e.g. editorial serif on paper vs. bold dark tech vs. soft pastel grid). Each direction: name (2–4 words), tone, palette {primary, accent, bg, surface, text} as hex, typography {heading, body font stacks, scale compact|regular|airy}, tokens {radius, spacing, shadow}, layout (e.g. centered 1100px sections / full-bleed hero + 12-col grid), density compact|regular|airy, docx_friendly (true only for Word-bound reports or when asked).',
  '- sections: every section with title, type (hero | summary | content | data | comparison | timeline | cta | footer), brief (what goes in, with the real facts and numbers), layout (hero_split | hero_typographic | hero_image_overlay | feature_grid | stats_row | big_number | two_column | comparison_table | timeline | steps | quote_band | gallery | faq | cta_band | prose; vary across sections, adjacent sections never repeat a pattern), optional image_queries.',
  '- meta: title, language (of the page, following the conversation), audience, length (e.g. one page / 3–5 screens / long-form).',
  '- context: the reference material for the page writer, who sees only the brief and this field, never the conversation: every fact, figure, quote, product detail, search finding and source the page must use, in full. Never invent facts.',
  '- Questionnaire answers and anything the user asked for are binding. A redo note overrides the previous proposal.',
  '- restyle: keep the existing sections (read them from the page), propose only a new style. extract: describe the page as it is, one direction is enough.',
  '',
  '## Output',
  'Reply with one JSON object only, no markdown fences, no commentary: {"core_hook", "style", "alternatives", "sections", "meta", "context"}. Strings use straight double quotes with inner quotes escaped.',
].join('\n')

const MODE_TASK: Record<PlanMode, string> = {
  new: 'A new page in an empty document: propose the full brief.',
  redesign: 'The user asked to rebuild the whole page: propose a full new brief.',
  restyle:
    'Restyle only: keep the structure and sections of the current page, propose a new style direction (plus alternatives).',
  extract:
    'Extract the brief of the current page as it is: one direction, its sections, no invention.',
}

function clip(text: string, cap: number): string {
  return text.length > cap ? `${text.slice(0, cap)}\n… (truncated)` : text
}

/** Flatten the agent history into a transcript the writer can read; the newest turns win the budget. */
export function transcriptOf(messages: readonly AgentMessage[]): string {
  const turns: string[] = []
  for (const m of messages) {
    if (m.role === 'user') turns.push(`### User\n${clip(m.text, TURN_CAP)}`)
    else if (m.role === 'assistant') {
      const calls = m.toolCalls?.map((c) => `[called ${c.name} ${JSON.stringify(c.input)}]`) ?? []
      const body = [m.text.trim(), ...calls].filter(Boolean).join('\n')
      if (body) turns.push(`### Assistant\n${clip(body, TURN_CAP)}`)
    } else {
      for (const r of m.results) turns.push(`### Tool result\n${clip(r.output, TURN_CAP)}`)
    }
  }
  const kept: string[] = []
  let size = 0
  for (let i = turns.length - 1; i >= 0; i--) {
    const turn = turns[i]!
    if (size + turn.length > TRANSCRIPT_CAP && kept.length) break
    kept.unshift(turn)
    size += turn.length
  }
  if (kept.length < turns.length) kept.unshift('… (earlier turns omitted)')
  return kept.join('\n\n')
}

export function buildBriefWriterRequest(
  spec: BriefPlanSpec,
  messages: readonly AgentMessage[],
  langDirective: string,
): { system: string; user: string } {
  const notes = spec.notes?.trim()
    ? `\n\n## Notes from the assistant\n${clip(spec.notes, NOTES_CAP)}`
    : ''
  const page =
    spec.page?.trim() && (spec.mode === 'restyle' || spec.mode === 'extract')
      ? `\n\n## Current page (head and start of body)\n${clip(spec.page, PAGE_CAP)}`
      : ''
  return {
    system: SYSTEM + langDirective,
    user: `## Conversation\n${transcriptOf(messages)}\n\n## Task\n${MODE_TASK[spec.mode]}${notes}${page}\n\nWrite the brief JSON now.`,
  }
}

/** Slice the object between the first and last brace, dropping fences and chatter. */
export function extractBriefJson(raw: string): { text: string; complete: boolean } {
  const start = raw.indexOf('{')
  if (start < 0) return { text: '', complete: false }
  const end = raw.lastIndexOf('}')
  const text = end > start ? raw.slice(start, end + 1) : raw.slice(start)
  return { text, complete: text.endsWith('}') && parseBriefJson(text) !== null }
}

export function parseBriefJson(text: string): Record<string, unknown> | null {
  for (const candidate of [text, text.replace(/,\s*([}\]])/g, '$1')]) {
    try {
      const parsed: unknown = JSON.parse(candidate)
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed))
        return parsed as Record<string, unknown>
    } catch {
      /* try the next repair */
    }
  }
  return null
}

export interface StreamBriefOptions {
  transport: AgentTransport
  system: string
  user: string
  signal?: AbortSignal
}

/** One tool-less streaming request; resolves with the parsed brief or the reason it failed. */
export async function streamBrief(opts: StreamBriefOptions): Promise<BriefPlanResult> {
  const outcome = await streamText({
    transport: opts.transport,
    system: opts.system,
    user: opts.user,
    signal: opts.signal,
    maxChars: BRIEF_MAX_CHARS,
    extract: extractBriefJson,
  })
  if (outcome.status === 'empty') return { ok: false, error: outcome.error }
  const raw = parseBriefJson(outcome.text)
  if (raw) return { ok: true, raw }
  return {
    ok: false,
    error:
      outcome.status === 'partial'
        ? `${outcome.reason}${outcome.error ? `: ${outcome.error}` : ''}`
        : 'the reply was not a JSON object',
  }
}
