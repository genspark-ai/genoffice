import type { AgentSkill, AgentToolCall, AgentToolDef, ToolExecution } from '@genoffice/agent-core'
import { t } from '../i18n/locale'
import type { ElementEntry, ParseMap } from '../document/parse-map'
import type { HtmlOp, OpError } from '../document/ops'
import { lineOf } from '../document/match'
import {
  briefPinEdit,
  briefSummary,
  injectBrief,
  parseBrief,
  type Brief,
  type BriefStyle,
} from '../document/brief'
import type { PageWriteResult, PageWriteSpec } from './page-writer'

export const CONTEXT_MAX_CHARS = 8000
export const OUTLINE_MAX_LINES = 80
export const READ_PAGE_CHARS = 12000
export const SELECTION_MAX_CHARS = 4000
const PREVIEW_CHARS = 60

/** What the skill needs from the app; every getter reads live state. */
export interface HtmlDocAccess {
  getText(): string
  getVersion(): number
  getMap(): ParseMap
  /** version of the last edit that did not come from the AI (typing, ribbon) */
  getLastManualVersion(): number
  getFilePath(): string | null
  /** element the user selected in the preview / source, if any */
  getSelectedSid(): number | null
  /** compile + apply atomically; returns the applied patch ranges (post-edit offsets) or errors */
  applyOps(
    ops: HtmlOp[],
    label: string,
  ): { ok: true; ranges: Array<[number, number]> } | { ok: false; errors: OpError[] }
  replaceAll(html: string, label: string): void
  /** questionnaire card; resolves with the user's answers or cancelled when skipped */
  askClarification?(questions: ClarifyQuestion[]): Promise<{ answers: string; cancelled?: boolean }>
  /** brief card; resolves with the confirmed (possibly edited) brief, a request for another one, or cancelled */
  confirmBrief?(brief: Brief): Promise<BriefDecision>
  /** whole-page generation as its own streamed request (see page-writer.ts); the app previews the draft and asks about partial pages */
  writePage?(spec: PageWriteSpec, signal?: AbortSignal): Promise<PageWriteResult>
  /** the instruction of the current run, handed to the page writer (which never sees the conversation) */
  getInstruction?(): string
}

export type PlanMode = 'new' | 'redesign' | 'restyle' | 'extract'

export interface ClarifyQuestion {
  id: string
  label: string
  description?: string
  /** ≤5 options; the card appends "decide for me" and a free-text "other" itself */
  options: string[]
  multi?: boolean
}

export type BriefDecision =
  { kind: 'confirmed'; brief: Brief } | { kind: 'redo'; note: string } | { kind: 'cancelled' }

const str = (v: unknown) => (v === undefined || v === null ? undefined : String(v))

function coerceStyle(raw: unknown): BriefStyle | null {
  const style = (raw ?? {}) as Record<string, unknown>
  if (typeof style !== 'object' || !String(style.tone ?? '').trim()) return null
  const palette = (style.palette ?? {}) as Record<string, unknown>
  const typography = (style.typography ?? {}) as Record<string, unknown>
  const tokens = (style.tokens ?? {}) as Record<string, unknown>
  return {
    name: str(style.name)?.trim() || undefined,
    tone: String(style.tone).trim(),
    palette: {
      primary: str(palette.primary),
      accent: str(palette.accent),
      bg: str(palette.bg),
      surface: str(palette.surface),
      text: str(palette.text),
    },
    typography: {
      heading: str(typography.heading),
      body: str(typography.body),
      scale: str(typography.scale),
    },
    tokens: {
      radius: str(tokens.radius),
      spacing: str(tokens.spacing),
      shadow: str(tokens.shadow),
    },
    layout: str(style.layout),
    density: str(style.density),
    docx_friendly: Boolean(style.docx_friendly),
  }
}

const MAX_ALTERNATIVES = 2

/** a missing/unknown mode is read from the document: empty → new page, otherwise only pin the brief */
function planMode(raw: unknown, text: string): PlanMode {
  if (raw === 'new' || raw === 'redesign' || raw === 'restyle' || raw === 'extract') return raw
  return text.trim() ? 'extract' : 'new'
}

function coerceBrief(input: Record<string, unknown>): Brief | string {
  const coreHook = String(input.core_hook ?? '').trim()
  const sections = Array.isArray(input.sections) ? input.sections : []
  if (!coreHook) return 'core_hook is required'
  const style = coerceStyle(input.style)
  if (!style) return 'style.tone is required'
  if (sections.length === 0) return 'sections must be non-empty'
  const alternatives = (Array.isArray(input.alternatives) ? input.alternatives : [])
    .map(coerceStyle)
    .filter((s): s is BriefStyle => s !== null)
    .slice(0, MAX_ALTERNATIVES)
  const meta = (input.meta ?? {}) as Record<string, unknown>
  return {
    core_hook: coreHook,
    style,
    alternatives: alternatives.length ? alternatives : undefined,
    sections: sections.map((raw) => {
      const sec = (raw ?? {}) as Record<string, unknown>
      return {
        title: String(sec.title ?? '').trim() || 'Untitled',
        type: str(sec.type),
        brief: String(sec.brief ?? '').trim(),
        layout: str(sec.layout),
        image_queries: Array.isArray(sec.image_queries) ? sec.image_queries.map(String) : undefined,
      }
    }),
    meta: {
      title: str(meta.title),
      language: str(meta.language),
      audience: str(meta.audience),
      length: str(meta.length),
    },
    version: 1,
  }
}

const STYLE_SCHEMA = {
  type: 'object',
  required: ['tone'],
  properties: {
    name: { type: 'string', description: 'short direction name, 2–4 words' },
    tone: {
      type: 'string',
      description: 'e.g. executive / editorial / playful / technical / luxurious / brutalist',
    },
    palette: {
      type: 'object',
      properties: {
        primary: { type: 'string', description: 'hex' },
        accent: { type: 'string', description: 'hex; one accent system for the whole page' },
        bg: { type: 'string', description: 'hex page background; dark or saturated is legitimate' },
        surface: { type: 'string', description: 'hex card / panel background' },
        text: { type: 'string', description: 'hex; must contrast with bg' },
      },
    },
    typography: {
      type: 'object',
      properties: {
        heading: { type: 'string', description: 'font-family stack' },
        body: { type: 'string', description: 'font-family stack' },
        scale: { type: 'string', description: 'compact | regular | airy' },
      },
    },
    tokens: {
      type: 'object',
      properties: {
        radius: { type: 'string', description: 'CSS length, e.g. 0 / 6px / 16px' },
        spacing: { type: 'string', description: 'base unit, e.g. 8px' },
        shadow: { type: 'string', description: 'CSS box-shadow or none' },
      },
    },
    layout: {
      type: 'string',
      description:
        'e.g. centered 1100px sections / full-bleed hero + 12-col grid / single column 760px',
    },
    density: { type: 'string', enum: ['compact', 'regular', 'airy'] },
    docx_friendly: {
      type: 'boolean',
      description:
        'default false; true only for reports meant for Word export or when the user asks — then no absolute positioning, gradients or transforms carrying meaning',
    },
  },
} as const

/** Skip structural/no-content elements in the outline; the model can still address them via read_source */
const OUTLINE_SKIP = new Set([
  'html',
  'head',
  'body',
  'meta',
  'link',
  'script',
  'style',
  'br',
  'wbr',
  'template',
  'noscript',
])

function textPreview(text: string, e: ElementEntry): string {
  const inner = text
    .slice(e.inner[0], e.inner[1])
    .replace(/<[^>]*>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
  return inner.length > PREVIEW_CHARS ? `${inner.slice(0, PREVIEW_CHARS)}…` : inner
}

function attrsPreview(text: string, e: ElementEntry): string {
  const tag = text.slice(e.startTag[0], e.startTag[1])
  const id = /\sid\s*=\s*["']([^"']+)["']/i.exec(tag)?.[1]
  const cls = /\sclass\s*=\s*["']([^"']+)["']/i.exec(tag)?.[1]
  const src = /\s(?:src|href)\s*=\s*["']([^"']+)["']/i.exec(tag)?.[1]
  const parts: string[] = []
  if (id) parts.push(`#${id}`)
  if (cls) parts.push(`.${cls.trim().split(/\s+/).slice(0, 3).join('.')}`)
  if (src && (e.tag === 'img' || e.tag === 'a' || e.tag === 'source'))
    parts.push(`(${src.slice(0, 40)})`)
  return parts.join('')
}

export function buildOutline(
  text: string,
  map: ParseMap,
  opts: { depth?: number; fromSid?: number; maxLines?: number } = {},
): string {
  const depthLimit = opts.depth ?? 3
  const root = opts.fromSid !== undefined ? map.bySid.get(opts.fromSid) : undefined
  const maxLines = opts.maxLines ?? OUTLINE_MAX_LINES
  const lines: string[] = []
  let skipped = 0
  const baseDepth = root ? root.depth : 0
  for (const e of map.elements) {
    if (root && !(e.range[0] >= root.range[0] && e.range[1] <= root.range[1])) continue
    if (OUTLINE_SKIP.has(e.tag) && e.tag !== 'style' && e.tag !== 'script') continue
    if (e.tag === 'style' || e.tag === 'script') {
      // keep one line for each so the model knows they exist
    }
    const rel = e.depth - baseDepth
    if (rel > depthLimit) {
      skipped++
      continue
    }
    if (lines.length >= maxLines) {
      skipped++
      continue
    }
    const l1 = lineOf(text, e.range[0])
    const l2 = lineOf(text, Math.max(e.range[0], e.range[1] - 1))
    const where = l1 === l2 ? `L${l1}` : `L${l1}-L${l2}`
    const indent = '  '.repeat(Math.max(0, rel - 1))
    const preview = OUTLINE_SKIP.has(e.tag) ? '' : textPreview(text, e)
    lines.push(
      `${where.padEnd(11)} sid=${String(e.sid).padEnd(4)} ${indent}${e.tag}${attrsPreview(text, e)}${preview ? `  "${preview}"` : ''}`,
    )
  }
  if (skipped > 0)
    lines.push(`… (${skipped} more elements; use read_source or get_outline with from_sid / depth)`)
  return lines.join('\n')
}

function numbered(text: string, fromLine: number): string {
  return text
    .split('\n')
    .map((l, i) => `${String(fromLine + i).padStart(5)}| ${l}`)
    .join('\n')
}

export const AGENT_TOOLS: AgentToolDef[] = [
  {
    name: 'get_outline',
    description:
      'Structural outline of the document: one line per element with its line range, stable element id (sid), tag, id/class and a text preview. Use it to find where to read or edit. Narrow with from_sid / depth for large documents.',
    inputSchema: {
      type: 'object',
      properties: {
        depth: { type: 'integer', description: 'Nesting depth to show (default 3)' },
        from_sid: { type: 'integer', description: 'Only this element and its subtree' },
      },
    },
  },
  {
    name: 'read_source',
    description:
      'Read a slice of the HTML source with line numbers. Address by line range, or by sid to read one element. Long results are paged (page starts at 0).',
    inputSchema: {
      type: 'object',
      properties: {
        start_line: { type: 'integer' },
        end_line: { type: 'integer' },
        sid: { type: 'integer' },
        page: { type: 'integer' },
      },
    },
  },
  {
    name: 'apply_ops',
    description:
      'Apply a batch of edits to the HTML source atomically: every op is validated first and either all are applied or none. Ops: ' +
      '`str_replace {old, new, sid?, replace_all?}` exact search-and-replace (old must match once; scope with sid when markup repeats); ' +
      '`replace_element {sid, html}` swap the whole element; `set_inner_html {sid, html}`; `set_text {sid, text}` (text is escaped for you); ' +
      '`insert_html {sid, position: before|after|prepend|append, html}`; `remove {sid}`; `move {sid, position: before|after, ref_sid}`; ' +
      '`set_attr {sid, name, value|null}`; `set_style {sid, styles: {prop: value|null}}` merges inline style. ' +
      'Prefer sid-addressed ops for single elements and str_replace for scattered text or CSS/JS edits. Keep unrelated markup, indentation and attribute order untouched.',
    inputSchema: {
      type: 'object',
      required: ['ops'],
      properties: {
        ops: {
          type: 'array',
          minItems: 1,
          maxItems: 40,
          items: {
            type: 'object',
            required: ['op'],
            properties: {
              op: {
                type: 'string',
                enum: [
                  'str_replace',
                  'replace_element',
                  'set_inner_html',
                  'set_text',
                  'insert_html',
                  'remove',
                  'move',
                  'set_attr',
                  'set_style',
                ],
              },
              old: { type: 'string' },
              new: { type: 'string' },
              replace_all: { type: 'boolean' },
              sid: { type: 'integer' },
              ref_sid: { type: 'integer' },
              html: { type: 'string' },
              text: { type: 'string' },
              position: { type: 'string', enum: ['before', 'after', 'prepend', 'append'] },
              name: { type: 'string' },
              value: { type: ['string', 'null'] },
              styles: { type: 'object', additionalProperties: { type: ['string', 'null'] } },
            },
          },
        },
        summary: { type: 'string', description: 'One short sentence shown to the user' },
      },
    },
  },
  {
    name: 'ask_clarification',
    description:
      '[Before creating a new page] Show a questionnaire card so the user makes the key choices (audience, scenario, tone, focus, length). Ask 2–4 questions that are real trade-offs for THIS topic, ≤5 options each; the card adds "decide for me" and a free-text option itself. Skip it when the request already answers these. Wait for the answers, then call plan_page. Do not repeat the questions in your reply.',
    inputSchema: {
      type: 'object',
      required: ['questions'],
      properties: {
        questions: {
          type: 'array',
          minItems: 1,
          maxItems: 4,
          items: {
            type: 'object',
            required: ['id', 'label', 'options'],
            properties: {
              id: { type: 'string' },
              label: { type: 'string' },
              description: { type: 'string' },
              options: { type: 'array', items: { type: 'string' }, maxItems: 5 },
              multi: { type: 'boolean' },
            },
          },
        },
      },
    },
  },
  {
    name: 'plan_page',
    description:
      '[Always before generating a new page or a full redesign, and when the user asks to restyle or "extract the brief" of an existing page] Propose the brief: the Core Hook (one sentence with tension, ideally a number or contrast), the recommended style direction plus 1–2 named alternatives that differ meaningfully (e.g. editorial serif on paper vs. bold dark tech vs. soft pastel grid), the section list and meta. The user sees the directions as style tiles, picks one, edits fields, and confirms or asks for another brief. With mode "new" or "redesign" the confirmed brief is written into a page by the system right away (you do not write HTML yourself); put every fact, figure and quote the page must use into `context`. Do not repeat the brief in your reply.',
    inputSchema: {
      type: 'object',
      required: ['core_hook', 'style', 'sections', 'mode'],
      properties: {
        mode: {
          type: 'string',
          enum: ['new', 'redesign', 'restyle', 'extract'],
          description:
            'new = empty document, the page is generated after confirmation; redesign = the user asked to rebuild the whole page, generated after confirmation; restyle = keep the structure, you update --brief-* variables with apply_ops afterwards; extract = only pin the inferred brief',
        },
        context: {
          type: 'string',
          description:
            'Reference material for the page writer: facts, figures, quotes, product details and sources gathered from the conversation, attachments and web_search. The writer never sees the conversation, only this.',
        },
        core_hook: { type: 'string' },
        style: { ...STYLE_SCHEMA, description: 'the recommended direction' },
        alternatives: {
          type: 'array',
          maxItems: MAX_ALTERNATIVES,
          items: STYLE_SCHEMA,
          description:
            '1–2 other directions with a different mood, palette and type pairing; the user picks one on the card',
        },
        sections: {
          type: 'array',
          minItems: 1,
          items: {
            type: 'object',
            required: ['title', 'brief'],
            properties: {
              title: { type: 'string' },
              type: {
                type: 'string',
                description:
                  'hero | summary | content | data | comparison | timeline | cta | footer',
              },
              brief: { type: 'string', description: 'what goes in, with real facts and numbers' },
              layout: {
                type: 'string',
                description:
                  'hero_split | hero_typographic | hero_image_overlay | feature_grid | stats_row | big_number | two_column | comparison_table | timeline | steps | quote_band | gallery | faq | cta_band | prose; vary across sections',
              },
              image_queries: { type: 'array', items: { type: 'string' } },
            },
          },
        },
        meta: {
          type: 'object',
          properties: {
            title: { type: 'string' },
            language: { type: 'string' },
            audience: { type: 'string' },
            length: { type: 'string', description: 'e.g. one page / 3–5 screens / long-form' },
          },
        },
      },
    },
  },
  {
    name: 'write_document',
    description:
      '[Content-first documents only: the user wants text written (article, guide, memo, announcement), not a designed page] Hands a plan to the system document writer, which writes the whole document as a clean single-column page. Give the title, a plan (outline, key points, length, voice) and the reference material in context. Everything else — new designed pages and full redesigns — goes through plan_page; single edits use apply_ops.',
    inputSchema: {
      type: 'object',
      required: ['plan'],
      properties: {
        title: { type: 'string' },
        plan: {
          type: 'string',
          description: 'outline and key points the document must cover, length, voice',
        },
        context: {
          type: 'string',
          description: 'facts, figures and source text the document must draw on',
        },
      },
    },
  },
]

const STALE_MESSAGE =
  'The document changed since you last looked at it (the user edited it). Call get_outline or read_source again and re-plan before editing.'

export function createHtmlSkillCore(access: HtmlDocAccess): {
  buildContext(): string
  executeTool(call: AgentToolCall, signal?: AbortSignal): ToolExecution | Promise<ToolExecution>
} {
  /** version the model last saw; index/sid-addressed writes refuse when the user edited after it */
  let seenVersion = -1
  const markSeen = () => {
    seenVersion = access.getVersion()
  }
  const stale = () => access.getLastManualVersion() > seenVersion
  let currentSignal: AbortSignal | undefined

  const landPage = async (spec: PageWriteSpec, preface: string): Promise<ToolExecution> => {
    const summary = t('aiToolWritePage')
    if (!access.writePage)
      return { output: 'page generation is not available here', isError: true, summary }
    const result = await access.writePage(spec, currentSignal)
    if (!result.ok || !result.html)
      return {
        output: `${preface}\nThe page writer produced nothing (${result.error ?? 'no output'}); the document is unchanged. Tell the user briefly and offer to try again.`,
        isError: true,
        summary: t('aiToolWritePageFailed'),
      }
    access.replaceAll(result.html, 'generate')
    markSeen()
    const lines = result.html.split('\n').length
    const sections = [...result.html.matchAll(/data-section="([^"]+)"/g)].map((m) => m[1])
    const note = result.truncated
      ? ' The stream ended early, so the page is INCOMPLETE (the user chose to keep it): the tail is missing. Say so and offer to finish the missing sections with apply_ops.'
      : ''
    return {
      output: `${preface}\nPage written by the system (${result.html.length} chars, ${lines} lines${sections.length ? `; sections: ${sections.join(', ')}` : ''}). Version ${access.getVersion()}.${note}\nReply with one or two sentences describing the page; do not paste HTML.`,
      mutated: true,
      summary: result.truncated ? t('aiToolWritePagePartial') : summary,
    }
  }

  /** write the brief meta into the open document as one targeted edit; true when the document changed */
  const pinBrief = (brief: Brief): boolean => {
    const current = access.getText()
    const edit = briefPinEdit(current, brief)
    if (edit) {
      const r = access.applyOps([{ op: 'str_replace', old: edit.old, new: edit.new }], 'pin brief')
      if (r.ok) {
        markSeen()
        return true
      }
    } else if (parseBrief(current)) return false
    // whole-document fallback: applyOps flushes pending live edits first, so re-read rather than reuse the snapshot
    access.replaceAll(injectBrief(access.getText(), brief), 'pin brief')
    markSeen()
    return true
  }

  return {
    buildContext() {
      markSeen()
      const text = access.getText()
      const map = access.getMap()
      const path = access.getFilePath()
      const title = /<title[^>]*>([^<]*)<\/title>/i.exec(text)?.[1]?.trim()
      const lines = text.split('\n').length
      const head = [
        '## Document',
        `file: ${path ? path.replace(/^.*[/\\]/, '') : '(untitled, not saved yet)'}  (${text.length.toLocaleString()} chars, ${lines} lines, version ${access.getVersion()})`,
        title ? `title: ${title}` : '',
        text.trim()
          ? ''
          : 'The document is empty: plan_page (mode new) or write_document creates it.',
      ].filter(Boolean)
      let outline = buildOutline(text, map, { depth: 2, maxLines: 60 })
      const selectedSid = access.getSelectedSid()
      const selected = selectedSid !== null ? map.bySid.get(selectedSid) : undefined
      let selection = ''
      if (selected) {
        const source = text.slice(selected.range[0], selected.range[1])
        const clipped =
          source.length > SELECTION_MAX_CHARS
            ? `${source.slice(0, SELECTION_MAX_CHARS)}\n… (truncated; read_source sid=${selected.sid} for the rest)`
            : source
        const l1 = lineOf(text, selected.range[0])
        const l2 = lineOf(text, Math.max(selected.range[0], selected.range[1] - 1))
        selection = `\n\n## User selection\nsid=${selected.sid} <${selected.tag}> L${l1}-L${l2}\nEdit/rewrite-style requests apply to this element by default.\n\`\`\`html\n${clipped}\n\`\`\``
      }
      const brief = parseBrief(text)
      const briefBlock = brief
        ? `\n\n## Brief (pinned in <meta name="genoffice:brief">)\n${briefSummary(brief)}\nNew content follows this core hook and style; a restyle starts by updating the --brief-* variables and this meta via plan_page.`
        : ''
      const budget =
        CONTEXT_MAX_CHARS - head.join('\n').length - selection.length - briefBlock.length - 40
      if (outline.length > budget)
        outline = `${outline.slice(0, Math.max(0, budget))}\n… (outline truncated; use get_outline)`
      return `${head.join('\n')}\n\noutline (sid = stable element id for apply_ops):\n${outline}${selection}${briefBlock}`
    },

    executeTool(call, signal) {
      currentSignal = signal
      const text = access.getText()
      const map = access.getMap()
      switch (call.name) {
        case 'get_outline': {
          markSeen()
          const depth = Number.isInteger(call.input.depth) ? Number(call.input.depth) : 3
          const fromSid = Number.isInteger(call.input.from_sid)
            ? Number(call.input.from_sid)
            : undefined
          if (fromSid !== undefined && !map.bySid.has(fromSid)) {
            return {
              output: `sid ${fromSid} does not exist`,
              isError: true,
              summary: t('aiToolOutline'),
            }
          }
          const out = buildOutline(text, map, { depth, fromSid, maxLines: 400 })
          return { output: out || '(empty document)', summary: t('aiToolOutline') }
        }
        case 'read_source': {
          markSeen()
          const page = Number.isInteger(call.input.page) ? Math.max(0, Number(call.input.page)) : 0
          let from = 0
          let to = text.length
          if (Number.isInteger(call.input.sid)) {
            const e = map.bySid.get(Number(call.input.sid))
            if (!e)
              return {
                output: `sid ${String(call.input.sid)} does not exist`,
                isError: true,
                summary: t('aiToolReadSource'),
              }
            ;[from, to] = e.range
          } else if (
            Number.isInteger(call.input.start_line) ||
            Number.isInteger(call.input.end_line)
          ) {
            const lines = text.split('\n')
            const s = Math.max(1, Number(call.input.start_line) || 1)
            const e = Math.min(lines.length, Number(call.input.end_line) || lines.length)
            from = lines.slice(0, s - 1).join('\n').length + (s > 1 ? 1 : 0)
            to = lines.slice(0, e).join('\n').length
          }
          const slice = text.slice(from, to)
          const pages = Math.max(1, Math.ceil(slice.length / READ_PAGE_CHARS))
          const chunk = slice.slice(page * READ_PAGE_CHARS, (page + 1) * READ_PAGE_CHARS)
          const startLine = lineOf(text, from + page * READ_PAGE_CHARS)
          const body = numbered(chunk, startLine)
          const footer =
            pages > 1 ? `\n(page ${page + 1} of ${pages}; pass page=${page + 1} for more)` : ''
          return { output: body + footer, summary: t('aiToolReadSource') }
        }
        case 'apply_ops': {
          const ops = call.input.ops
          if (!Array.isArray(ops) || ops.length === 0) {
            return {
              output: 'ops must be a non-empty array',
              isError: true,
              summary: t('aiToolApplyOpsFailed'),
            }
          }
          if (stale())
            return { output: STALE_MESSAGE, isError: true, summary: t('aiToolApplyOpsFailed') }
          const label =
            typeof call.input.summary === 'string' && call.input.summary
              ? call.input.summary
              : `apply_ops (${ops.length})`
          const result = access.applyOps(ops as HtmlOp[], label)
          if (!result.ok) {
            const lines = result.errors.map(
              (e) =>
                `Op ${e.index + 1}/${ops.length} FAILED (${e.kind.toUpperCase()}): ${e.message}`,
            )
            return {
              output: `0 of ${ops.length} ops applied.\n${lines.join('\n')}`,
              isError: true,
              summary: t('aiToolApplyOpsFailed'),
            }
          }
          markSeen()
          const after = access.getText()
          const where = result.ranges.slice(0, 12).map(([a, b]) => {
            const l1 = lineOf(after, a)
            const l2 = lineOf(after, Math.max(a, b - 1))
            return l1 === l2 ? `L${l1}` : `L${l1}-L${l2}`
          })
          const health = access.getMap().errorCount - map.errorCount
          const warn =
            health > 5
              ? `\nWARNING: the edit introduced ${health} new HTML parse recoveries (unbalanced tags?). Check the preview.`
              : ''
          return {
            output: `Applied ${ops.length} op(s) at ${where.join(', ')}. Document is now version ${access.getVersion()}.${warn}`,
            mutated: true,
            summary: t('aiToolApplyOps', { n: ops.length }),
          }
        }
        case 'ask_clarification': {
          if (!access.askClarification)
            return {
              output:
                'questionnaire cards are not available here; decide sensible defaults and continue',
              isError: true,
              summary: t('aiToolClarify'),
            }
          const raw = Array.isArray(call.input.questions) ? call.input.questions : []
          const questions: ClarifyQuestion[] = raw
            .map((q: Record<string, unknown>, i: number) => ({
              id: String(q.id ?? `q${i + 1}`),
              label: String(q.label ?? '').trim(),
              description: q.description ? String(q.description) : undefined,
              options: Array.isArray(q.options)
                ? q.options.map((o: unknown) => String(o)).slice(0, 5)
                : [],
              multi: !!q.multi,
            }))
            .filter((q) => q.label && q.options.length > 0)
            .slice(0, 4)
          if (questions.length === 0)
            return {
              output: 'questions must be non-empty and every question needs options',
              isError: true,
              summary: t('aiToolClarify'),
            }
          return access.askClarification(questions).then((r) =>
            r.cancelled
              ? {
                  output:
                    'The user skipped the questionnaire. Decide the core hook and style yourself, then call plan_page.',
                  summary: t('aiToolClarifySkipped'),
                }
              : {
                  output: `User answers:\n${r.answers}\nNow call plan_page with a brief that follows these choices.`,
                  summary: t('aiToolClarify'),
                },
          )
        }
        case 'plan_page': {
          if (!access.confirmBrief)
            return {
              output:
                'brief cards are not available here; describe the brief in your reply instead',
              isError: true,
              summary: t('aiToolPlan'),
            }
          const brief = coerceBrief(call.input)
          if (typeof brief === 'string')
            return { output: brief, isError: true, summary: t('aiToolPlanRejected') }
          const mode = planMode(call.input.mode, text)
          const context = str(call.input.context)
          return access.confirmBrief(brief).then(async (decision) => {
            if (decision.kind === 'cancelled')
              return {
                output:
                  'The user dismissed the brief without confirming. Ask what they would like to change, or stop.',
                summary: t('aiToolPlanRejected'),
              }
            if (decision.kind === 'redo')
              return {
                output: `The user asked for a different brief${decision.note ? `: ${decision.note}` : ''}. Call plan_page again with a revised proposal.`,
                summary: t('aiToolPlanRejected'),
              }
            const edited = decision.brief.user_edited?.length
              ? `\nFields the user edited (keep them exactly): ${decision.brief.user_edited.join(', ')}`
              : ''
            // the user may have typed while the card was open: judge the live document, not the pre-card snapshot
            if (mode === 'new' || mode === 'redesign') {
              if (access.getText().trim() && stale())
                return { output: STALE_MESSAGE, isError: true, summary: t('aiToolPlanRejected') }
              return landPage(
                {
                  kind: 'design',
                  brief: decision.brief,
                  instruction: access.getInstruction?.() ?? '',
                  context,
                },
                `Brief confirmed.${edited}`,
              )
            }
            if (stale())
              return { output: STALE_MESSAGE, isError: true, summary: t('aiToolPlanRejected') }
            const pinned = pinBrief(decision.brief)
            return {
              output: `Brief confirmed (${mode}).${edited}\n${briefSummary(decision.brief)}\n${
                pinned
                  ? 'The brief is now pinned in <head> (meta name="genoffice:brief"). '
                  : 'The pinned brief was already current. '
              }${
                mode === 'restyle'
                  ? 'Now update the --brief-* variables and related CSS rules with apply_ops (str_replace inside <style>); keep the structure and text.'
                  : 'No page is generated in extract mode.'
              }`,
              mutated: pinned,
              summary: t('aiToolPlan', { n: decision.brief.sections.length }),
            }
          })
        }
        case 'write_document': {
          const plan = String(call.input.plan ?? '').trim()
          if (!plan)
            return {
              output: 'plan must be non-empty',
              isError: true,
              summary: t('aiToolWritePage'),
            }
          if (text.trim() && stale())
            return { output: STALE_MESSAGE, isError: true, summary: t('aiToolWritePage') }
          return landPage(
            {
              kind: 'content',
              title: str(call.input.title),
              plan,
              instruction: access.getInstruction?.() ?? '',
              context: str(call.input.context),
            },
            '',
          )
        }
        default:
          return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
    },
  }
}

export function createHtmlSkill(access: HtmlDocAccess, systemPrompt: string): AgentSkill {
  const core = createHtmlSkillCore(access)
  return {
    id: 'html',
    systemPrompt,
    tools: AGENT_TOOLS,
    buildContext: () => core.buildContext(),
    executeTool: (call, signal) => core.executeTool(call, signal),
  }
}
