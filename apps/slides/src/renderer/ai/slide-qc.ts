/**
 * Post-generation layout QC: each cloud-generated page gets one focused vision pass —
 * screenshot + element inventory → a restricted agent fixes objective layout defects
 * with execute_slide_script. Runs in its own AgentLoop per page (fresh context, so the
 * QC cost doesn't ride on the main conversation), orchestrated by AiPanel.
 */
import {
  AgentLoop,
  type AgentImage,
  type AgentSkill,
  type AgentTransport,
} from '@genoffice/agent-core'
import { auditSlideLayout } from './layout-audit'
import { createSlidesSkill, formatSlideDump, type DeckAccess } from './slides-skill'

/** Kill switch: localStorage 'ai-slides-qc' = '0' disables the automatic pass */
export function isQcEnabled(): boolean {
  return localStorage.getItem('ai-slides-qc') !== '0'
}

/** Cost ceiling per generation run — beyond this the tail pages are skipped (reported to the user) */
export const QC_MAX_PAGES = 20

/**
 * Pages produced by a generateFromHtml/regenerateSlide call, as 0-based indexes.
 * replace lands a whole new deck; append starts at appendedFrom; insert_at/replace_at touch one page.
 */
export function generatedPageRange(
  mode: 'replace' | 'append' | 'insert_at' | 'replace_at',
  r: { pages?: number; appendedFrom?: number; insertedIndex?: number },
): number[] {
  const total = r.pages ?? 0
  switch (mode) {
    case 'replace':
      return Array.from({ length: total }, (_, i) => i)
    case 'append': {
      const from = r.appendedFrom ?? 0
      return Array.from({ length: Math.max(0, total - from) }, (_, i) => from + i)
    }
    case 'insert_at':
    case 'replace_at':
      return typeof r.insertedIndex === 'number' ? [r.insertedIndex] : []
  }
}

/**
 * Fold one landing's pages into the run's pending-QC set.
 * replace discards earlier pendings (whole new deck); insert_at shifts pendings at/after
 * the insertion point before adding it, keeping indexes valid.
 */
export function mergeQcPages(
  prev: number[],
  mode: 'replace' | 'append' | 'insert_at' | 'replace_at',
  r: { pages?: number; appendedFrom?: number; insertedIndex?: number },
): number[] {
  const range = generatedPageRange(mode, r)
  const sortDedupe = (pages: number[]) => [...new Set(pages)].sort((a, b) => a - b)
  switch (mode) {
    case 'replace':
      return range
    case 'append':
    case 'replace_at':
      return sortDedupe([...prev, ...range])
    case 'insert_at': {
      const at = r.insertedIndex
      if (typeof at !== 'number') return prev
      return sortDedupe([...prev.map((p) => (p >= at ? p + 1 : p)), at])
    }
  }
}

/** Only the two tools the QC pass needs: fresh geometry reads + atomic layout scripts */
const QC_TOOL_ALLOWLIST = new Set(['read_slide', 'execute_slide_script'])

const QC_SYSTEM_PROMPT = `You are a slide layout QA and polish fixer. Each request gives you ONE slide: a rendered screenshot (attached image) and an element inventory (ids, geometry, colors, text — the same ids the tools accept).

First fix objective defects:
- text overflowing its box, colliding with a neighbor, or clipped by the canvas edge
- elements overlapping unintentionally (a text block over another text block; content under an image)
- unreadable contrast (text color too close to what it sits on)
- distorted or badly cropped images

Then apply a restrained professional polish when the screenshot clearly needs it:
- establish a clear visual hierarchy between title, subtitle, body, captions, and key figures
- align related elements to shared edges or centers; make columns, cards, and repeated items consistent
- normalize spacing and padding so groups are visually connected and sections have breathing room
- improve typography using the page's existing font family: adjust font size, weight, line height, and text-box size for readability
- rebalance whitespace and visual weight by moving or resizing existing elements
- improve text contrast only when needed, using colors already present on the page

Use execute_slide_script and batch every change for this page into as few calls as possible; call read_slide first if you need fresher geometry than the inventory. Preserve the page's content, visual identity, and intended composition. Prefer a small coordinated set of high-confidence changes over many cosmetic tweaks. After editing, use the tool's layout-audit feedback to correct any new defect.

STRICTLY FORBIDDEN: regenerating or redesigning the page, changing the theme or font family, rewriting copy, changing facts or numbers, adding or deleting elements, introducing a new color palette, or touching elements without a clear visual benefit. When the page is already clean, balanced, and readable, make NO tool call.

Final reply: one short line (under 15 words) stating what you fixed, or exactly "OK" if nothing needed fixing.`

export interface QcPageResult {
  /** page still exists and the pass ran */
  ok: boolean
  /** at least one mutating tool call was applied */
  edited: boolean
  /** model's final one-liner ('OK' when clean) */
  reply: string
  /** deterministic audit issue counts before/after (rollback signal: after > before) */
  preIssues: number
  postIssues: number
  error?: string
}

export interface QcPageOptions {
  access: DeckAccess
  transport: AgentTransport
  pageIndex: number
  /** pixelRatio-1 PNG of the page's current rendering; null runs a geometry-only pass */
  screenshot: AgentImage | null
  systemSuffix?: () => string
  signal?: AbortSignal
}

/** Wrap createSlidesSkill with the QC system prompt and the two-tool allowlist (executor shared) */
export function createSlideFixSkill(access: DeckAccess): AgentSkill {
  const full = createSlidesSkill(access)
  return {
    id: 'slides-qc',
    systemPrompt: QC_SYSTEM_PROMPT,
    tools: full.tools.filter((tool) => QC_TOOL_ALLOWLIST.has(tool.name)),
    executeTool: full.executeTool,
  }
}

function buildQcInstruction(pageIndex: number, dump: string, issues: string[]): string {
  const auditStr = issues.length
    ? `Deterministic geometry audit already flags:\n${issues.map((s) => `- ${s}`).join('\n')}\n(These are hints — the screenshot is the ground truth; it may show more or reveal a flagged item is fine.)`
    : 'The deterministic geometry audit found nothing — trust the screenshot for visual defects it cannot measure (contrast, alignment, crowding).'
  return `Slide ${pageIndex + 1} (slideIndex ${pageIndex}) was just auto-generated. The attached image is its current rendering.

Element inventory:
${dump}

${auditStr}

Inspect the screenshot, fix objective layout defects first, then apply restrained professional polish when clearly beneficial.`
}

/**
 * One page, one focused QC run. The caller owns history batching (rollback via
 * aiSnapshotRestore) and deciding what to do with the result.
 */
export function qcSlidePage(opts: QcPageOptions): Promise<QcPageResult> {
  const { access, transport, pageIndex, screenshot, systemSuffix, signal } = opts
  const slide = access.getSlides()[pageIndex]
  if (!slide) {
    return Promise.resolve({
      ok: false,
      edited: false,
      reply: '',
      preIssues: 0,
      postIssues: 0,
      error: `slideIndex ${pageIndex} out of range`,
    })
  }
  const preIssues = auditSlideLayout(slide)
  const instruction = buildQcInstruction(pageIndex, formatSlideDump(slide), preIssues)

  return new Promise((resolve) => {
    let edited = false
    const finish = (r: { reply: string; error?: string }) => {
      const after = access.getSlides()[pageIndex]
      resolve({
        ok: true,
        edited,
        reply: r.reply.trim(),
        preIssues: preIssues.length,
        postIssues: after ? auditSlideLayout(after).length : 0,
        ...(r.error !== undefined ? { error: r.error } : {}),
      })
    }
    const loop = new AgentLoop({
      transport,
      skill: createSlideFixSkill(access),
      // audit feedback inside execute_slide_script output drives at most a couple of fix rounds
      maxTurns: 6,
      ...(systemSuffix ? { systemSuffix } : {}),
      events: {
        onToolExecuted: ({ execution }) => {
          if (execution.mutated) edited = true
        },
        onDone: ({ text }) => finish({ reply: text }),
        onError: (error) => finish({ reply: '', error }),
      },
    })
    const onAbort = () => loop.cancel()
    signal?.addEventListener('abort', onAbort, { once: true })
    loop.run(instruction, screenshot ? [screenshot] : [])
  })
}
