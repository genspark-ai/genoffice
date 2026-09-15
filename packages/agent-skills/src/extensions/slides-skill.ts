/**
 * slides-skill extension — GenOffice PowerPoint/slides AI tools as a pi extension.
 *
 * Mirrors docs-skill.ts / sheets-skill.ts shape:
 *   - `SlidesEditor` interface abstracts the deck runtime
 *   - One `defineTool()` per migrated tool
 *   - `createSlidesSkillExtension()` is the entry point for createOfficeSession
 *
 * Phase 1 (W7) ships: read_slide, plan_deck, execute_slide_script, regenerate_slide.
 * The 11 other tools (image/clipart/web) follow in post-W7 work.
 */

import { Type, type Static } from "@earendil-works/pi-ai";
import { defineTool, type ExtensionAPI } from "@earendil-works/pi-coding-agent";
import type { ReactUIAdapter } from "@genoffice/agent-runtime";

// ============================================================================
// SlidesEditor contract
// ============================================================================

export interface SlideSummary {
  index: number;
  title: string;
  layoutName: string;
  hasImages: boolean;
  hasCharts: boolean;
}

export interface SlideContent {
  index: number;
  title: string;
  body: string;
  notes: string;
  layoutName: string;
  /** Slot descriptors (position + content for placeholders). */
  slots: Array<{ kind: string; text?: string; imageRef?: string }>;
}

export interface DeckPlan {
  title: string;
  audience: string;
  pages: Array<{ title: string; keyPoints: string[]; layoutHint?: string }>;
}

export interface SlideScript {
  /** Operations in order; format depends on host runtime. */
  operations: ReadonlyArray<unknown>;
}

export interface SlidesEditor {
  // ---- Read API ----
  /** High-level deck summary (page count, layouts, media presence). */
  getDeckSummary(): { pageCount: number; title: string; pages: SlideSummary[] };
  /** Read a single slide's full content (title, body, notes, slots). */
  readSlide(index: number): SlideContent | null;

  // ---- Write API ----
  /** Replace a slide's content from a generated plan. */
  regenerateSlide?(index: number, plan: { title: string; body: string; notes?: string }): void;
  /** Execute a precomputed script of layout operations on a slide. */
  executeSlideScript?(index: number, script: SlideScript): { applied: number; errors: string[] };
  /** Apply a deck-wide plan (creates/updates pages). */
  applyDeckPlan?(plan: DeckPlan): { updated: number; created: number };
}

// ============================================================================
// read_slide
// ============================================================================

const ReadSlideParams = Type.Object({
  index: Type.Integer({ minimum: 0, description: "0-based slide index" }),
});

export function createReadSlideTool(opts: { uiAdapter: ReactUIAdapter }) {
  const { uiAdapter } = opts;
  return defineTool<typeof ReadSlideParams, { found: boolean; index: number }>({
    name: "read_slide",
    label: "Read Slide",
    description: "Read a slide's full content: title, body, notes, layout name, and slot descriptors.",
    promptSnippet: "read_slide(index) — get title, body, notes, slots",
    parameters: ReadSlideParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<SlidesEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No deck available." }], details: { found: false, index: params.index } };
      }
      const slide = editor.readSlide(params.index);
      if (!slide) {
        return {
          content: [{ type: "text", text: `Slide ${params.index} not found. Deck has ${editor.getDeckSummary().pageCount} page(s).` }],
          details: { found: false, index: params.index },
        };
      }
      const slotSummary = slide.slots.map((s) => `[${s.kind}${s.text ? ` "${s.text.slice(0, 30)}…"` : ""}${s.imageRef ? " (image)" : ""}]`).join(" ");
      return {
        content: [{
          type: "text",
          text: `Slide ${slide.index}: "${slide.title}" (layout: ${slide.layoutName})\nBody: ${slide.body}\nNotes: ${slide.notes || "(none)"}\nSlots: ${slotSummary}`,
        }],
        details: { found: true, index: slide.index },
      };
    },
  });
}

// ============================================================================
// plan_deck
// ============================================================================

const DeckPlanPageSchema = Type.Object({
  title: Type.String(),
  keyPoints: Type.Array(Type.String()),
  layoutHint: Type.Optional(Type.String()),
});

const PlanDeckParams = Type.Object({
  title: Type.String({ description: "Overall deck title" }),
  audience: Type.String({ description: "Target audience description" }),
  pages: Type.Array( DeckPlanPageSchema, { description: "Ordered list of pages" }),
});

type PlanDeckArgs = Static<typeof PlanDeckParams>;

export function createPlanDeckTool(opts: { uiAdapter: ReactUIAdapter }) {
  const { uiAdapter } = opts;
  return defineTool<typeof PlanDeckParams, { pageCount: number; applied: boolean }>({
    name: "plan_deck",
    label: "Plan Deck",
    description:
      "Apply a deck-wide plan: title, audience, and ordered pages with key points. If the host supports applyDeckPlan, it will materialize the plan into the active deck. Otherwise the plan is returned for review.",
    promptSnippet: "plan_deck(title, audience, pages) — outline a full deck",
    parameters: PlanDeckParams,
    async execute(_id, params: PlanDeckArgs, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<SlidesEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No deck available." }], details: { pageCount: 0, applied: false } };
      }
      const plan: DeckPlan = { title: params.title, audience: params.audience, pages: params.pages };
      if (typeof editor.applyDeckPlan === "function") {
        const result = editor.applyDeckPlan(plan);
        return {
          content: [{ type: "text", text: `Applied deck plan: updated ${result.updated} page(s), created ${result.created} new page(s).` }],
          details: { pageCount: params.pages.length, applied: true },
        };
      }
      // Fallback: just report the plan back
      const planText = params.pages.map((p, i) => `Page ${i + 1}: "${p.title}" — ${p.keyPoints.join("; ")}`).join("\n");
      return {
        content: [{ type: "text", text: `Deck plan (not applied — host missing applyDeckPlan):\nTitle: ${params.title}\nAudience: ${params.audience}\n${planText}` }],
        details: { pageCount: params.pages.length, applied: false },
      };
    },
  });
}

// ============================================================================
// execute_slide_script
// ============================================================================

const SlideScriptOpSchema = Type.Unknown();

const ExecuteSlideScriptParams = Type.Object({
  index: Type.Integer({ minimum: 0 }),
  operations: Type.Array(SlideScriptOpSchema, { description: "Array of layout operations to apply in order" }),
});

export function createExecuteSlideScriptTool(opts: { uiAdapter: ReactUIAdapter }) {
  const { uiAdapter } = opts;
  return defineTool<typeof ExecuteSlideScriptParams, { applied: number; errors: string[] }>({
    name: "execute_slide_script",
    label: "Execute Slide Script",
    description: "Apply a script of layout operations to a single slide. Operations are applied in order; any error aborts the rest.",
    promptSnippet: "execute_slide_script(index, operations) — run layout ops on one slide",
    parameters: ExecuteSlideScriptParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<SlidesEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No deck available." }], details: { applied: 0, errors: ["no editor"] } };
      }
      if (typeof editor.executeSlideScript !== "function") {
        return {
          content: [{ type: "text", text: "Host does not implement executeSlideScript." }],
          details: { applied: 0, errors: ["not_supported"] },
        };
      }
      const result = editor.executeSlideScript(params.index, { operations: params.operations });
      const errorSuffix = result.errors.length > 0 ? `\nErrors: ${result.errors.join("; ")}` : "";
      return {
        content: [{ type: "text", text: `Executed ${result.applied} op(s) on slide ${params.index}.${errorSuffix}` }],
        details: { applied: result.applied, errors: result.errors },
      };
    },
  });
}

// ============================================================================
// regenerate_slide
// ============================================================================

const RegenerateSlideParams = Type.Object({
  index: Type.Integer({ minimum: 0 }),
  title: Type.String(),
  body: Type.String(),
  notes: Type.Optional(Type.String()),
});

export function createRegenerateSlideTool(opts: { uiAdapter: ReactUIAdapter }) {
  const { uiAdapter } = opts;
  return defineTool<typeof RegenerateSlideParams, { regenerated: boolean; index: number }>({
    name: "regenerate_slide",
    label: "Regenerate Slide",
    description: "Replace a slide's content (title, body, notes) with new generated content. Layout is preserved.",
    promptSnippet: "regenerate_slide(index, title, body, notes?) — rewrite one slide",
    parameters: RegenerateSlideParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<SlidesEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No deck available." }], details: { regenerated: false, index: params.index } };
      }
      if (typeof editor.regenerateSlide !== "function") {
        return {
          content: [{ type: "text", text: "Host does not implement regenerateSlide." }],
          details: { regenerated: false, index: params.index },
        };
      }
      editor.regenerateSlide(params.index, {
        title: params.title,
        body: params.body,
        ...(params.notes !== undefined ? { notes: params.notes } : {}),
      });
      return {
        content: [{ type: "text", text: `Regenerated slide ${params.index} with title "${params.title}".` }],
        details: { regenerated: true, index: params.index },
      };
    },
  });
}

// ============================================================================
// Extension factory
// ============================================================================

export type SlidesToolName = "read_slide" | "plan_deck" | "execute_slide_script" | "regenerate_slide";

export const ALL_SLIDES_TOOL_NAMES: readonly SlidesToolName[] = [
  "read_slide",
  "plan_deck",
  "execute_slide_script",
  "regenerate_slide",
];

export interface SlidesSkillOptions {
  uiAdapter: ReactUIAdapter;
  enabledTools?: ReadonlyArray<SlidesToolName>;
}

export function createSlidesSkillExtension(opts: SlidesSkillOptions) {
  const { uiAdapter, enabledTools } = opts;
  const enabled = new Set(enabledTools ?? ALL_SLIDES_TOOL_NAMES);

  return (pi: ExtensionAPI) => {
    if (enabled.has("read_slide")) pi.registerTool(createReadSlideTool({ uiAdapter }));
    if (enabled.has("plan_deck")) pi.registerTool(createPlanDeckTool({ uiAdapter }));
    if (enabled.has("execute_slide_script")) pi.registerTool(createExecuteSlideScriptTool({ uiAdapter }));
    if (enabled.has("regenerate_slide")) pi.registerTool(createRegenerateSlideTool({ uiAdapter }));

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (pi as any).on("before_agent_start", async () => ({
      systemPromptAppend:
        "\n## Slide Editing Rules\n- Slide indices are 0-based.\n- Use read_slide first to understand the current content and layout.\n- plan_deck replaces/creates the whole deck outline; use it for new decks, regenerate_slide for individual edits.",
    }));
  };
}

export default createSlidesSkillExtension;
