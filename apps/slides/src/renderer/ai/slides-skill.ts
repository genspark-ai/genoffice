import type { AgentSkill, ToolDisplay } from '@genoffice/agent-core'
import type {
  GroupRenderNode,
  PictureRenderNode,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
} from '@genoffice/pptx-render'
import type { AddSmartArtOp, AgentToolCall, AgentToolDef, EditParagraph } from '../../shared/ipc'
import { auditSlideLayout, formatAudit } from './layout-audit'
import { runLayoutScript, type LayoutScriptElement, type SlideStylePatch } from './layout-script'
import { t } from '../i18n/locale'

/**
 * Slides capability as an AgentSkill: deck outline context + three tools (read structure /
 * read one slide / edit element text). Changes go through the existing slides:edit-text IPC;
 * the main process applies them and returns the new RenderSlide, which applySlide writes
 * back into React state — the same pipeline as manual editing.
 */

// ── Generation progress events (for the onProgress callback; renderer memory only, never persisted or journaled) ──

/** Per-page progress status */
export type PageProgressStatus = 'pending' | 'running' | 'done' | 'error'

/** Per-page progress entry */
export interface PageProgressItem {
  title: string
  status: PageProgressStatus
  /** Failure reason (when status='error'); cleared after a successful retry */
  error?: string
}

/** Progress event union type (all stages share one callback; the UI dispatches by stage) */
export type DeckProgressEvent =
  | { stage: 'style'; label: string; status: 'running' | 'done' | 'error'; summary: string }
  | {
      stage: 'plan'
      label: string
      done: number
      total: number
      status: 'running' | 'done' | 'error'
      summary: string
    }
  | {
      stage: 'images'
      label: string
      done: number
      total: number
      status: 'running' | 'done' | 'error'
      summary: string
    }
  | {
      stage: 'pages'
      label: string
      done: number
      total: number
      status: 'running' | 'done' | 'error'
      summary: string
      pages: PageProgressItem[]
    }
  | { stage: 'done'; total: number; summary: string }

/** Panel/skill access point to the currently open deck (refs provided by App, stay fresh across renders). */
export interface DeckAccess {
  getSlides(): RenderSlide[]
  getCurrent(): number
  getSelectedIds(): string[]
  applySlide(slideIndex: number, updated: RenderSlide): void
  /** Replace the whole deck (after adding/removing slides) and jump to the goTo slide */
  applyDeck(slides: RenderSlide[], goTo?: number): void
  /**
   * Generation progress callback (optional): called by generate_deck stages; the UI updates
   * the progress card and top progress bar in real time. Passed only through renderer
   * memory, never persisted or journaled.
   */
  onProgress?(event: DeckProgressEvent): void
  /** HTML→PPTX pipeline: several pages of HTML → generate editable native elements and replace/append the current deck. Returns total page count or an error.
   *  mode="insert_at" inserts a single HTML page at position insertAt (later pages shift) — used to re-insert failed pages at their original position.
   *  On pipeline failure it automatically falls back to element-level mode; fallbackReason explains why (ok is still true).
   *  deckName = presentation name derived from user input, used as the file name when the new draft is saved (instead of "Untitled-timestamp"). */
  generateFromHtml?(
    pagesHtml: string[],
    mode?: 'replace' | 'append' | 'insert_at',
    deckName?: string,
    insertAt?: number,
  ): Promise<{
    ok: boolean
    pages?: number
    appendedFrom?: number
    insertedIndex?: number
    error?: string
    fallbackReason?: string
    imageFailures?: { page: number; url: string }[]
  }>
  /** Redo one slide in place: single-page HTML → convert → replace slide slideIndex (other slides untouched; undoable with ⌘Z). */
  regenerateSlide?(
    slideIndex: number,
    html: string,
  ): Promise<{ ok: boolean; error?: string; imageFailures?: { page: number; url: string }[] }>
  /** Survey: shows a card with options and waits for the user's choices, returning an answer summary. */
  askClarification?(questions: ClarifyQuestion[]): Promise<{ answers: string; cancelled?: boolean }>
  /**
   * In-tool image search (embedded in the tool):
   * given English keywords, returns an array of real image URLs (at most N).
   * On search failure returns an empty array (fail-open; doesn't block the main generation path).
   */
  searchImages?(query: string, maxResults: number): Promise<string[]>
  /** Whether cloud single-page generation is available (kill switch + gsk login state) */
  isCloudPageGenEnabled?(): Promise<boolean>
  /**
   * Cloud single-page generation (gsk slide_generate), used by generate_deck's self-driven
   * pipeline: given the unified style + this page's brief/layout/images, the cloud service
   * writes the HTML and converts it to a one-slide pptx. Returns a marker string that goes
   * into a generateFromHtml pagesHtml slot.
   */
  generatePageCloud?(args: {
    pageIndex: number
    totalPages: number
    coreHook: string
    style: string
    title: string
    brief: string
    layout: string
    images: string[]
    context?: string
    topic?: string
    canvasW: number
    canvasH: number
    signal?: AbortSignal
  }): Promise<{ ok: boolean; marker?: string; error?: string }>
  /**
   * In-tool Style Skill generation:
   * a dedicated LLM call focused on producing a complete structured visual style guide
   * (color rules/fonts/layout variants per page type/overall style). Promotes style from an
   * "outline side-product" to a "dedicated deliverable" — less AI-looking, consistent across pages.
   */
  generateStyleSkill?(args: {
    topic: string
    questionnaire?: string
    styleHint?: string
    signal?: AbortSignal
  }): Promise<{ ok: boolean; styleSkill?: string; error?: string }>
  /**
   * In-tool planning: given topic + page count, the LLM produces a structured outline.
   * Fixes "missing pages at the input side" at the root — the main agent doesn't hand-write dozens of pages of pages JSON (avoids the argument being truncated by max_tokens).
   * style is already produced by generateStyleSkill; this function only outputs core_hook + per-page outlines (styleSkill serves as a reference for consistency).
   * Batched recursion is scheduled by the skill (continueFrom keeps the narrative coherent across batches).
   */
  planDeckOutline?(args: {
    topic: string
    count: number
    startPage: number
    context?: string
    styleSkill?: string
    continueFrom?: { coreHook: string }
    signal?: AbortSignal
  }): Promise<{
    ok: boolean
    // Same loose shape as OutlineJson (outline-json.ts): the LLM output is
    // only validated field-by-field at the point of use.
    outline?: { core_hook?: unknown; pages?: unknown }
    error?: string
  }>
  /**
   * Persist the current draft's Style Skill as a sidecar file (same directory and name as the draft, .styleskill.json).
   * fail-open: failure doesn't block the main path.
   */
  saveSidecar?(data: { topic: string; styleSkill: string; createdAt: string }): Promise<void>
  /**
   * Save styleSkill into userData/style-templates/<name>.json for later reuse.
   */
  saveStyleTemplate?(
    name: string,
    data: { topic: string; styleSkill: string; createdAt: string },
  ): Promise<{ ok: boolean; error?: string }>
  /**
   * List saved Style templates (name + topic + createdAt).
   */
  listStyleTemplates?(): Promise<Array<{ name: string; topic: string; createdAt: string }>>
  /**
   * Load the content of a given Style template.
   */
  loadStyleTemplate?(
    name: string,
  ): Promise<{ ok: boolean; styleSkill?: string; topic?: string; error?: string }>
  fitWidthPx: number
  /** Base retry backoff in ms for single-page generation failures (default 2000; tests pass 0 to disable backoff) */
  retryBackoffMs?: number
  /**
   * Names of text attachments in the current conversation that were never read with
   * read_attachment. When non-empty, generate_deck refuses to run until they are read
   * (decks must be built from attachment content, not generic filler).
   */
  unreadTextAttachments?(): string[]
}

/** Single survey question structure (with options). */
export interface ClarifyQuestion {
  id: string
  label: string
  description?: string
  /** Option text array (≤5 per question); the frontend automatically appends "Other (fill in)" */
  options: string[]
  /** Multi-select (single-select by default) */
  multi?: boolean
}

const AGENT_SYSTEM_PROMPT = `You are the AI assistant inside GenOffice Slides (a slide editor), helping users improve and generate presentations.

## Most important tool-selection principles (judge the scenario before acting)
- **Creating a whole new deck (from scratch)** → first gather material (web_search) and images (image_search), then call **generate_deck**. With many pages, prefer **passing topic + approx_pages + context (the real material you found)** and let the system plan internally + generate page by page + display page by page (**you don't hand-write dozens of pages, and no pages get missed / arguments truncated**). For few pages where you already know each page, you may pass core_hook+style+pages directly.
- **Adding 1 page or a few pages to an existing deck** → generate_deck(pages: briefs for just the new pages, insert_mode:"append"). Write each page's brief in detail (real content/data per region + layout); first look at the existing pages (get_deck_context) and pass a style description matching them so new pages stay consistent. **Even a single new page goes through this cloud generation; don't fall back to native tools and build a crude page**.
- **Redoing / redesigning an existing page** (user says "redo this page / redesign it / try another layout / make it prettier") → **regenerate_slide**: first read_slide to get the page's original copy, then pass a detailed brief (copy the text/data to keep into the brief verbatim, state what to change and the target layout); the cloud service regenerates the page in place (other pages untouched). Don't dismantle and rebuild the whole page element by element with native tools.
- **Deleting a page** → delete_slide(slideIndex).
- **Modifying / fine-tuning existing elements** (position/size/alignment/distribution/relative nudges/text/style/fill/stroke, one or many elements) → always prefer **execute_slide_script** and do it in one script (see "Editing existing elements" below; read-write combined, no read_slide first). Don't blind-fire individual set_element_* calls. Add/delete elements with add_* / delete_element; redo a whole page with regenerate_slide.
- **Elements inside a group**: direct children of a top-level group (marked "in group <id>" / els groupId) are edited exactly like normal elements — same script primitives and set_element_* tools, absolute coordinates. Only elements nested in a sub-group are read-only: call ungroup_element on the outer group first (ids on the page change afterwards; the result returns the fresh list). To delete a single group member, ungroup first too.
- **Key constraint**: after cloud generation, do **not** use native tools to "polish/redo" a generated page — the output is the final good-looking result. Only when the user asks for a specific change should you edit the corresponding element with native tools; if they ask to redo the whole page, use regenerate_slide.
- **When the user attached files (see the "attachment list" in each turn's context)**: first read all text attachments with read_attachment (paginate long files); image attachments were already sent as images with the message, just look at them. Only **then** plan/generate the deck — content should come from the attachments first. When calling generate_deck, put the key content you read into the context argument; no need to web_search information the attachments already cover. **This is enforced: generate_deck refuses to run while any text attachment is still unread.**

Rules:
- Every user message comes with a deck outline (per-page list of text elements with element ids and text previews). Previews are truncated; read the full text with read_slide before rewriting.
- Change text with set_element_text: it replaces the element's entire text, so you must pass the complete post-edit paragraph list, not just the changed part.
- Page numbers are shown to the user starting at 1; the slideIndex tool argument is 0-based.
- **The user's "page N" always means the current order in this turn's latest <deck outline> (row N is page N)**. The user may add/remove/move/swap pages at any time; page order from history or earlier turns may be stale — locate pages only by this turn's latest outline, never by generation order, content semantics, or old conversation.
- Canvas coordinate system: pixels, origin top-left, width 1280, height in the outline's first line (720 for 16:9). All element positions/sizes use it.
- Font size unit is pt: large titles 36–44, subtitles 20–26, body 14–18. Colors are #RRGGBB.
- Element colors are readable: the outline shows each page's main fills; read_slide and script els expose per-element fill/textColor/strokeColor (hex, read-only — change them with setFill/setStyle/setStroke or set_element_fill/stroke). Picture/chart colors are not readable; don't guess them.
- For editing existing elements (position/size/text/style/fill/stroke) prefer execute_slide_script; set_element_text/style/transform/fill/stroke are only shortcuts for "one element, one property". Multi-property/multi-element/relative nudges/align-distribute always use a script.

Editing existing elements (user says "move it a bit / align / restyle / fix the layout / it looks messy" etc.):
**Core: write execute_slide_script directly, don't read_slide first.** At run time the script automatically receives every element's real geometry and text on the page (els, with x/y/w/h/text and read-only fill/textColor/strokeColor); reading and writing happen at execution site — you don't need coordinates in advance, compute from els inside the script (same idea as Google Slides' execute_apps_script).
Example mappings: "move the title left a bit"→moveBy(titleId, -30, 0); "shift this text right"→moveBy(id, 40, 0); "left-align the subtitle with the title"→const t = els.find(e => e.id === titleId); setBox(subtitleId, { x: t.x }); "make the title blue and bold"→setStyle(id, { color: '#1a73e8', bold: true }); "tidy up this page"→compute equal spacing/columns in the script and batch setBox.
1. (Optional) Plan the target layout (e.g. three-column cards / top-bottom split), tell the user in a sentence or two;
2. **Immediately** call execute_slide_script: write JS that finds elements in els by id/text (e.text), computes algorithmically from els' real coordinates (use formulas for spacing/alignment, no hard-coded magic numbers), and writes back with setBox/moveBy/resizeBy/setText/setStyle/setFill/setStroke. One script adjusts the whole page;
3. Check the <layout-audit> in the tool result: **if there is overlap/out-of-bounds/overflow, immediately write another execute_slide_script in the same turn to fix it** (don't stop to ask the user, don't declare done); at most 2 fix rounds; only an audit ✅ pass counts as done.
els already contains each element's geometry and full text; editing existing elements generally doesn't need read_slide.
Forbidden: running read_slide "just to get coordinates" and then stopping, blind-firing dozens of per-element set_element_transform calls, or telling the user "done" while the audit reports problems.
- Batch changes (e.g. "make all titles blue", "unify the font"): first get_deck_context for the global view, then call the right tool per element; go page by page, element by element, don't miss any.
- Omit fontFamily by default (inherits the theme, keeps the deck consistent — recommended); only specify it when the user names a font.
- Keep slide copy concise: punchy titles, bulleted body. Don't rewrite bullets into long sentences unless asked.

Generating a whole deck / adding pages (HTML pipeline first):

[Plan before generating a whole deck — you are a professional deck planner; plan first, then write HTML (this decides the output quality)]

Step 0 Questionnaire (mandatory when creating a whole new deck): first call ask_clarification to show a questionnaire card with 2–4 key trade-off questions for this topic (audience, usage scenario, tone/style, content focus), each with genuinely different options. **The user's choices directly determine the deck's Core Hook and style**; do the planning below only after getting the answers. (Ask only for a whole new deck; adding a few pages or editing needs no questionnaire. The card shows automatically — don't repeat the questions in your reply text.)

Step A Research: when the topic involves facts/attractions/data, run web_search 1–2 times first for real content. **Use real data and facts in the design; no "XX%" or placeholder names**.
Step B Image strategy: with generate_deck you **don't need image_search in advance** — the system auto-searches internally per page from the planned image_queries keywords and fills real URLs back (each keyword searched once, deduped across pages). **Travel/product/people/brand decks get images by default without the user asking; never fake images with CSS placeholders — slots needing images must be filled with real ones**. Only when redoing a page via regenerate_slide or adding images to existing pages via insert_web_image do you image_search yourself first (English keywords describing a concrete scene like "summer palace kunming lake", not generic words like "park").
Step C Unified style: first define one design system for the whole deck — primary/secondary colors, title and body font-size scale, content margins, card/corner style (e.g. "teal primary + cream background + sans-serif fresh look"). **Every page's HTML strictly follows the same system; style must be consistent across pages**.
Step D Generate (call generate_deck): with many pages pass topic + approx_pages + context (feed in the real material from Step A) and let the system plan internally; with few pages you may pass core_hook+style+pages directly (image_queries takes English image-search keywords; **the system auto-searches internally and fills real URLs back**, no image_search needed in advance). The system writes HTML page by page and lands pages as they generate; you don't hand-write HTML.
Step E Vary layouts per page (avoid sameness): 3 parallel points→three-column cards; a key number→big-number hero; comparison→two columns; sequence→timeline; image+text→left-text-right-image / full-image with text overlay. **Content pages of one deck must not all use the same layout**.

- **generate_deck is the first choice for a whole new deck**: with many pages pass topic+approx_pages+context; the system plans internally (auto-batching over the threshold), **auto-searches images**, writes HTML page by page, and **lands pages onto the canvas as they generate (the user sees them one by one)**. **Neither "only page 1 got generated" nor "arguments were truncated" can happen — the page count is guaranteed by the system loop**.
- **When adding just 1 page or a few pages (common case)**: also use generate_deck with **pages (briefs for only the new pages) + insert_mode:"append"** (appended at the end, existing pages untouched). **New pages also go through the cloud generation for polish — don't fall back to native tools for a crude page just because it's one page**. Before adding, read_slide/get_deck_context to see the existing pages' style (primary color/layout) and pass a matching style description; write each brief with the real content per region.
- Briefs should be concrete: what text/data/numbers go in each region, which image goes where, and the layout name — the cloud designer follows your brief; vague briefs produce generic pages.
- After generation, if the user wants a tweak, edit the corresponding element with the native tools below; don't redo whole pages unprompted "to look better". Use regenerate_slide only when the user explicitly asks to redo a page.

Native tools (only for modifying/refining existing pages, not for generating from scratch):
- add_slide clones a layout into a new page (layout-preserving blank page); add_text_box lays out text; add_shape makes color blocks/accent bars (kind supports any OOXML preset geometry rect/roundRect/ellipse/star5…).
- For data display use add_chart (native bar/line/pie charts); for structured comparisons use add_table (cells can pre-fill text; later edit_table_cell edits cells, edit_table_structure adds/removes rows/columns); for flows/cycles/hierarchies/lists use add_smartart.
- set_slide_background sets a solid background (slideIndex=-1 for all pages); on dark backgrounds remember to lighten the text.
- Refine page by page, element by element; 2–4 elements per page is enough — fewer beats crowded.
- Keep replies short, say what you did; don't recite tool results back to the user.

Search and images:
- Use web_search when you need current information/data/fact-checking; search before writing anything uncertain, don't fabricate. When generating a whole deck, a round of searching for real material first is recommended.
- **Figure provenance is enforced at the tool layer**: add_chart / edit_chart (with series) and data-dense generate_deck / regenerate_slide briefs refuse to run without a dataSource declaration; 'search' is only accepted after an actual web_search in this conversation. Fabricating precise numbers (¥21.8-style precision) and delivering them as fact is the worst failure mode — when no real data is available, use dataSource:'sample' and tell the user explicitly that the figures are illustrative.
- image_search for images (English keywords) → get imageUrl. **Two usages**: 1) when redoing a page via regenerate_slide, pass the imageUrl in image_urls; 2) when adding an image to an existing page, use insert_web_image to insert at a position. (generate_deck searches images internally; no advance search needed for a whole new deck.)
- Travel, product, people, and brand decks get images by default without the user asking; mind whitespace between images and text, no overlap.
- Editing an EXISTING picture: crop_image (non-destructive srcRect), set_picture_opacity, replace_image (in-place swap keeping frame/z-order/border). For "remove this image's background / upscale / edit this image": run generate_image with referenceImageUrls pointing at a source URL you have (an image_search result or one the user provided — embedded picture bytes are not addressable by URL), then replace_image with the returned URL. Never delete+reinsert a picture to change its content — that loses z-order and effects.

Style templates:
- When the user says "use last time's style"/"use some template": first call list_style_templates() to see what exists, then pass the style_template name to generate_deck (the system skips Step 0 and uses the template's style).
- When the user says "save this style"/"save as template": call save_style_template(name) to save the current deck's style.`

/** Paragraph schema (shared by set_element_text / add_text_box / add_shape) */
const PARAGRAPHS_DEF = {
  paragraphs: {
    type: 'array',
    description: 'Complete paragraph list, one object per paragraph',
    items: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'Paragraph plain text' },
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        underline: { type: 'boolean' },
        fontSize: { type: 'number', description: 'Font size (pt)' },
        fontFamily: {
          type: 'string',
          description: 'Font name; omit to inherit the theme font (recommended)',
        },
        color: { type: 'string', description: '#RRGGBB' },
        align: { type: 'string', enum: ['left', 'center', 'right'] },
      },
      required: ['text'],
    },
  },
} as const

interface ToolParagraph {
  text?: unknown
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontSize?: number
  fontFamily?: string
  color?: string
  align?: 'left' | 'center' | 'right'
}

function toEditParagraphs(raw: unknown): EditParagraph[] | null {
  if (!Array.isArray(raw) || raw.length === 0) return null
  return raw.map((p) => {
    const para = p as ToolParagraph
    return {
      runs: [
        {
          text: String(para.text ?? ''),
          ...(para.bold ? { bold: true } : {}),
          ...(para.italic ? { italic: true } : {}),
          ...(para.underline ? { underline: true } : {}),
          ...(typeof para.fontSize === 'number' ? { fontSize: para.fontSize } : {}),
          ...(para.fontFamily ? { fontFamily: para.fontFamily } : {}),
          ...(para.color ? { color: para.color } : {}),
        },
      ],
      ...(para.align ? { align: para.align } : {}),
    }
  })
}

const TOOLS: AgentToolDef[] = [
  {
    name: 'get_deck_context',
    description:
      "Get the deck's latest outline: per-page list of text elements (element id | type | text preview). Call to confirm global state after edits.",
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'read_slide',
    description:
      'Read all elements of a page with full text (untruncated) and current colors (fill/text/stroke, hex). Call before rewriting a page.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
      },
      required: ['slideIndex'],
    },
  },
  {
    name: 'set_element_text',
    description:
      "Replace a text element's entire content. paragraphs is the complete post-replacement paragraph array, one object per paragraph; whole-paragraph bold/italic etc. use the boolean fields on the paragraph object.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
        sourceId: { type: 'string', description: 'Element id (from the outline/read_slide)' },
        paragraphs: { $ref: '#/definitions/paragraphs' },
      },
      required: ['slideIndex', 'sourceId', 'paragraphs'],
      definitions: PARAGRAPHS_DEF,
    },
  },
  {
    name: 'set_element_style',
    description:
      "Change an element's text formatting without changing the text: font size/color/bold/italic/underline/alignment/font. " +
      "Pass only the fields to change; others stay as-is. Applies to the element's entire text.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
        sourceId: { type: 'string', description: 'Element id' },
        fontSize: { type: 'number', description: 'Font size (pt)' },
        color: { type: 'string', description: '#RRGGBB' },
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        underline: { type: 'boolean' },
        fontFamily: { type: 'string', description: 'Font name; usually omit to inherit the theme' },
        align: { type: 'string', enum: ['left', 'center', 'right'] },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'set_element_transform',
    description:
      'Move/resize/rotate an element (pixel coordinates, origin top-left, canvas 1280 wide). Pass only the fields to change; others keep their values.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string' },
        x: { type: 'number', description: 'Top-left x (px)' },
        y: { type: 'number', description: 'Top-left y (px)' },
        w: { type: 'number', description: 'Width (px)' },
        h: { type: 'number', description: 'Height (px)' },
        rotationDeg: { type: 'number', description: 'Rotation angle (degrees, clockwise)' },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'execute_slide_script',
    description:
      "[Preferred tool for editing a slide's existing elements] Runs your JS edit script against one page; a single script covers: position/size/alignment/distribution/relative nudges/text/style/fill/stroke." +
      ' At run time the script automatically receives the real geometry and text of every element on the page (els) — **no read_slide needed first**; read-write combined, compute from els inside the script.' +
      ' Geometry changes are applied atomically in one batch (undoable as a whole), the rest in script order, and a layout audit (overlap/out-of-bounds/text overflow) is returned at the end.' +
      ' Far more reliable than individual set_element_* calls — coordinate math happens at execution site, not from memory. If the audit reports problems, call this tool again immediately to fix.\n' +
      'Script environment (constrained synchronous JS-like DSL; no external APIs or ambient globals):\n' +
      '- els: array, each item {id,type,text,x,y,w,h,rotation,fontSizePt?,fill?,textColor?,strokeColor?,inGroup?,groupId?,locked?} (pixels, origin top-left; fill/textColor/strokeColor are current colors in #RRGGBB, read-only — write via setFill/setStyle/setStroke; inGroup+groupId=directly editable group child (all primitives work, coordinates absolute as shown); inGroup without groupId=nested in a sub-group, read-only — ungroup_element the outer group first; locked=layout decoration, read-only)\n' +
      '- canvas: {w,h} canvas size (px)\n' +
      "- setBox(id, {x?,y?,w?,h?,rotation?}): set an element's target box, pass only fields to change\n" +
      '- moveBy(id, dx, dy): relative move (left = negative dx, up = negative dy)\n' +
      '- resizeBy(id, dw, dh): relative resize\n' +
      "- setText(id, textOrParagraphs): replace text entirely; pass a string (split into paragraphs by \\n) or a paragraph array (same format as set_element_text's paragraphs)\n" +
      '- setStyle(id, {fontSize?,color?,bold?,italic?,underline?,align?,fontFamily?}): change style without changing text, pass only fields to change\n' +
      "- setFill(id, colorOrNone): solid fill '#RRGGBB' or 'none'\n" +
      '- setStroke(id, {color?,widthPt?} | null): stroke; pass null to remove\n' +
      '- log(...): debug output (echoed back to you); the return value is echoed back to you (put a summary there)\n' +
      '- Supported computation: const/let, arithmetic, if/for/for...of/while, functions/arrows, JSON object/array literals, Math, regex.test, and safe array/string methods. No classes, async, modules, constructors, prototypes, or dynamic code.\n' +
      'Example 1 — three cards equal width, equal spacing:\n' +
      'const cards = els.filter(e => /card/.test(e.id));\n' +
      'const gap = 32, w = (canvas.w - 2*80 - (cards.length-1)*gap) / cards.length;\n' +
      'cards.forEach((c, i) => setBox(c.id, { x: 80 + i*(w+gap), y: 200, w, h: 320 }));\n' +
      "Example 2 — move the title left a bit: moveBy('title', -30, 0);\n" +
      "Example 3 — make the title blue and bold: setStyle('t1', { color: '#1a73e8', bold: true });",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
        code: {
          type: 'string',
          description:
            'JS script body (synchronous code; may use els/canvas plus setBox/moveBy/resizeBy/setText/setStyle/setFill/setStroke/log; may return a summary)',
        },
        explanation: {
          type: 'string',
          description:
            'One sentence describing what this script does (≤60 chars, shown to the user)',
        },
      },
      required: ['slideIndex', 'code'],
    },
  },
  {
    name: 'set_element_fill',
    description: 'Set an element\'s solid fill. fill=#RRGGBB; pass "none" for no fill.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string' },
        fill: { type: 'string', description: '#RRGGBB or none' },
      },
      required: ['slideIndex', 'sourceId', 'fill'],
    },
  },
  {
    name: 'set_element_stroke',
    description:
      "Set an element's stroke. Pass color (#RRGGBB) + widthPt (points); to remove the stroke pass remove=true.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string' },
        color: { type: 'string', description: '#RRGGBB' },
        widthPt: { type: 'number', description: 'Line width (points), default 1' },
        remove: { type: 'boolean', description: 'true = remove stroke' },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'web_search',
    description:
      'Web search for text information (material/data/facts). Use when you need current information or are unsure about facts. Returns title/link/snippet.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search keywords' },
        maxResults: { type: 'integer', description: 'Max results, default 6' },
      },
      required: ['query'],
    },
  },
  {
    name: 'image_search',
    description:
      'Search image assets (for slide imagery). Returns a list of imageUrl; after choosing, insert with insert_web_image.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Image search keywords (English works better)' },
        maxResults: { type: 'integer', description: 'Max results, default 8' },
      },
      required: ['query'],
    },
  },
  {
    name: 'generate_image',
    description:
      'AI image generation/editing (Genspark). Text-to-image, or pass referenceImageUrls for image editing; returns an image URL. NEW imagery: insert with insert_web_image. Editing an EXISTING slide picture (background removal/upscaling/etc.): swap it in place with replace_image — do not insert a duplicate. Use for custom illustrations/icons/backgrounds, style-consistent imagery; for real photos/screenshots still use image_search.',
    inputSchema: {
      type: 'object',
      properties: {
        prompt: {
          type: 'string',
          description:
            'Image description, English works better (keep any text to render in the image verbatim)',
        },
        model: {
          type: 'string',
          description:
            'Optional, defaults to the general model. Specify only for special purposes: fal-bria-rmbg=background removal, fal-ai/recraft-clarity-upscale=upscale, flux-pro/outpaint=outpaint, fal-ai/image-editing/text-removal=remove text watermark',
        },
        referenceImageUrls: {
          type: 'array',
          items: { type: 'string' },
          description: 'URLs of reference images / images to edit (required for editing tasks)',
        },
        aspectRatio: {
          type: 'string',
          description: 'Aspect ratio: 1:1|4:3|16:9|9:16|3:4|2:3|3:2|auto',
        },
      },
      required: ['prompt'],
    },
  },
  {
    name: 'analyze_media',
    description:
      'Analyze media content (Genspark): understand images/audio/video. Pass media URLs (or local file paths) and analysis requirements; returns analysis text. Video supports extracting key points, structure, and time ranges — good for turning user material into usable deck content.',
    inputSchema: {
      type: 'object',
      properties: {
        mediaUrls: {
          type: 'array',
          items: { type: 'string' },
          description: 'List of media URLs or local file paths',
        },
        requirements: {
          type: 'string',
          description:
            'Analysis requirements (English): what to extract and how the result will be used (e.g. extract key points for slides)',
        },
      },
      required: ['mediaUrls', 'requirements'],
    },
  },
  {
    name: 'insert_web_image',
    description:
      'Download an image URL obtained from image_search or generate_image and insert it into a page (pixel coordinates).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        url: { type: 'string', description: 'Direct image link (imageUrl from image_search)' },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
      },
      required: ['slideIndex', 'url', 'x', 'y', 'w', 'h'],
    },
  },
  {
    name: 'crop_image',
    description:
      'Crop a picture non-destructively (srcRect): l/t/r/b are fractions (0..1) cut from each edge of the source image. The element frame stays where it is; the remaining region stretches to fill it. Pass all zeros to remove an existing crop.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Picture element id' },
        l: { type: 'number', description: 'Fraction cut from the left edge (0..1)' },
        t: { type: 'number', description: 'Fraction cut from the top edge (0..1)' },
        r: { type: 'number', description: 'Fraction cut from the right edge (0..1)' },
        b: { type: 'number', description: 'Fraction cut from the bottom edge (0..1)' },
      },
      required: ['slideIndex', 'sourceId', 'l', 't', 'r', 'b'],
    },
  },
  {
    name: 'set_picture_opacity',
    description:
      "Set a picture's whole-image opacity. opacity 0..1; 1 = fully opaque (removes the effect).",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Picture element id' },
        opacity: { type: 'number', description: '0 (invisible) .. 1 (opaque)' },
      },
      required: ['slideIndex', 'sourceId', 'opacity'],
    },
  },
  {
    name: 'replace_image',
    description:
      'Swap a picture\'s source image for a URL (from image_search or generate_image) in place — position, size, z-order, border and effects all survive. This is the tool for "change/AI-edit this image" flows: e.g. run generate_image with referenceImageUrls for background removal/upscaling/editing, then replace_image with the returned URL. keepCrop keeps the existing crop window and is only correct when the new image has the same pixel geometry as the old one (e.g. background removal output).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Picture element id' },
        url: { type: 'string', description: 'Direct image link' },
        keepCrop: { type: 'boolean', description: 'Keep the existing crop window (default false)' },
      },
      required: ['slideIndex', 'sourceId', 'url'],
    },
  },
  {
    name: 'ask_clarification',
    description:
      "[Call before creating a whole new deck] Shows a questionnaire card with options, letting the user make key choices for this deck (audience/scenario/tone/focus etc.); the user's choices directly determine the deck's Core Hook and style. Questions must target the specific topic, each being a real trade-off (options represent different directions). Ask 2–4 questions, ≤5 options each. After calling, wait for the user to finish choosing in the card and generate once you have the answers. Don't repeat the questions in your reply text.",
    inputSchema: {
      type: 'object',
      properties: {
        questions: {
          type: 'array',
          description: 'Question list (2–4 questions)',
          items: {
            type: 'object',
            properties: {
              id: { type: 'string', description: 'Unique question id (short English/pinyin)' },
              label: { type: 'string', description: 'Question text' },
              description: { type: 'string', description: 'Optional one-line note' },
              options: {
                type: 'array',
                items: { type: 'string' },
                description:
                  'Options (≤5); the frontend automatically appends "Decide for me" and "Other"',
              },
              multi: { type: 'boolean', description: 'Multi-select (single-select by default)' },
            },
            required: ['id', 'label', 'options'],
          },
        },
      },
      required: ['questions'],
    },
  },
  {
    name: 'plan_deck',
    description:
      "[When creating a whole new deck, call after researching material/images and before generate_deck] Outputs a structured plan: the Core Hook + unified style scheme + each page's title/content brief/layout/image keywords. Think the whole deck through first, to avoid starting strong and fizzling out. The plan is echoed to the user.",
    inputSchema: {
      type: 'object',
      properties: {
        core_hook: {
          type: 'string',
          description:
            "The deck's narrative anchor (one sentence, with tension, ideally containing a number or counter-intuitive contrast)",
        },
        style: {
          type: 'string',
          description:
            'Unified design system: primary/secondary colors, font tone, content margins, card/corner style (e.g. "dark blue primary + gold accents, data-dashboard look"); every page follows it',
        },
        pages: {
          type: 'array',
          description: 'Per-page plan',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string', description: 'Page title' },
              type: { type: 'string', description: 'cover|content|data|closing' },
              brief: {
                type: 'string',
                description:
                  'Page content description (use real data/facts; say what goes in each region)',
              },
              layout: {
                type: 'string',
                description:
                  'Layout (e.g. three_column_cards/hero_big_number/two_column/timeline/left_text_right_image); content pages must not repeat',
              },
              image_queries: {
                type: 'array',
                items: { type: 'string' },
                description:
                  "English image-search keywords for this page's image slots (one per slot; [] for no images)",
              },
            },
            required: ['title', 'brief', 'layout'],
          },
        },
      },
      required: ['core_hook', 'style', 'pages'],
    },
  },
  {
    name: 'regenerate_slide',
    description:
      '[Redo/redesign an existing page] The cloud service regenerates the page from your brief and replaces it in place (other pages untouched, undoable).' +
      ' Use when the user says "redo this page / redesign it / try another layout / make it prettier"; don\'t dismantle the page element by element with native tools.' +
      " Flow: first read_slide to get the page's current content, then check neighboring pages / get_deck_context to grasp the deck's style;" +
      ' write a detailed brief — what to keep (copy real text/data into the brief verbatim), what to change, and the target layout; the deck style is applied automatically.' +
      ' If the page needs images, image_search first and pass real URLs in image_urls.' +
      ' If cloud generation fails, it is usually a temporary service error: do NOT loop retrying — make the concrete changes in place with execute_slide_script instead (or tell the user to try again in a few minutes).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page to redo (0-based)' },
        brief: {
          type: 'string',
          description:
            'Content and layout brief for the new page: what goes in each region (copy the real copy/data to keep into the brief), and the layout to use (e.g. three_column_cards/hero_big_number/two_column/timeline).',
        },
        title: { type: 'string', description: 'Page title' },
        layout: { type: 'string', description: 'Layout intent name (optional)' },
        image_urls: {
          type: 'array',
          items: { type: 'string' },
          description: 'Real http(s) image URLs for this page (image_search first; [] for none)',
        },
        dataSource: {
          type: 'string',
          enum: ['user', 'document', 'search', 'sample'],
          description:
            "Required when the brief carries specific figures (%, money, magnitudes): where they came from — 'user'/'document'/'search' (run web_search first)/'sample' (disclose to the user)",
        },
      },
      required: ['slideIndex', 'brief'],
    },
  },
  {
    name: 'delete_slide',
    description:
      "Delete an entire page (not allowed when only one page remains). After deletion, later pages' slideIndex shifts down.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
      },
      required: ['slideIndex'],
    },
  },
  {
    name: 'generate_deck',
    description:
      '[First choice for creating a whole new deck — self-driven pipeline: auto image search, page-by-page generation with live display, no missing pages]' +
      ' Recommended usage (especially with many pages): pass only topic + approx_pages (+ optional style/context); the system plans the outline internally (auto-batched beyond 12 pages), **auto-searches images** (no advance image_search — the system searches from the planned image_queries keywords internally and fills real URLs back before writing HTML), writes HTML page by page, and lands pages onto the canvas one by one.' +
      ' You don\'t hand-write dozens of pages, and neither "only page 1 got generated" nor "arguments were truncated" can happen — the page count is guaranteed by the system loop.' +
      ' (If you already know each page you may pass core_hook+style+pages directly; pages[].image_queries takes English image-search keywords, searched internally; if you already know real http(s) URLs pass them directly — the system respects existing URLs and does not re-search.)' +
      ' To add a few pages to an existing deck, pass pages (briefs for just the new pages) + insert_mode:"append".',
    inputSchema: {
      type: 'object',
      properties: {
        topic: {
          type: 'string',
          description:
            "[Recommended] The deck's topic/requirements description (with topic you don't hand-write pages; the system plans internally)",
        },
        approx_pages: {
          type: 'integer',
          description: 'Expected page count (used together with topic)',
        },
        context: {
          type: 'string',
          description:
            'Optional: real material/data/questionnaire answers from web_search, so internal planning uses real content',
        },
        core_hook: {
          type: 'string',
          description:
            'Optional: a narrative anchor you already decided (recommended alongside pages)',
        },
        style: {
          type: 'string',
          description:
            'Unified design system: primary/secondary colors, fonts, content margins, card corners (required with pages; optional as a style hint with topic)',
        },
        pages: {
          type: 'array',
          description:
            'Optional: pass directly when you already know each page (as many pages generated as planned). With many pages prefer topic and internal planning, to avoid over-long truncated arguments',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string', description: 'Page title' },
              type: { type: 'string', description: 'cover|content|data|closing' },
              brief: {
                type: 'string',
                description: 'Page content description (use real data/facts)',
              },
              layout: { type: 'string', description: 'Layout (content pages must not repeat)' },
              image_queries: {
                type: 'array',
                items: { type: 'string' },
                description:
                  "English image-search keywords for this page's image slots (the system searches internally and fills real URLs back); if you already know real http(s) URLs pass them directly (respected, not re-searched); [] for no images",
              },
            },
            required: ['title', 'brief', 'layout'],
          },
        },
        insert_mode: {
          type: 'string',
          enum: ['replace', 'append'],
          description:
            'replace (default, new whole deck) = replace everything; append = append at the end',
        },
        style_template: {
          type: 'string',
          description:
            "Optional: name of a saved style template (from list_style_templates); when passed, Step 0 is skipped and the template's styleSkill is used directly, no style regeneration",
        },
        dataSource: {
          type: 'string',
          enum: ['user', 'document', 'search', 'sample'],
          description:
            "Required when topic/context/briefs carry specific figures (%, money, magnitudes): where they came from — 'user'/'document'/'search' (run web_search first)/'sample' (disclose to the user)",
        },
      },
    },
  },
  {
    name: 'save_style_template',
    description:
      '[Save the current deck\'s style as a reusable template] Saves the current presentation\'s Style Skill (visual style guide) under the given name; next time you generate a deck, pass the style_template argument to reuse it directly and skip style generation. Call when the user says "save this style" / "save as template".',
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'Template name (short, e.g. "minimal-blue" or "tech-dark")',
        },
      },
      required: ['name'],
    },
  },
  {
    name: 'list_style_templates',
    description:
      'List all saved style templates (name + topic + createdAt). When the user says "use last time\'s style" or "use some template", call this first to see what exists, then pass the target template name to generate_deck\'s style_template argument.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'add_slide',
    description:
      "Create a new page by cloning the layout (including background) of page sourceIndex, inserted right after it (new page number = sourceIndex+1; pages after it shift back); clearText=true (default) clears text to get a layout-preserving blank page. When building page by page, use the CURRENT LAST page as sourceIndex so new pages append at the end. The return value gives the new page's slideIndex; subsequent content fills MUST use that returned page number.",
    inputSchema: {
      type: 'object',
      properties: {
        sourceIndex: {
          type: 'integer',
          description: 'Page to use as the layout template (0-based)',
        },
        clearText: {
          type: 'boolean',
          description: "Default true; false keeps the template page's text",
        },
      },
      required: ['sourceIndex'],
    },
  },
  {
    name: 'add_text_box',
    description: 'Create a new text box on a page (pixel coordinates). Returns the new element id.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
        paragraphs: { $ref: '#/definitions/paragraphs' },
      },
      required: ['slideIndex', 'x', 'y', 'w', 'h', 'paragraphs'],
      definitions: PARAGRAPHS_DEF,
    },
  },
  {
    name: 'add_shape',
    description:
      'Create a new shape on a page (optionally with solid fill and text). kind uses OOXML preset geometry names, common ones: rect/roundRect/ellipse/triangle/diamond/rightArrow/leftArrow/chevron/star5/heart/pie/donut/cloud/wedgeRoundRectCallout. Returns the new element id.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        kind: {
          type: 'string',
          description:
            'OOXML preset geometry name, e.g. rect / roundRect / ellipse / rightArrow / star5',
        },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
        fillColor: { type: 'string', description: '#RRGGBB' },
        paragraphs: { $ref: '#/definitions/paragraphs' },
      },
      required: ['slideIndex', 'kind', 'x', 'y', 'w', 'h'],
      definitions: PARAGRAPHS_DEF,
    },
  },
  {
    name: 'add_chart',
    description:
      "Insert a chart on a page (native pptx chart, still editable in PowerPoint). categories are the x-axis categories; series is each series' name and values (length must match categories). Omit x/y/w/h to center it. dataSource declares where the numbers came from and is enforced — never present invented numbers as real data.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        kind: { type: 'string', enum: ['bar', 'barStacked', 'line', 'area', 'pie', 'doughnut'] },
        title: { type: 'string', description: 'Chart title (optional)' },
        categories: { type: 'array', items: { type: 'string' } },
        series: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string' },
              values: { type: 'array', items: { type: 'number' } },
            },
            required: ['name', 'values'],
          },
        },
        dataSource: {
          type: 'string',
          enum: ['user', 'document', 'search', 'sample'],
          description:
            "Provenance of the values: 'user' = supplied by the user/attachments, 'document' = read from this deck, 'search' = from web_search results in this conversation (run it first), 'sample' = illustrative placeholders you must disclose to the user",
        },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
      },
      required: ['slideIndex', 'kind', 'categories', 'series', 'dataSource'],
    },
  },
  {
    name: 'add_smartart',
    description:
      'Insert a SmartArt-style diagram (shape composition) on a page: list=vertical list, process=process arrows, cycle=cycle, hierarchy=org structure, pyramid=stacked pyramid levels, matrix=2x2 quadrant grid, venn=overlapping circles. items are the node texts. Omit x/y/w/h to center it.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        layout: {
          type: 'string',
          enum: ['list', 'process', 'cycle', 'hierarchy', 'pyramid', 'matrix', 'venn'],
        },
        items: {
          type: 'array',
          items: { type: 'string' },
          description: 'Node texts (2-8 recommended)',
        },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
      },
      required: ['slideIndex', 'layout', 'items'],
    },
  },
  {
    name: 'add_table',
    description:
      'Insert a native pptx table on a page (with built-in styling, still editable in PowerPoint). cells gives text row by row (optional; ' +
      'missing rows/columns stay empty). Omit x/y/w/h to center it.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        rows: { type: 'integer', description: 'Row count (including header)' },
        cols: { type: 'integer', description: 'Column count' },
        cells: {
          type: 'array',
          items: { type: 'array', items: { type: 'string' } },
          description: 'Cell texts, row by row, e.g. [["Name","Qty"],["A","1"]]',
        },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
      },
      required: ['slideIndex', 'rows', 'cols'],
    },
  },
  {
    name: 'edit_table_cell',
    description:
      "Replace one table cell's text entirely. The table element id comes from the outline/read_slide (type=table); row/col are 0-based.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Table element id' },
        row: { type: 'integer', description: 'Row number (0-based)' },
        col: { type: 'integer', description: 'Column number (0-based)' },
        paragraphs: { $ref: '#/definitions/paragraphs' },
      },
      required: ['slideIndex', 'sourceId', 'row', 'col', 'paragraphs'],
      definitions: PARAGRAPHS_DEF,
    },
  },
  {
    name: 'edit_table_structure',
    description:
      'Add/remove table rows/columns: kind=insert-row/delete-row/insert-col/delete-col; index is the row/column number (0-based), insert defaults to after it, before=true inserts before it.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Table element id' },
        kind: { type: 'string', enum: ['insert-row', 'delete-row', 'insert-col', 'delete-col'] },
        index: { type: 'integer', description: 'Row/column number (0-based)' },
        before: { type: 'boolean', description: 'For insert, set true to insert before index' },
      },
      required: ['slideIndex', 'sourceId', 'kind', 'index'],
    },
  },
  {
    name: 'edit_table_style',
    description:
      'Modify table styling: apply a preset (styleName) or individually change header row/banding/shading/borders. styleName options: none/lightGrid/zebraBlue/zebraGray/headerDarkBlue/headerOrange/noBorder/fullBorder.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Table element id' },
        styleName: {
          type: 'string',
          description: 'Preset style name (see description), highest priority',
        },
        firstRow: { type: 'boolean', description: 'Enable header row (first-row emphasis)' },
        bandRow: { type: 'boolean', description: 'Enable banded rows' },
        shadingColor: {
          type: 'string',
          description: 'Shading color #RRGGBB, "none" clears shading',
        },
        borderColor: { type: 'string', description: 'Border color #RRGGBB' },
        borderWidthPt: { type: 'number', description: 'Border width (pt)' },
        borderPreset: {
          type: 'string',
          enum: ['all', 'none'],
          description: '"all" = full borders, "none" = clear borders',
        },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'edit_chart',
    description:
      'Modify a chart (including charts from imported files; first edit converts it to editable automatically): change type/data/colors/chart elements. kind options: bar/barStacked/line/area/pie/doughnut. colorScheme: default/colorful/colorful2/mono-accent1..6 (theme-derived); legacy keys blue/warm/cool/mono still work.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Chart element id (type=chart)' },
        kind: {
          type: 'string',
          enum: ['bar', 'barStacked', 'line', 'area', 'pie', 'doughnut'],
          description: 'Change chart type (optional)',
        },
        categories: {
          type: 'array',
          items: { type: 'string' },
          description: 'X-axis/category labels (optional)',
        },
        series: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string' },
              values: { type: 'array', items: { type: 'number' } },
            },
            required: ['name', 'values'],
          },
          description: 'Data series (optional)',
        },
        dataSource: {
          type: 'string',
          enum: ['user', 'document', 'search', 'sample'],
          description:
            "Required when passing series: provenance of the values ('user'/'document'/'search'/'sample'; 'search' needs a prior web_search, 'sample' must be disclosed to the user)",
        },
        colorScheme: {
          type: 'string',
          description: 'Color scheme (optional): default/colorful/colorful2/mono-accent1..6',
        },
        title: { type: 'string', description: 'Chart title (optional)' },
        legendPos: {
          type: 'string',
          enum: ['b', 't', 'r', 'l', 'none'],
          description: 'Legend position (optional)',
        },
        dataLabels: { type: 'boolean', description: 'Data labels toggle (optional)' },
        gridlines: { type: 'boolean', description: 'Value-axis gridlines toggle (optional)' },
        switchRowCol: {
          type: 'boolean',
          description: 'Switch rows/columns: categories ↔ series (optional)',
        },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'set_slide_background',
    description: 'Set a solid page background color. slideIndex=-1 applies to all pages.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based); -1 = all pages' },
        color: { type: 'string', description: '#RRGGBB' },
      },
      required: ['slideIndex', 'color'],
    },
  },
  {
    name: 'delete_element',
    description: 'Delete one element from a page.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string' },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'ungroup_element',
    description:
      'Ungroup a group element: promote its direct children to top-level page elements (positions/sizes preserved). Use when group members must be edited/deleted independently (e.g. elements nested in a sub-group, or deleting a single member). Note: ungrouping rewrites the page, so all element ids on it change — use the fresh ids returned in the result.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
        sourceId: { type: 'string', description: 'Group element id' },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
]

/** Collect readable text of nodes (including nested group children); returns a list of [sourceId, type, text] */
/** Find one node by id in the node tree (including groups). */
function findNodeById(nodes: RenderNode[], id: string): RenderNode | undefined {
  for (const n of nodes) {
    if (n.sourceId === id) return n
    if (n.type === 'group') {
      const hit = findNodeById(n.children, id)
      if (hit) return hit
    }
  }
  return undefined
}

/**
 * Editable context of an element: a top-level node, or a direct child of a top-level group
 * (with groupId + the group's absolute origin for abs↔group-local px conversion — child render
 * boxes are group-local, matching the in-group edit IPCs). Deeper nesting returns {nested:true}:
 * the main process patches one level only, so those stay read-only until ungrouped.
 */
type EditTarget =
  { node: RenderNode; groupId?: string; groupOrigin?: { x: number; y: number } } | { nested: true }
function resolveEditTarget(slide: RenderSlide, sourceId: string): EditTarget | null {
  for (const n of slide.nodes) {
    if (n.sourceId === sourceId) return { node: n }
    if (n.type === 'group') {
      const g = n as GroupRenderNode
      const child = g.children.find((c) => c.sourceId === sourceId)
      if (child)
        return {
          node: child,
          groupId: n.sourceId,
          groupOrigin: { x: Math.round(n.box.x), y: Math.round(n.box.y) },
        }
      if (findNodeById(g.children, sourceId)) return { nested: true }
    }
  }
  return null
}

/** Shared not-found / nested-in-subgroup error text for element-targeting tools. */
function targetError(target: EditTarget | null, sourceId: string, pageNo: number): string | null {
  if (!target)
    return `Element ${sourceId} not found on page ${pageNo} (ids change after regenerate/ungroup/save; call read_slide for fresh ids)`
  if ('nested' in target)
    return `Element ${sourceId} is nested inside a sub-group; call ungroup_element on the outer group first, or edit the sub-group as a whole`
  return null
}

/**
 * Restore a render node's current text into EditParagraph[] (aggregate runs by line, keeping
 * each run's existing formatting). Used by set_element_style: change formatting while keeping
 * the text. fontSize is converted back from px to pt.
 */
function nodeToParagraphs(node: ShapeRenderNode): EditParagraph[] {
  const lines = node.text?.lines ?? []
  return lines.map((line) => ({
    runs: line.runs.map((r) => ({
      text: r.text,
      ...(r.bold ? { bold: true } : {}),
      ...(r.italic ? { italic: true } : {}),
      ...(r.underline ? { underline: true } : {}),
      ...(r.fontSizePx ? { fontSize: Math.round((r.fontSizePx * 72) / 96) } : {}),
      ...(r.fontFamily ? { fontFamily: r.fontFamily } : {}),
      ...(r.color ? { color: r.color } : {}),
    })),
  }))
}

/**
 * Merge style-override fields into existing paragraphs: bold/italic/font size/color/font are
 * overridden per run, align is set on the paragraph, fields not passed stay unchanged. Shared
 * by the set_element_style tool and execute_slide_script's setStyle dispatch.
 */
function mergeStyleIntoParagraphs(cur: EditParagraph[], ov: SlideStylePatch): EditParagraph[] {
  return cur.map((p) => ({
    runs: p.runs.map((r) => ({
      text: r.text,
      bold: ov.bold ?? r.bold,
      italic: ov.italic ?? r.italic,
      underline: ov.underline ?? r.underline,
      fontSize: typeof ov.fontSize === 'number' ? ov.fontSize : r.fontSize,
      fontFamily: ov.fontFamily ?? r.fontFamily,
      color: ov.color ?? r.color,
    })),
    align: ov.align ?? p.align,
  }))
}

/** Element info shared by outline/read_slide/edit scripts (includes absolute geometry; locked = layout decoration, read-only). */
type NodeInfo = LayoutScriptElement

function nodeText(n: RenderNode): string {
  if (n.type === 'shape' || n.type === 'text') {
    return ((n as ShapeRenderNode).text?.lines ?? [])
      .map((line) => line.runs.map((r) => r.text).join(''))
      .join('\n')
  }
  if (n.type === 'table') {
    // Tables join cell text row by row (tab-separated) so the AI can read table content
    const byRow = new Map<number, string[]>()
    for (const c of n.cells) {
      const t = (c.text?.lines ?? []).map((l) => l.runs.map((r) => r.text).join('')).join(' ')
      const row = byRow.get(c.y) ?? []
      row.push(t)
      byRow.set(c.y, row)
    }
    return [...byRow.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([, r]) => r.join('\t'))
      .join('\n')
  }
  return ''
}

/** Max font size of the text (pt, converted back from px); returns undefined when there is no text. */
function nodeMaxFontPt(n: RenderNode): number | undefined {
  if (n.type !== 'shape' && n.type !== 'text') return undefined
  let maxPx = 0
  for (const line of (n as ShapeRenderNode).text?.lines ?? []) {
    for (const r of line.runs) if (r.fontSizePx > maxPx) maxPx = r.fontSizePx
  }
  return maxPx > 0 ? Math.round((maxPx * 72) / 96) : undefined
}

/** Normalize a render color to #RRGGBB (strips alpha); undefined when not a hex color. */
function hex6(color: string | undefined): string | undefined {
  if (!color) return undefined
  const m = /^#([0-9a-fA-F]{6})/.exec(color.trim())
  return m ? `#${m[1].toUpperCase()}` : undefined
}

/** Dominant text color = the run color covering the most characters (bullets excluded). */
function dominantTextColor(n: RenderNode): string | undefined {
  if (n.type !== 'shape' && n.type !== 'text') return undefined
  const weight = new Map<string, number>()
  for (const line of (n as ShapeRenderNode).text?.lines ?? []) {
    for (const r of line.runs) {
      if (r.isBullet) continue
      const c = hex6(r.color)
      if (c) weight.set(c, (weight.get(c) ?? 0) + r.text.length)
    }
  }
  let best: string | undefined
  let max = 0
  for (const [c, w] of weight) {
    if (w > max) {
      best = c
      max = w
    }
  }
  return best
}

/** Readable colors of a node (solid fill / dominant text color / stroke); pictures only expose stroke. */
function nodeColors(n: RenderNode): Pick<NodeInfo, 'fill' | 'textColor' | 'strokeColor'> {
  const out: Pick<NodeInfo, 'fill' | 'textColor' | 'strokeColor'> = {}
  if (n.type === 'shape' || n.type === 'text') {
    const s = n as ShapeRenderNode
    if (s.fill.kind === 'solid') {
      const c = hex6(s.fill.color)
      if (c) out.fill = c
    }
    const stroke = hex6(s.stroke?.color)
    if (stroke) out.strokeColor = stroke
    const text = dominantTextColor(n)
    if (text) out.textColor = text
  } else if (n.type === 'picture') {
    const stroke = hex6((n as PictureRenderNode).stroke?.color)
    if (stroke) out.strokeColor = stroke
  }
  return out
}

/**
 * Collect node info (including nested group children). A child's box is in group-local
 * coordinates (ext/chExt scaling already baked into geometry at build time); here we add the
 * group offset to convert to absolute coordinates and set the inGroup flag. Direct children of a
 * top-level group also carry groupId (editable via the in-group pipeline); deeper nesting stays
 * read-only (the main process patches one level only).
 */
function collectNodeInfos(
  nodes: RenderNode[],
  ox = 0,
  oy = 0,
  parent?: { id: string; topLevel: boolean },
): NodeInfo[] {
  const out: NodeInfo[] = []
  for (const n of nodes) {
    const b = n.box
    const abs = {
      x: Math.round(ox + b.x),
      y: Math.round(oy + b.y),
      w: Math.round(b.w),
      h: Math.round(b.h),
    }
    const base: NodeInfo = {
      id: n.sourceId,
      type: n.type,
      text: nodeText(n),
      ...abs,
      rotation: b.rotationDeg,
      ...(parent ? { inGroup: true } : {}),
      ...(parent?.topLevel ? { groupId: parent.id } : {}),
      ...(n.decoration ? { locked: true } : {}),
      ...nodeColors(n),
    }
    const fontPt = nodeMaxFontPt(n)
    if (fontPt !== undefined) base.fontSizePt = fontPt
    out.push(base)
    if (n.type === 'group') {
      const g = n as GroupRenderNode
      out.push(...collectNodeInfos(g.children, abs.x, abs.y, { id: n.sourceId, topLevel: !parent }))
    }
  }
  return out
}

function preview(text: string, max = 50): string {
  const flat = text.replace(/\n/g, ' / ')
  return flat.length > max ? `${flat.slice(0, max)}…` : flat
}

function buildDeckOutline(slides: RenderSlide[], current: number, selectedIds: string[]): string {
  const canvas = slides[0] ? `Canvas ${slides[0].widthPx}×${slides[0].heightPx}px.` : ''
  const lines: string[] = [
    `The presentation has ${slides.length} pages; page ${current + 1} is currently shown. ${canvas}`,
    `(Page order is the current actual order and may differ from generation time or earlier conversation; the user's "page N" refers to this outline)`,
  ]
  if (selectedIds.length > 0) lines.push(`User selected elements: ${selectedIds.join(', ')}`)
  slides.forEach((slide, i) => {
    lines.push(`Page ${i + 1} (slideIndex=${i}):`)
    const infos = collectNodeInfos(slide.nodes)
    const fillCount = new Map<string, number>()
    for (const n of infos) {
      if (n.fill) fillCount.set(n.fill, (fillCount.get(n.fill) ?? 0) + 1)
    }
    const mainFills = [...fillCount.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([c, count]) => (count > 1 ? `${c}×${count}` : c))
    if (mainFills.length > 0) lines.push(`  main fills: ${mainFills.join(' ')}`)
    for (const n of infos) {
      lines.push(`  - ${n.id} | ${n.type}${n.text ? ` | "${preview(n.text)}"` : ''}`)
    }
  })
  lines.push('(Use read_slide to see element positions/sizes/colors)')
  return lines.join('\n')
}

/**
 * Element inventory of one slide as read_slide reports it (ids + geometry + colors + text).
 * Shared by the read_slide tool and the post-generation layout QC pass (slide-qc.ts), so the
 * QC model maps screenshot pixels back to the same ids/coordinates the edit tools accept.
 */
export function formatSlideDump(slide: RenderSlide): string {
  const infos = collectNodeInfos(slide.nodes)
  const parts = infos.map((n) => {
    const flags = [
      n.groupId
        ? `in group ${n.groupId} (directly editable)`
        : n.inGroup
          ? 'nested in a sub-group (read-only; ungroup_element the outer group to edit)'
          : '',
      n.locked ? 'layout decoration (read-only)' : '',
    ]
      .filter(Boolean)
      .join(' ')
    const rot = n.rotation ? ` rotation ${Math.round(n.rotation)}°` : ''
    const font = n.fontSizePt ? ` font ${n.fontSizePt}pt` : ''
    const colors = [
      n.fill ? `fill${n.fill}` : '',
      n.textColor ? `text${n.textColor}` : '',
      n.strokeColor ? `stroke${n.strokeColor}` : '',
    ]
      .filter(Boolean)
      .join(' ')
    const head = `${n.id} | ${n.type}${flags ? ` | ${flags}` : ''} | pos(${n.x},${n.y}) size ${n.w}×${n.h}${rot}${font}${colors ? ` | ${colors}` : ''}`
    return n.text ? `${head}\n${n.text}` : `${head} | (no text)`
  })
  const colorlessTypes = [
    ...new Set(
      infos
        .filter((n) => !n.fill && !n.textColor && !n.strokeColor)
        .filter((n) => n.type === 'picture' || n.type === 'chart')
        .map((n) => n.type),
    ),
  ]
  const colorNote = colorlessTypes.length
    ? `\n(${colorlessTypes.join('/')} colors not available)`
    : ''
  return `Canvas ${slide.widthPx}×${slide.heightPx}px\n${parts.join('\n---\n') || '(no elements on this page)'}${colorNote}`
}

export function createSlidesSkill(access: DeckAccess): AgentSkill {
  // The HTML pipeline was already used in this conversation → later calls without an explicit mode default to append.
  // Safety net for when the AI ignores the "pass all pages at once" constraint: separate calls no longer overwrite each other (P0-1).
  const state: SkillState = { htmlGenerated: false }
  return {
    id: 'slides',
    systemPrompt: AGENT_SYSTEM_PROMPT,
    tools: TOOLS,
    buildContext: () => {
      const outline = `<deck outline>\n${buildDeckOutline(access.getSlides(), access.getCurrent(), access.getSelectedIds())}\n</deck outline>`
      const progress = buildProgressNote(state)
      return progress ? `${outline}\n${progress}` : outline
    },
    executeTool: (call, signal) => executeTool(access, call, state, signal),
  }
}

interface SkillState {
  htmlGenerated: boolean
  /** A web_search ran in this conversation — unlocks dataSource:'search' in the figure gate */
  webSearched?: boolean
  /** Number of pages most recently planned by plan_deck, used by the progress checklist to remind the AI to finish */
  plannedPages?: number
  /** Per-page titles planned by plan_deck (order = page order), used to name unfinished pages in the checklist */
  plannedTitles?: string[]
  /** Whether each planned page has been generated (aligned with plannedTitles) — names unfinished pages accurately even when a middle page fails */
  pageDone?: boolean[]
  /** Style Skill produced by the most recent generate_deck (used by save_style_template) */
  lastStyleSkill?: string
  /** Topic of the most recent generate_deck (used by save_style_template) */
  lastTopic?: string
}

/**
 * [Hard constraint against "hand-building from scratch"] Sometimes the AI skips the HTML
 * pipeline and assembles a whole deck element by element with add_text_box/add_shape/add_smartart —
 * such hand-built pages look crude and the layout falls apart (root cause of screenshot issues).
 * "From-scratch" detection: this session hasn't used the HTML pipeline (!htmlGenerated) AND the
 * deck has almost no real content (≤ 2 non-decoration elements with text, i.e. blank/initial
 * template). If so, reject and steer toward generate_deck. Adding a single element to an
 * existing rich deck / fine-tuning after the HTML pipeline are unaffected.
 */
function blockScratchBuild(
  toolName: string,
  slides: RenderSlide[],
  state?: SkillState,
): { output: string; isError: true; mutated: false; summary: string } | null {
  if (state?.htmlGenerated) return null // Went through the HTML pipeline; subsequent native edits are legitimate
  let contentEls = 0
  for (const slide of slides) {
    for (const n of collectNodeInfos(slide.nodes)) {
      if (!n.locked && n.text && n.text.trim() !== '') contentEls += 1
    }
  }
  if (contentEls > 2) return null // Deck already has real content; this is a refinement scenario, allow it
  const label = toolName === 'add_smartart' ? t('aiLabelInsertSmartart') : t('aiFailNewElement')
  return {
    output:
      "For blank/from-scratch scenarios don't hand-assemble pages element by element with add_text_box/add_shape/add_smartart (crude layout). " +
      'Use cloud generation instead: new whole deck → generate_deck; new pages for an existing deck → generate_deck(pages, insert_mode:"append"). ' +
      'Write it beautifully in HTML/CSS and the system converts it into editable elements. Use native tools only when the deck already has polished content and one element needs refining.',
    isError: true,
    mutated: false,
    summary: t('aiSumFromScratchGuard', { label }),
  }
}

/**
 * Build the generation progress checklist text — injected to the AI each turn via buildContext
 * so it "sees" which pages are still missing, a mechanical reminder to finish (rather than a
 * one-shot prompt constraint). Returns an empty string when there is no plan.
 */
function buildProgressNote(state?: SkillState): string {
  if (!state || !state.plannedPages) return ''
  const planned = state.plannedPages
  const flags = state.pageDone ?? new Array(planned).fill(false)
  const done = flags.filter(Boolean).length
  const titles = state.plannedTitles ?? []
  if (done >= planned) {
    return `<generation-progress>\n✅ Complete: ${planned} pages planned, all generated (${done} pages).\n</generation-progress>`
  }
  // Name unfinished pages one by one from pageDone (page numbers stay accurate when a middle page fails)
  const remaining: string[] = []
  for (let i = 0; i < planned; i++) {
    if (!flags[i]) remaining.push(`page ${i + 1}${titles[i] ? ` "${titles[i]}"` : ''}`)
  }
  return (
    `<generation-progress>\n` +
    `⚠️ Incomplete: ${planned} pages planned, ${done} generated, ${planned - done} still missing.\n` +
    `Unfinished: ${remaining.join(', ')}.\n` +
    `Immediately fill in the unfinished pages above with generate_deck(pages: briefs for the missing pages, insert_mode:"append"); do not stop and do not substitute native tools.\n` +
    `</generation-progress>`
  )
}

const fail = (summary: string, output: string) => ({
  output,
  isError: true,
  mutated: false,
  summary,
})

// ── Figure-provenance gate ────────────────────────────────────
// Prompt rules ("search before writing data") did not stop invented numbers being
// delivered as fact, so provenance is enforced at the tool layer: chart data and
// data-dense briefs must declare a dataSource, 'search' is only accepted after a
// real web_search in this conversation, and 'sample' figures must be disclosed.

/** Specific figures: percentages, money, magnitude units, decimals — not bare years/counts */
const SPECIFIC_FIGURE_RE =
  /\d[\d,.]*\s*(?:%|％|亿|萬|万|兆|billion|million|\bbn\b|\bmn\b)|[¥￥$€£]\s*\d|\d+\.\d+/g

function countSpecificFigures(text: string): number {
  return text.match(SPECIFIC_FIGURE_RE)?.length ?? 0
}

/**
 * Returns an error message when the declared dataSource does not justify the figures
 * this call carries, null when the call may proceed.
 */
function dataSourceGateError(call: AgentToolCall, state: SkillState | undefined): string | null {
  const src = String(call.input.dataSource ?? '')
  if (src === 'user' || src === 'document' || src === 'sample') return null
  if (src === 'search') {
    if (state?.webSearched) return null
    return (
      "dataSource is 'search' but no web_search has run in this conversation. " +
      'Run web_search first and build the figures from the results (cite them to the user), ' +
      "or declare 'user'/'document' if the figures actually came from the user or this deck."
    )
  }
  return (
    'This call carries specific figures, so dataSource is required: ' +
    "'user' (figures supplied by the user or attachments), 'document' (read from this deck), " +
    "'search' (from web_search results — run it first), or 'sample' (illustrative placeholders; " +
    'you must tell the user they are NOT real data). Never present invented numbers as facts.'
  )
}

/** Appended to a successful tool output when the model declared the figures illustrative. */
const SAMPLE_DATA_NOTE =
  '\nNOTE: dataSource is "sample" — you MUST tell the user these figures are illustrative placeholders, not real data, and offer to research real numbers with web_search.'

/** Template placeholder strings that must never survive into a delivered page */
const PLACEHOLDER_PATTERNS: RegExp[] = [
  /copy paste fonts\. choose the only option to retain text/i,
  /click to edit master title style/i,
  /\btext here\b/i,
  /supporting text here/i,
  /lorem ipsum/i,
]

/**
 * Cheap post-generation audit of one page's HTML: flags leftover template
 * placeholders and near-empty pages so the model is forced to redo them instead of
 * delivering filler as success. Returns a short reason, or null when the page looks fine.
 */
export function auditPageHtml(html: string): string | null {
  const text = html
    .replace(/<(style|script)\b[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
  for (const re of PLACEHOLDER_PATTERNS) {
    const m = re.exec(text)
    if (m) return `contains template placeholder text "${m[0]}"`
  }
  if (text.length < 10) return 'has almost no text content'
  return null
}

/** Append the missing-image report to the tool output: the model learns which pages lack images and how to fix them, instead of silently treating it as success. */
function imageFailNote(fails?: { page: number; url: string }[]): string {
  if (!fails?.length) return ''
  const detail = fails.map((f) => `page ${f.page} (${f.url})`).join(', ')
  return `\n⚠️ Missing images: ${detail} failed to download/convert; those image slots are blank on the page. Re-run image_search with more generic English keywords, pick a working image, patch it onto the page with insert_web_image (slideIndex = page number - 1), then reply to the user.`
}

async function executeTool(
  access: DeckAccess,
  call: AgentToolCall,
  state?: SkillState,
  signal?: AbortSignal,
) {
  const slides = access.getSlides()
  switch (call.name) {
    case 'get_deck_context':
      return {
        output: buildDeckOutline(slides, access.getCurrent(), access.getSelectedIds()),
        mutated: false,
        summary: t('aiSumDeckContext'),
      }

    case 'read_slide': {
      const idx = Number(call.input.slideIndex)
      const slide = slides[idx]
      if (!slide)
        return fail(t('aiFailReadSlide'), `slideIndex out of range (0-${slides.length - 1})`)
      return {
        output: formatSlideDump(slide),
        mutated: false,
        summary: t('aiSumReadSlide', { n: idx + 1 }),
      }
    }

    case 'set_element_text': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      const slide = slides[idx]
      if (!slide)
        return fail(t('aiFailEditText'), `slideIndex out of range (0-${slides.length - 1})`)
      const paragraphs = toEditParagraphs(call.input.paragraphs)
      if (!paragraphs) return fail(t('aiFailEditText'), 'paragraphs must be a non-empty array')
      const target = resolveEditTarget(slide, sourceId)
      const terr = targetError(target, sourceId, idx + 1)
      if (terr || !target || 'nested' in target) return fail(t('aiFailEditText'), terr!)
      const updated = await window.slidesApi.editText({
        slideIndex: idx,
        sourceId,
        paragraphs,
        ...(target.groupId ? { groupId: target.groupId } : {}),
      })
      if (!updated)
        return fail(
          t('aiFailEditText'),
          `Element ${sourceId} (${target.node.type}) does not support text editing` +
            (target.node.type === 'table'
              ? '; use edit_table_cell for tables'
              : target.node.type === 'chart'
                ? '; use edit_chart for charts'
                : ''),
        )
      access.applySlide(idx, updated)
      return {
        output: `Replaced the text of element ${sourceId} on page ${idx + 1} (${paragraphs.length} paragraphs).`,
        mutated: true,
        summary: t('aiSumEditText', { n: idx + 1 }),
      }
    }

    case 'set_element_style': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      const slide = slides[idx]
      if (!slide) return fail(t('aiFailStyle'), `slideIndex out of range (0-${slides.length - 1})`)
      const target = resolveEditTarget(slide, sourceId)
      const terr = targetError(target, sourceId, idx + 1)
      if (terr || !target || 'nested' in target) return fail(t('aiFailStyle'), terr!)
      const node = target.node
      if (!(node.type === 'text' || node.type === 'shape')) {
        return fail(t('aiFailStyle'), `Element ${sourceId} (${node.type}) has no editable text`)
      }
      const cur = nodeToParagraphs(node as ShapeRenderNode)
      if (!cur.length) return fail(t('aiFailStyle'), 'This element has no text to format')
      const ov = call.input as SlideStylePatch
      const paragraphs = mergeStyleIntoParagraphs(cur, ov)
      const updated = await window.slidesApi.editText({
        slideIndex: idx,
        sourceId,
        paragraphs,
        ...(target.groupId ? { groupId: target.groupId } : {}),
      })
      if (!updated)
        return fail(t('aiFailStyle'), `Element ${sourceId} does not support format editing`)
      access.applySlide(idx, updated)
      return {
        output: `Updated the formatting of element ${sourceId} on page ${idx + 1}.`,
        mutated: true,
        summary: t('aiSumStyle', { n: idx + 1 }),
      }
    }

    case 'set_element_transform': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      const slide = slides[idx]
      if (!slide)
        return fail(t('aiFailTransform'), `slideIndex out of range (0-${slides.length - 1})`)
      const target = resolveEditTarget(slide, sourceId)
      const terr = targetError(target, sourceId, idx + 1)
      if (terr || !target || 'nested' in target) return fail(t('aiFailTransform'), terr!)
      const b = target.node.box
      // Group-child render boxes are group-local; the tool takes absolute px, so convert both ways via the group origin
      const origin = target.groupOrigin ?? { x: 0, y: 0 }
      const inp = call.input as {
        x?: number
        y?: number
        w?: number
        h?: number
        rotationDeg?: number
      }
      const updated = await window.slidesApi.editTransform({
        slideIndex: idx,
        sourceId,
        ...(target.groupId ? { groupId: target.groupId } : {}),
        xPx: (typeof inp.x === 'number' ? inp.x : origin.x + b.x) - origin.x,
        yPx: (typeof inp.y === 'number' ? inp.y : origin.y + b.y) - origin.y,
        wPx: typeof inp.w === 'number' ? inp.w : b.w,
        hPx: typeof inp.h === 'number' ? inp.h : b.h,
        rotationDeg: typeof inp.rotationDeg === 'number' ? inp.rotationDeg : b.rotationDeg,
        fitWidthPx: access.fitWidthPx,
      })
      if (!updated) return fail(t('aiFailTransform'), 'Transform failed')
      access.applySlide(idx, updated)
      const afterTarget = resolveEditTarget(updated, sourceId)
      const after = afterTarget && !('nested' in afterTarget) ? afterTarget : null
      const nb = after
        ? {
            ...after.node.box,
            x: (after.groupOrigin?.x ?? 0) + after.node.box.x,
            y: (after.groupOrigin?.y ?? 0) + after.node.box.y,
          }
        : undefined
      const boxStr = nb
        ? `New geometry: pos(${Math.round(nb.x)},${Math.round(nb.y)}) size ${Math.round(nb.w)}×${Math.round(nb.h)}.`
        : ''
      const issues = auditSlideLayout(updated)
      const auditStr = issues.length
        ? `\n⚠️ The layout audit found ${issues.length} issue(s) on this page:\n${issues.map((s) => `- ${s}`).join('\n')}\nFor multi-element layout adjustments switch to execute_slide_script (it reads every element's real geometry and applies atomically).`
        : ''
      return {
        output: `Adjusted the position/size of element ${sourceId} on page ${idx + 1}. ${boxStr}${auditStr}`,
        mutated: true,
        summary: t('aiSumTransform', { n: idx + 1 }),
      }
    }

    // execute_layout_script is a legacy alias (avoids breaking existing sessions/prompts); both share the same logic
    case 'execute_layout_script':
    case 'execute_slide_script': {
      const idx = Number(call.input.slideIndex)
      const slide = slides[idx]
      if (!slide) return fail(t('aiFailScript'), `slideIndex out of range (0-${slides.length - 1})`)
      const code = String(call.input.code ?? '').trim()
      if (!code) return fail(t('aiFailScript'), 'code must not be empty')
      const infos = collectNodeInfos(slide.nodes)
      const r = runLayoutScript(code, infos, { w: slide.widthPx, h: slide.heightPx })
      const logsStr = r.logs.length ? `\nlog output:\n${r.logs.join('\n')}` : ''
      if (r.error) {
        return fail(
          t('aiFailScript'),
          `Script execution error: ${r.error}${logsStr}\n(You can fix the script and retry; see els for the element list and geometry)`,
        )
      }
      const returnedStr = r.returned !== undefined ? `\nScript returned: ${r.returned}` : ''
      if (r.ops.length === 0 && r.edits.length === 0) {
        return {
          output: `Script finished but called no edit primitives (setBox/moveBy/setText/setStyle/setFill/setStroke); the page was not modified.${returnedStr}${logsStr}`,
          mutated: false,
          summary: t('aiSumScriptNoop', { n: idx + 1 }),
        }
      }
      // ── Dispatch: geometry applied atomically once via batchEditTransform, the rest serially in script order,
      //   each step using the returned new slide as the next step's current state. One failure doesn't crash the whole run; report faithfully.
      // Balanced with the end in the finally below; an unbalanced pair would leave
      // the session mid-batch, where undo/redo refuse to run
      const batchOpened = (await window.slidesApi.beginHistoryBatch?.()) === true
      try {
        let current = slide
        const failures: string[] = []
        let boxApplied = 0
        const topOps = r.ops.filter((op) => !op.groupId)
        const grpOps = r.ops.filter((op) => op.groupId)
        if (topOps.length > 0) {
          const updated = await window.slidesApi.batchEditTransform({
            slideIndex: idx,
            fitWidthPx: access.fitWidthPx,
            items: topOps.map((op) => ({
              sourceId: op.id,
              xPx: op.x,
              yPx: op.y,
              wPx: op.w,
              hPx: op.h,
              rotationDeg: op.rotation,
            })),
          })
          if (updated) {
            current = updated
            access.applySlide(idx, updated)
            boxApplied = topOps.length
          } else {
            failures.push(
              `Batch geometry apply failed (${topOps.length} items; some elements may no longer exist on this page)`,
            )
          }
        }
        // Group children go through the per-element in-group transform IPC (abs px → group-local px via the group origin)
        for (const op of grpOps) {
          const gnode = findNodeById(current.nodes, op.groupId!)
          const updated = gnode
            ? await window.slidesApi.editTransform({
                slideIndex: idx,
                sourceId: op.id,
                groupId: op.groupId!,
                xPx: op.x - Math.round(gnode.box.x),
                yPx: op.y - Math.round(gnode.box.y),
                wPx: op.w,
                hPx: op.h,
                rotationDeg: op.rotation,
                fitWidthPx: access.fitWidthPx,
              })
            : null
          if (!updated) {
            failures.push(`setBox("${op.id}"): geometry apply inside group ${op.groupId} failed`)
            continue
          }
          current = updated
          access.applySlide(idx, updated)
          boxApplied += 1
        }
        const counts = { text: 0, style: 0, fill: 0, stroke: 0 }
        for (const e of r.edits) {
          const grp = e.groupId ? { groupId: e.groupId } : {}
          let updated: RenderSlide | null
          if (e.kind === 'text') {
            updated = await window.slidesApi.editText({
              slideIndex: idx,
              sourceId: e.id,
              paragraphs: e.paragraphs,
              ...grp,
            })
            if (!updated) {
              failures.push(
                `setText("${e.id}"): element not found or does not support text editing`,
              )
              continue
            }
          } else if (e.kind === 'style') {
            // Changing style = read current paragraphs (including earlier edit results), merge override fields, write back whole
            const node = findNodeById(current.nodes, e.id)
            if (!node || !(node.type === 'text' || node.type === 'shape')) {
              failures.push(`setStyle("${e.id}"): text element not found`)
              continue
            }
            const cur = nodeToParagraphs(node as ShapeRenderNode)
            if (!cur.length) {
              failures.push(`setStyle("${e.id}"): this element has no text to format`)
              continue
            }
            updated = await window.slidesApi.editText({
              slideIndex: idx,
              sourceId: e.id,
              paragraphs: mergeStyleIntoParagraphs(cur, e.style),
              ...grp,
            })
            if (!updated) {
              failures.push(`setStyle("${e.id}"): element does not support format editing`)
              continue
            }
          } else if (e.kind === 'fill') {
            updated = await window.slidesApi.editFill({
              slideIndex: idx,
              sourceId: e.id,
              fill: e.fill,
              ...grp,
            })
            if (!updated) {
              failures.push(`setFill("${e.id}"): element does not support fill`)
              continue
            }
          } else {
            updated = await window.slidesApi.editStroke({
              slideIndex: idx,
              sourceId: e.id,
              stroke: e.stroke,
              ...grp,
            })
            if (!updated) {
              failures.push(`setStroke("${e.id}"): element does not support stroke`)
              continue
            }
          }
          current = updated
          access.applySlide(idx, updated)
          counts[e.kind] += 1
        }
        const totalApplied = boxApplied + counts.text + counts.style + counts.fill + counts.stroke
        if (totalApplied === 0) {
          return fail(
            t('aiFailScript'),
            `All operations collected by the script failed to apply:\n${failures.map((f) => `- ${f}`).join('\n')}${returnedStr}${logsStr}`,
          )
        }
        const parts: string[] = []
        if (boxApplied > 0) parts.push(`layout ${boxApplied} element(s)`)
        if (counts.text > 0) parts.push(`text ${counts.text} item(s)`)
        if (counts.style > 0) parts.push(`style ${counts.style} item(s)`)
        if (counts.fill > 0) parts.push(`fill ${counts.fill} item(s)`)
        if (counts.stroke > 0) parts.push(`stroke ${counts.stroke} item(s)`)
        const failStr = failures.length
          ? `\n⚠️ ${failures.length} operation(s) failed (the rest took effect):\n${failures.map((f) => `- ${f}`).join('\n')}`
          : ''
        const issues = auditSlideLayout(current)
        return {
          output: `Applied the edit script to page ${idx + 1}: ${parts.join(', ')}.${returnedStr}${logsStr}${failStr}${formatAudit(issues)}`,
          mutated: true,
          summary: t('aiSumScript', { n: idx + 1, count: totalApplied }),
        }
      } finally {
        if (batchOpened) await window.slidesApi.endHistoryBatch?.()
      }
    }

    case 'set_element_fill': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailFill'), `slideIndex out of range (0-${slides.length - 1})`)
      const target = resolveEditTarget(slides[idx]!, sourceId)
      const terr = targetError(target, sourceId, idx + 1)
      if (terr || !target || 'nested' in target) return fail(t('aiFailFill'), terr!)
      const updated = await window.slidesApi.editFill({
        slideIndex: idx,
        sourceId,
        fill: String(call.input.fill),
        ...(target.groupId ? { groupId: target.groupId } : {}),
      })
      if (!updated) return fail(t('aiFailFill'), `Element ${sourceId} does not support fill`)
      access.applySlide(idx, updated)
      return {
        output: `Set the fill of element ${sourceId} on page ${idx + 1}.`,
        mutated: true,
        summary: t('aiSumFill', { n: idx + 1 }),
      }
    }

    case 'set_element_stroke': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailStroke'), `slideIndex out of range (0-${slides.length - 1})`)
      const remove = call.input.remove === true
      const stroke = remove
        ? null
        : { color: String(call.input.color ?? '#000000'), widthPt: Number(call.input.widthPt ?? 1) }
      const target = resolveEditTarget(slides[idx]!, sourceId)
      const terr = targetError(target, sourceId, idx + 1)
      if (terr || !target || 'nested' in target) return fail(t('aiFailStroke'), terr!)
      const updated = await window.slidesApi.editStroke({
        slideIndex: idx,
        sourceId,
        stroke,
        ...(target.groupId ? { groupId: target.groupId } : {}),
      })
      if (!updated) return fail(t('aiFailStroke'), `Element ${sourceId} does not support stroke`)
      access.applySlide(idx, updated)
      return {
        output: `${remove ? 'Removed' : 'Set'} the stroke of element ${sourceId} on page ${idx + 1}.`,
        mutated: true,
        summary: t('aiSumStroke', { n: idx + 1 }),
      }
    }

    case 'web_search': {
      const query = String(call.input.query ?? '').trim()
      if (!query) return fail(t('aiFailWebSearch'), 'query must not be empty')
      const r = await window.slidesApi.webSearch(query, Number(call.input.maxResults) || 6)
      if (state) state.webSearched = true
      // output for the LLM: title+URL+summary (each summary truncated to 120 chars to stay lean)
      const SNIPPET_MAX = 120
      const lines: string[] = []
      if (r.answer) lines.push(`Direct answer: ${r.answer.slice(0, 300)}\n`)
      r.results.forEach((it, i) => {
        const snip =
          it.snippet.length > SNIPPET_MAX ? it.snippet.slice(0, SNIPPET_MAX) + '…' : it.snippet
        lines.push(`${i + 1}. ${it.title}\n   ${it.url}\n   ${snip}`)
      })
      // display side channel: link list (full title+URL for the UI, not in LLM context)
      const display: ToolDisplay = {
        kind: 'links',
        items: r.results.map((it) => ({ url: it.url, title: it.title })),
      }
      return {
        output: lines.join('\n') || '(no results)',
        mutated: false,
        summary: t('aiSumWebSearch', { query, count: r.results.length }),
        display,
      }
    }

    case 'image_search': {
      const query = String(call.input.query ?? '').trim()
      if (!query) return fail(t('aiFailImageSearch'), 'query must not be empty')
      const r = await window.slidesApi.imageSearch(query, Number(call.input.maxResults) || 8)
      // output for the LLM: keep the existing format (the LLM needs to read URLs into image_queries; format unchanged)
      const lines = r.images.map(
        (im, i) =>
          `${i + 1}. ${im.title || '(untitled)'} [${im.width ?? '?'}x${im.height ?? '?'}]\n   ${im.imageUrl}`,
      )
      // display side channel: image list (for UI thumbnails, not in LLM context)
      const display: ToolDisplay = {
        kind: 'images',
        items: r.images.map((im) => ({ url: im.imageUrl, title: im.title || undefined })),
      }
      return {
        output: lines.join('\n') || '(no images)',
        mutated: false,
        summary: t('aiSumImageSearch', { query, count: r.images.length }),
        display,
      }
    }

    case 'generate_image': {
      const prompt = String(call.input.prompt ?? '').trim()
      if (!prompt) return fail(t('aiFailGenImage'), 'prompt must not be empty')
      const refs = Array.isArray(call.input.referenceImageUrls)
        ? (call.input.referenceImageUrls as unknown[]).map(String).filter(Boolean)
        : undefined
      const r = await window.slidesApi.generateImage({
        prompt,
        model: call.input.model ? String(call.input.model) : undefined,
        referenceImageUrls: refs,
        aspectRatio: call.input.aspectRatio ? String(call.input.aspectRatio) : undefined,
      })
      if (!r.url) return fail(t('aiFailGenImage'), r.error ?? 'Generation failed')
      const display: ToolDisplay = {
        kind: 'images',
        items: [{ url: r.url, title: prompt.slice(0, 60) }],
      }
      return {
        output:
          `Image generated, URL: ${r.url}\n` +
          'New imagery: insert it with insert_web_image. If this edits an existing slide picture (e.g. background removal), swap it in place with replace_image instead.',
        mutated: false,
        summary: t('aiSumGenImage', {
          prompt: `${prompt.slice(0, 20)}${prompt.length > 20 ? '…' : ''}`,
        }),
        display,
      }
    }

    case 'analyze_media': {
      const mediaUrls = Array.isArray(call.input.mediaUrls)
        ? (call.input.mediaUrls as unknown[]).map(String).filter(Boolean)
        : []
      const requirements = String(call.input.requirements ?? '').trim()
      if (!mediaUrls.length) return fail(t('aiFailMedia'), 'mediaUrls must not be empty')
      if (!requirements) return fail(t('aiFailMedia'), 'requirements must not be empty')
      const r = await window.slidesApi.analyzeMedia({ mediaUrls, requirements })
      if (!r.text) return fail(t('aiFailMedia'), r.error ?? 'Analysis failed')
      // Analysis text can be very long; truncate to protect context (first 6000 chars are enough to generate deck content)
      const MAX_LEN = 6000
      const text = r.text.length > MAX_LEN ? r.text.slice(0, MAX_LEN) + '\n…(truncated)' : r.text
      return {
        output: text,
        mutated: false,
        summary: t('aiSumParseMedia', { count: mediaUrls.length }),
      }
    }

    case 'insert_web_image': {
      const idx = Number(call.input.slideIndex)
      if (!slides[idx])
        return fail(t('aiFailInsertImage'), `slideIndex out of range (0-${slides.length - 1})`)
      const url = String(call.input.url ?? '')
      if (!/^https?:\/\//.test(url)) return fail(t('aiFailInsertImage'), 'Invalid url')
      const r = await window.slidesApi.insertImageUrl({
        slideIndex: idx,
        url,
        xPx: Number(call.input.x),
        yPx: Number(call.input.y),
        wPx: Number(call.input.w),
        hPx: Number(call.input.h),
        fitWidthPx: access.fitWidthPx,
      })
      if (!r)
        return fail(
          t('aiFailInsertImage'),
          'Download or insertion failed (the image may be inaccessible)',
        )
      access.applySlide(idx, r.slide)
      return {
        output: `Inserted the image on page ${idx + 1}, element id=${r.sourceId}.`,
        mutated: true,
        summary: t('aiSumInsertImage', { n: idx + 1 }),
      }
    }

    case 'crop_image':
    case 'set_picture_opacity':
    case 'replace_image': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      const failKey =
        call.name === 'crop_image'
          ? ('aiFailCropImage' as const)
          : call.name === 'set_picture_opacity'
            ? ('aiFailPictureOpacity' as const)
            : ('aiFailReplaceImage' as const)
      const slide = slides[idx]
      if (!slide) return fail(t(failKey), `slideIndex out of range (0-${slides.length - 1})`)
      const target = resolveEditTarget(slide, sourceId)
      const terr = targetError(target, sourceId, idx + 1)
      if (terr || !target || 'nested' in target) return fail(t(failKey), terr!)
      if (target.node.type !== 'picture')
        return fail(t(failKey), `Element ${sourceId} is not a picture (type: ${target.node.type})`)
      if (target.groupId)
        return fail(
          t(failKey),
          `Element ${sourceId} is inside a group; this tool only supports top-level pictures — ungroup_element first`,
        )

      if (call.name === 'crop_image') {
        const frac = (v: unknown) => Math.min(1, Math.max(0, Number(v) || 0))
        const cl = frac(call.input.l)
        const ct = frac(call.input.t)
        const cr = frac(call.input.r)
        const cb = frac(call.input.b)
        if (cl + cr >= 0.99 || ct + cb >= 0.99)
          return fail(t(failKey), 'Crop removes the whole image (l+r and t+b must be < 1)')
        const srcRect = cl || ct || cr || cb ? { l: cl, t: ct, r: cr, b: cb } : null
        const updated = await window.slidesApi.editPictureSrcRect({
          slideIndex: idx,
          sourceId,
          srcRect,
        })
        if (!updated) return fail(t(failKey), 'Crop failed')
        access.applySlide(idx, updated)
        return {
          output: srcRect
            ? `Cropped picture ${sourceId} on page ${idx + 1} (l=${cl} t=${ct} r=${cr} b=${cb}).`
            : `Removed the crop of picture ${sourceId} on page ${idx + 1}.`,
          mutated: true,
          summary: t('aiSumCropImage', { n: idx + 1 }),
        }
      }

      if (call.name === 'set_picture_opacity') {
        const opacity = Number(call.input.opacity)
        if (!Number.isFinite(opacity) || opacity < 0 || opacity > 1)
          return fail(t(failKey), 'opacity must be between 0 and 1')
        const updated = await window.slidesApi.editPictureOpacity({
          slideIndex: idx,
          sourceId,
          opacity,
        })
        if (!updated) return fail(t(failKey), 'Opacity change failed')
        access.applySlide(idx, updated)
        return {
          output: `Set the opacity of picture ${sourceId} on page ${idx + 1} to ${opacity}.`,
          mutated: true,
          summary: t('aiSumPictureOpacity', { n: idx + 1 }),
        }
      }

      const url = String(call.input.url ?? '')
      if (!/^https?:\/\//.test(url)) return fail(t(failKey), 'Invalid url')
      const updated = await window.slidesApi.replacePictureUrl({
        slideIndex: idx,
        sourceId,
        url,
        ...(call.input.keepCrop ? { keepSrcRect: true } : {}),
      })
      if (!updated)
        return fail(
          t(failKey),
          'Replacement failed (the image may be inaccessible, or the element is not a replaceable picture)',
        )
      access.applySlide(idx, updated)
      return {
        output: `Replaced the image of picture ${sourceId} on page ${idx + 1} in place (frame/z-order/effects kept).`,
        mutated: true,
        summary: t('aiSumReplaceImage', { n: idx + 1 }),
      }
    }

    case 'ask_clarification': {
      if (!access.askClarification)
        return fail(
          t('aiFailClarify'),
          'The current environment does not support questionnaire cards',
        )
      const raw = Array.isArray(call.input.questions) ? call.input.questions : []
      const questions: ClarifyQuestion[] = raw
        .map((q: Record<string, unknown>, i: number) => ({
          id: String(q.id ?? `q${i + 1}`),
          label: String(q.label ?? ''),
          description: q.description ? String(q.description) : undefined,
          options: Array.isArray(q.options)
            ? q.options.map((o: unknown) => String(o)).slice(0, 5)
            : [],
          multi: !!q.multi,
        }))
        .filter((q) => q.label && q.options.length > 0)
      if (questions.length === 0)
        return fail(
          t('aiFailClarify'),
          'questions must be non-empty and every question needs options',
        )
      const r = await access.askClarification(questions)
      if (r.cancelled) {
        return {
          output:
            'The user skipped the questionnaire. Decide the Core Hook and style yourself based on professional judgment and generate directly.',
          mutated: false,
          summary: t('aiSumClarifySkipped'),
        }
      }
      return {
        output: `User questionnaire answers:\n${r.answers}\nDecide the Core Hook and style accordingly, then generate with generate_deck.`,
        mutated: false,
        summary: t('aiSumClarifyDone'),
      }
    }

    case 'plan_deck': {
      const coreHook = String(call.input.core_hook ?? '').trim()
      const style = String(call.input.style ?? '').trim()
      const pages = Array.isArray(call.input.pages) ? call.input.pages : []
      if (!coreHook || !style || pages.length === 0) {
        return fail(t('aiFailPlan'), 'plan_deck requires core_hook + style + non-empty pages')
      }
      // Planning summary echoed back to the user
      const lines = pages.map((p: Record<string, unknown>, i: number) => {
        const q =
          Array.isArray(p.image_queries) && p.image_queries.length
            ? ` [images: ${p.image_queries.length}]`
            : ''
        return `Page ${i + 1} [${String(p.layout ?? '')}] ${String(p.title ?? '')} — ${String(p.brief ?? '').slice(0, 40)}${q}`
      })
      if (state) {
        state.plannedPages = pages.length
        state.plannedTitles = pages.map(
          (p: Record<string, unknown>) => String(p.title ?? '').trim() || 'Untitled',
        )
        state.pageDone = new Array(pages.length).fill(false) // New planning round, reset per-page progress
      }
      const summary = t('aiSumPlan', { count: pages.length, hook: coreHook })
      return {
        output: `Plan confirmed:\nCore Hook: ${coreHook}\nStyle: ${style}\n${lines.join('\n')}\nNow follow this plan and call generate_deck once, passing core_hook + style + all ${pages.length} pages (each page's brief strictly following the plan above). Each turn a <generation-progress> note tells you how many pages remain; do not stop before they are complete.`,
        mutated: false,
        summary,
      }
    }

    case 'regenerate_slide': {
      const idx = Number(call.input.slideIndex)
      if (!slides[idx])
        return fail(t('aiFailRegen'), `slideIndex out of range (0-${slides.length - 1})`)
      if (!access.regenerateSlide || !access.generatePageCloud)
        return fail(
          t('aiFailRegen'),
          'The current environment does not support the page-redo pipeline',
        )
      const brief = String(call.input.brief ?? '').trim()
      if (!brief) return fail(t('aiFailRegen'), 'brief must not be empty')
      // Figure-provenance gate
      if (countSpecificFigures(`${String(call.input.title ?? '')}\n${brief}`) >= 2) {
        const gateErr = dataSourceGateError(call, state)
        if (gateErr) return fail(t('aiFailRegen'), gateErr)
      }
      const regenImages = Array.isArray(call.input.image_urls)
        ? (call.input.image_urls as unknown[]).map(String).filter((u) => /^https?:\/\//.test(u))
        : []
      // Cloud generation, one retry then give up (same semantics as generate_deck pages)
      const backoff = access.retryBackoffMs ?? 2000
      let marker: string | null = null
      let lastErr = ''
      for (let attempt = 0; attempt < 2 && !marker; attempt++) {
        if (attempt > 0 && backoff > 0) await new Promise((r) => setTimeout(r, backoff))
        const res = await access.generatePageCloud({
          pageIndex: idx + 1,
          totalPages: slides.length,
          coreHook: '',
          style: state?.lastStyleSkill ?? '',
          title: String(call.input.title ?? ''),
          brief,
          layout: String(call.input.layout ?? ''),
          images: regenImages,
          canvasW: 1280,
          canvasH: 720,
        })
        if (res.ok && res.marker) marker = res.marker
        else lastErr = res.error ?? t('aiErrUnknown')
      }
      if (!marker)
        return fail(
          t('aiFailRegen'),
          `Cloud page generation failed (2 attempts): ${lastErr}. This is usually a temporary cloud service error — do not keep calling regenerate_slide in a loop. Instead, make the requested changes in place with execute_slide_script / set_element_* (group children are editable too), or tell the user to retry in a few minutes. The page was not modified.`,
        )
      const r = await access.regenerateSlide(idx, marker)
      if (!r.ok)
        return fail(
          t('aiFailRegen'),
          `${r.error || 'Redo failed'} (the page was not modified; retry once, or edit it in place with execute_slide_script)`,
        )
      if (state) state.htmlGenerated = true
      return {
        output:
          `Redid page ${idx + 1} in place from the brief via cloud generation (other pages untouched; the user can undo). Fine-tune afterwards with execute_slide_script / set_element_* tools.` +
          imageFailNote(r.imageFailures),
        mutated: true,
        summary: t('aiSumRegen', { n: idx + 1 }),
      }
    }

    case 'delete_slide': {
      const idx = Number(call.input.slideIndex)
      if (!slides[idx])
        return fail(t('aiFailDeleteSlide'), `slideIndex out of range (0-${slides.length - 1})`)
      if (slides.length <= 1)
        return fail(t('aiFailDeleteSlide'), 'Only one page remains; cannot delete')
      const r = await window.slidesApi.deleteSlide(idx)
      if (!r) return fail(t('aiFailDeleteSlide'), 'Deletion failed')
      access.applyDeck(r, Math.max(0, Math.min(idx, r.length - 1)))
      return {
        output: `Deleted page ${idx + 1}; the deck now has ${r.length} pages. Note that slideIndex of pages after it shifted down by 1.`,
        mutated: true,
        summary: t('aiSumDeleteSlide', { n: idx + 1 }),
      }
    }

    case 'generate_deck': {
      // ── Self-driven pipeline:
      //   1) Plan: use pages if passed; with topic, the tool plans the outline via LLM (batched recursion over threshold) — fixes missing pages at the input side.
      //   2) Generate: batched concurrent cloud page generation (gsk slide_generate, one retry per page), **each batch lands immediately → frontend shows pages one by one**.
      if (!access.generatePageCloud || !(await access.isCloudPageGenEnabled?.().catch(() => false)))
        return fail(
          t('aiFailGenDeck'),
          'Cloud slide generation is unavailable — sign in to Genspark (gsk) first',
        )
      if (!access.generateFromHtml)
        return fail(
          t('aiFailGenDeck'),
          'The current environment does not support the HTML→pptx pipeline',
        )

      // Hard gate: with unread text attachments present, refuse to generate.
      // The prompt already demands reading them first, but prompt rules alone are not
      // enforcement — this check is, and it is fully computable from the tool-call history.
      {
        const unread = access.unreadTextAttachments?.() ?? []
        if (unread.length > 0) {
          return fail(
            t('aiFailGenDeck'),
            `Text attachment(s) not read yet: ${unread.join(', ')}. Read each one with read_attachment first (paginate long files), then call generate_deck again and pass the key facts you read (names, figures, deals, next steps) in the context argument — deck content must come from the attachments, not generic filler.`,
          )
        }
      }

      const canvasW = 1280
      const canvasH = 720
      const insertMode: 'replace' | 'append' =
        call.input.insert_mode === 'append' ? 'append' : 'replace'
      const PLAN_BATCH = 12 // Per-batch planning cap (kept slightly conservative against truncation)
      const GEN_BATCH = 2 // Per-page generation concurrency (opus large output + proxy concurrent streams time out easily; lowered to 2, stability first)
      const BACKOFF_MS = access.retryBackoffMs ?? 2000 // Retry backoff base (rate limits/overload are mostly transient; immediate retries would hit them again)

      let coreHook = String(call.input.core_hook ?? '').trim()
      const style = String(call.input.style ?? '').trim()
      const pages: Array<Record<string, unknown>> = Array.isArray(call.input.pages)
        ? (call.input.pages as Array<Record<string, unknown>>)
        : []
      const topic = String(call.input.topic ?? '').trim()
      const context = String(call.input.context ?? '').trim() || undefined
      // Every per-page request re-sends the context; cap it so N pages don't multiply a huge attachment
      const PAGE_CONTEXT_MAX = 8000
      const pageContext =
        context && context.length > PAGE_CONTEXT_MAX ? context.slice(0, PAGE_CONTEXT_MAX) : context
      const styleTemplateName = String(call.input.style_template ?? '').trim() || undefined

      // Figure-provenance gate: a data-dense request must say where its numbers came from
      {
        const briefText = [
          topic,
          context ?? '',
          ...pages.map((p) => `${String(p.title ?? '')} ${String(p.brief ?? '')}`),
        ].join('\n')
        if (countSpecificFigures(briefText) >= 2) {
          const gateErr = dataSourceGateError(call, state)
          if (gateErr) return fail(t('aiFailGenDeck'), gateErr)
        }
      }

      // ── Step 0: generate the Style Skill independently — one focused LLM call thinking only about the design system, less AI-looking.
      // Prefer the user-passed style as the style preference; without pages, generate a full Style Skill from topic.
      // When full pages+style are passed, respect the user's style (don't regenerate).
      // When style_template is passed, load the template directly and skip Step 0 (no LLM style generation).
      let styleSkill = ''
      if (styleTemplateName && access.loadStyleTemplate) {
        // Preferred: load from a saved template (fail-open: on load failure continue normal generation)
        try {
          const tr = await access.loadStyleTemplate(styleTemplateName)
          if (tr.ok && tr.styleSkill) styleSkill = tr.styleSkill
        } catch {
          /* fail-open */
        }
      }
      // User clicked stop → checked inside each stage loop, abort as soon as possible (pages already landed are kept)
      const cancelled = () => signal?.aborted === true
      const cancelResult = (landed: number, totalPages: number) => ({
        output: `The user stopped generation. ${landed} page(s) already landed${totalPages ? ` (${totalPages} planned)` : ''} and remain on the canvas.`,
        mutated: landed > 0,
        summary: landed > 0 ? t('aiSumStoppedKept', { n: landed }) : t('aiErrStopped'),
      })
      if (cancelled()) return cancelResult(0, 0)

      const needStyleGen =
        !styleSkill && access.generateStyleSkill && (topic || pages.length === 0 || !style)
      if (needStyleGen) {
        access.onProgress?.({
          stage: 'style',
          label: t('aiStageStyle'),
          status: 'running',
          summary: t('aiStageStyleRunning'),
        })
        const styleTopic =
          topic || coreHook || (pages[0] ? String(pages[0].title ?? '') : 'Presentation')
        const sr = await access.generateStyleSkill!({
          topic: styleTopic,
          ...(style ? { styleHint: style } : {}),
          ...(context ? { questionnaire: context } : {}),
          ...(signal ? { signal } : {}),
        })
        if (sr.ok && sr.styleSkill) styleSkill = sr.styleSkill
        access.onProgress?.({
          stage: 'style',
          label: t('aiStageStyle'),
          status: 'done',
          summary: t('aiStageStyleDone'),
        })
      }
      if (!styleSkill) styleSkill = style // Fallback: use the user-passed style, or empty

      // ── Step 1: plan the outline — without pages, plan in-tool from topic (batched recursion over PLAN_BATCH; layouts chosen per the Style Skill).
      const approxForProgress = Math.max(
        1,
        parseInt(String(call.input.approx_pages ?? '0'), 10) || pages.length || 1,
      )
      if (pages.length === 0) {
        const approx = approxForProgress
        if (!topic || !approx) {
          return fail(
            t('aiFailGenDeck'),
            'generate_deck requires [topic + approx_pages] (system plans internally), or pass [core_hook + style + pages] directly.',
          )
        }
        if (!access.planDeckOutline)
          return fail(
            t('aiFailGenDeck'),
            'The current environment does not support internal planning; pass pages directly.',
          )
        access.onProgress?.({
          stage: 'plan',
          label: t('aiStagePlan'),
          done: 0,
          total: approx,
          status: 'running',
          summary: t('aiStagePlanRunning'),
        })
        let planned = 0
        while (planned < approx) {
          if (cancelled()) return cancelResult(0, approx)
          const count = Math.min(PLAN_BATCH, approx - planned)
          const r = await access.planDeckOutline({
            topic,
            count,
            startPage: planned + 1,
            ...(context ? { context } : {}),
            ...(styleSkill ? { styleSkill } : {}),
            ...(planned > 0 && coreHook ? { continueFrom: { coreHook } } : {}),
            ...(signal ? { signal } : {}),
          })
          if (
            !r.ok ||
            !r.outline ||
            !Array.isArray(r.outline.pages) ||
            r.outline.pages.length === 0
          ) {
            if (pages.length === 0) {
              // On failure the progress must be finalized, otherwise the UI spins forever at "Planning outline…"
              access.onProgress?.({
                stage: 'plan',
                label: t('aiStagePlan'),
                done: 0,
                total: approx,
                status: 'error',
                summary: t('aiStagePlanFailed'),
              })
              return fail(t('aiFailGenDeck'), `Planning failed: ${r.error || 'outline is empty'}`)
            }
            break // A later planning batch failed; at least generate what was already planned
          }
          if (planned === 0) {
            coreHook = String(r.outline.core_hook ?? coreHook).trim()
          }
          pages.push(...r.outline.pages)
          planned += r.outline.pages.length
          access.onProgress?.({
            stage: 'plan',
            label: t('aiStagePlan'),
            done: planned,
            total: approx,
            status: 'running',
            summary: t('aiStagePlanProgress', { done: planned, total: approx }),
          })
          if (r.outline.pages.length < count) break // The model returned fewer than requested; stop
        }
        access.onProgress?.({
          stage: 'plan',
          label: t('aiStagePlan'),
          done: pages.length,
          total: pages.length,
          status: 'done',
          summary: t('aiStagePlanDone', { n: pages.length }),
        })
      }

      // ── Step 1.5: in-tool image search —
      // walk every page's image_queries and replace "English keywords (non-URL)" with real image URLs.
      // Entries that are already http(s) URLs are respected upstream, not re-searched; pages whose search failed keep an empty array (fail-open).
      // The same keyword is searched once per deck (fetching several candidates at once); allocation skips already-used URLs to avoid duplicate images across pages.
      if (access.searchImages) {
        const isUrl = (s: string) => /^https?:\/\//i.test(s)
        const normKw = (s: string) => s.toLowerCase().replace(/\s+/g, ' ')
        // Collect deduplicated search keywords across the deck
        const uniqueKeywords: string[] = []
        const seenKw = new Set<string>()
        for (const p of pages) {
          if (!Array.isArray(p.image_queries)) continue
          for (const q of p.image_queries as unknown[]) {
            const s = String(q).trim()
            if (!s || isUrl(s)) continue
            const k = normKw(s)
            if (!seenKw.has(k)) {
              seenKw.add(k)
              uniqueKeywords.push(s)
            }
          }
        }
        // Each keyword is searched only once, fetching several candidates so cross-page dedup can pick unused images
        const candidatesByKw = new Map<string, string[]>()
        const totalSearches = uniqueKeywords.length
        if (totalSearches > 0) {
          access.onProgress?.({
            stage: 'images',
            label: t('aiStageImages'),
            done: 0,
            total: totalSearches,
            status: 'running',
            summary: t('aiStageImagesRunning', { done: 0, total: totalSearches }),
          })
          let imagesDone = 0
          await Promise.all(
            uniqueKeywords.map(async (kw) => {
              if (cancelled()) return
              try {
                const urls = await access.searchImages!(kw, 5)
                candidatesByKw.set(normKw(kw), urls)
              } catch {
                /* fail-open */
              }
              imagesDone++
              access.onProgress?.({
                stage: 'images',
                label: t('aiStageImages'),
                done: imagesDone,
                total: totalSearches,
                status: 'running',
                summary: t('aiStageImagesRunning', { done: imagesDone, total: totalSearches }),
              })
            }),
          )
        }
        // Allocate page by page in order: explicit URLs are kept as-is (and counted as used); keywords pick from candidates, preferring images the deck hasn't used yet
        const usedUrls = new Set<string>()
        for (const p of pages) {
          if (!Array.isArray(p.image_queries)) continue
          for (const q of p.image_queries as unknown[]) {
            const s = String(q).trim()
            if (isUrl(s)) usedUrls.add(s)
          }
        }
        for (const p of pages) {
          if (!Array.isArray(p.image_queries) || p.image_queries.length === 0) continue
          const resolvedUrls: string[] = []
          for (const q of p.image_queries as unknown[]) {
            const s = String(q).trim()
            if (!s) continue
            if (isUrl(s)) {
              resolvedUrls.push(s)
              continue
            }
            const candidates = candidatesByKw.get(normKw(s)) ?? []
            // If all candidates are used, fall back to reusing the first one (an image beats no image)
            const pick = candidates.find((u) => !usedUrls.has(u)) ?? candidates[0]
            if (pick) {
              usedUrls.add(pick)
              resolvedUrls.push(pick)
            }
          }
          p.image_queries = resolvedUrls
        }
        if (totalSearches > 0) {
          access.onProgress?.({
            stage: 'images',
            label: t('aiStageImages'),
            done: totalSearches,
            total: totalSearches,
            status: 'done',
            summary: t('aiStageImagesDone', { n: totalSearches }),
          })
        }
      }

      if (!coreHook) coreHook = topic || 'Presentation'
      if (!styleSkill)
        styleSkill =
          'Unified clean modern style: main background #FFFFFF, main text #1A1A2E, primary accent #2563EB, secondary accent #F59E0B, cards #F8FAFC, borders #E2E8F0; sans-serif fonts; titles 40-56px, body 16-20px'
      const total = pages.length
      if (total === 0) return fail(t('aiFailGenDeck'), 'No pages to generate (the plan is empty).')

      // Store the plan in the progress state (shared with the todolist mechanism; buildContext injects it every turn)
      // Also record styleSkill+topic for the save_style_template tool
      if (state) {
        state.plannedPages = total
        state.plannedTitles = pages.map((p) => String(p.title ?? '').trim() || 'Untitled')
        state.pageDone = new Array(total).fill(false)
        state.lastStyleSkill = styleSkill
        state.lastTopic = topic || coreHook || ''
      }

      // Presentation name: prefer the cover (page 1) title, then the user-entered topic / coreHook.
      // New drafts are saved under this name instead of "Untitled-timestamp".
      const deckName = String(pages[0]?.title ?? '').trim() || topic || coreHook

      // ── Step 2: generate page by page + land as we go (frontend shows pages one by one).
      // The cloud service (gsk slide_generate) writes each page's HTML and converts it to a
      // one-slide pptx; genOne returns a marker and landing reads the bytes.
      // Land strictly in page order: nextToLand pointer; a page lands only when its marker is ready, keeping page order intact.
      const htmlByIndex: (string | null)[] = new Array(total).fill(null)
      // Per-page completion flags (aligned with pages; same reference as state.pageDone, used by buildContext progress injection)
      const doneFlags: boolean[] = state?.pageDone ?? new Array(total).fill(false)
      const genFailed: number[] = [] // Page indexes (0-based) whose HTML generation failed
      const landFailed: number[] = [] // Page indexes (0-based) whose HTML generated but conversion/landing failed
      const degraded: number[] = [] // Page indexes (0-based) that "landed" via the plain-text fallback — must be reported, otherwise dead pages appear silently
      const deckImageFails: { page: number; url: string }[] = [] // Image download/conversion failures (page numbers are deck-global 1-based)
      const pageErrors: (string | undefined)[] = new Array(total).fill(undefined) // Last failure reason per page
      const auditWarns: string[] = [] // Content audit findings: placeholder text / near-empty pages
      let landedPages = 0
      let firstDone = false
      let baseOffset = 0 // Number of existing pages before generated page 0 in the deck (>0 in append mode); used to re-insert retries at their original position
      let nextToLand = 0 // Index of the next page to land (0-based)

      // Initialize per-page progress state (all pages pending)
      const pageProgressItems: PageProgressItem[] = pages.map((p) => ({
        title: String(p.title ?? '').trim() || t('aiPageN', { n: pages.indexOf(p) + 1 }),
        status: 'pending',
      }))

      // Send initial pages progress
      access.onProgress?.({
        stage: 'pages',
        label: t('aiStagePages'),
        done: 0,
        total,
        status: 'running',
        summary: t('aiStagePageRunning', { n: 1, total }),
        pages: [...pageProgressItems],
      })

      const genOne = async (p: Record<string, unknown>, pageIndex: number) => {
        // Mark as running
        pageProgressItems[pageIndex - 1] = {
          ...pageProgressItems[pageIndex - 1]!,
          status: 'running',
        }
        access.onProgress?.({
          stage: 'pages',
          label: t('aiStagePages'),
          done: landedPages,
          total,
          status: 'running',
          summary: t('aiStagePageRunning', { n: pageIndex, total }),
          pages: [...pageProgressItems],
        })
        const images = Array.isArray(p.image_queries)
          ? (p.image_queries as unknown[])
              .map((x) => String(x))
              .filter((x) => /^https?:\/\//.test(x))
          : []
        let lastErr = ''
        // One retry, then the page is skipped (marked failed in the progress card and the
        // final summary) and the rest of the deck keeps generating.
        for (let attempt = 0; attempt < 2; attempt++) {
          if (cancelled()) return null
          if (attempt > 0 && BACKOFF_MS > 0) await new Promise((r) => setTimeout(r, BACKOFF_MS))
          const res = await access.generatePageCloud!({
            pageIndex,
            totalPages: total,
            coreHook,
            style: styleSkill,
            title: String(p.title ?? ''),
            brief: String(p.brief ?? ''),
            layout: String(p.layout ?? ''),
            images,
            ...(pageContext ? { context: pageContext } : {}),
            ...(topic ? { topic } : {}),
            canvasW,
            canvasH,
            ...(signal ? { signal } : {}),
          })
          if (res.ok && res.marker) {
            pageErrors[pageIndex - 1] = undefined
            return res.marker
          }
          lastErr = res.error ?? t('aiErrUnknown')
        }
        pageErrors[pageIndex - 1] = lastErr
        return null
      }

      // Land in page order: starting from nextToLand, land as many as possible (stop at a gap and wait for it to generate).
      const flushLanded = async () => {
        while (nextToLand < total && htmlByIndex[nextToLand] !== null) {
          if (cancelled()) return
          const html = htmlByIndex[nextToLand] as string
          if (html.length > 0) {
            const m: 'replace' | 'append' = firstDone ? 'append' : insertMode
            const r = await access.generateFromHtml!([html], m, deckName)
            if (r.ok) {
              if (r.fallbackReason) {
                degraded.push(nextToLand)
                pageErrors[nextToLand] = r.fallbackReason
              }
              if (r.imageFailures) deckImageFails.push(...r.imageFailures)
              if (!firstDone) baseOffset = r.appendedFrom ?? 0
              landedPages += 1
              firstDone = true
              doneFlags[nextToLand] = true
              pageProgressItems[nextToLand] = { ...pageProgressItems[nextToLand]!, status: 'done' }
              access.onProgress?.({
                stage: 'pages',
                label: t('aiStagePages'),
                done: landedPages,
                total,
                status: 'running',
                summary:
                  landedPages < total
                    ? t('aiStagePageRunning', { n: landedPages + 1, total })
                    : t('aiStageFinishing'),
                pages: [...pageProgressItems],
              })
            } else {
              landFailed.push(nextToLand)
              pageErrors[nextToLand] = t('aiErrLandFailed', { err: r.error ?? t('aiErrUnknown') })
              pageProgressItems[nextToLand] = {
                ...pageProgressItems[nextToLand]!,
                status: 'error',
                error: pageErrors[nextToLand],
              }
            }
          }
          nextToLand += 1
        }
      }

      for (let start = 0; start < total; start += GEN_BATCH) {
        if (cancelled()) break
        const batchIdxs = []
        for (let k = start; k < Math.min(start + GEN_BATCH, total); k++) batchIdxs.push(k)
        const results = await Promise.all(
          batchIdxs.map(async (idx) => ({ idx, html: await genOne(pages[idx], idx + 1) })),
        )
        if (cancelled()) break
        for (const { idx, html } of results) {
          if (html && html.length > 0) {
            htmlByIndex[idx] = html
            const audit = auditPageHtml(html)
            if (audit) auditWarns.push(`page ${idx + 1} ${audit}`)
          } else {
            htmlByIndex[idx] = '' // Empty-string placeholder; doesn't block subsequent landings
            genFailed.push(idx)
            pageProgressItems[idx] = {
              ...pageProgressItems[idx]!,
              status: 'error',
              error: pageErrors[idx],
            }
          }
        }
        // Batch generated → immediately land everything that can land (frontend shows pages one by one)
        await flushLanded()
      }
      if (!cancelled()) await flushLanded() // Finalize

      // ── One re-land round for landing-failed pages (their one-slide pptx already exists, so
      //   landing again is cheap), re-inserted at their original page position with insert_at
      //   (target position = existing-page offset + pages completed before this one).
      //   Generation-failed pages already spent their single retry and stay skipped.
      if (!cancelled()) {
        const retryIdxs = [...new Set(landFailed)].sort((a, b) => a - b)
        for (const idx of retryIdxs) {
          if (cancelled()) break
          const html = htmlByIndex[idx]
          if (!html) continue
          const isFirstLand = !firstDone
          const r = isFirstLand
            ? await access.generateFromHtml!([html], insertMode, deckName)
            : await access.generateFromHtml!(
                [html],
                'insert_at',
                deckName,
                baseOffset + doneFlags.slice(0, idx).filter(Boolean).length,
              )
          if (r.ok) {
            if (isFirstLand) {
              baseOffset = r.appendedFrom ?? 0
              firstDone = true
            }
            landedPages += 1
            doneFlags[idx] = true
            pageErrors[idx] = r.fallbackReason
            if (r.fallbackReason) degraded.push(idx)
            if (r.imageFailures) deckImageFails.push(...r.imageFailures)
            pageProgressItems[idx] = {
              ...pageProgressItems[idx]!,
              status: 'done',
              error: undefined,
            }
            access.onProgress?.({
              stage: 'pages',
              label: t('aiStagePages'),
              done: landedPages,
              total,
              status: 'running',
              summary: t('aiStagePageRestored', { n: idx + 1 }),
              pages: [...pageProgressItems],
            })
          } else {
            pageErrors[idx] = t('aiErrReinsertFailed', { err: r.error ?? t('aiErrUnknown') })
            pageProgressItems[idx] = {
              ...pageProgressItems[idx]!,
              status: 'error',
              error: pageErrors[idx],
            }
          }
        }
      }

      // User stopped midway: keep landed pages and finish honestly (no more retries / no failed-page reporting)
      if (cancelled()) {
        if (state) state.htmlGenerated = landedPages > 0 || state.htmlGenerated
        access.onProgress?.({
          stage: 'done',
          total: landedPages,
          summary: t('aiSumStoppedKept', { n: landedPages }),
        })
        return cancelResult(landedPages, total)
      }

      if (state) state.htmlGenerated = true

      // ── Sidecar persistence (fail-open): write styleSkill to .styleskill.json next to the draft
      if (landedPages > 0 && access.saveSidecar && styleSkill) {
        try {
          await access.saveSidecar({
            topic: topic || coreHook || '',
            styleSkill,
            createdAt: new Date().toISOString(),
          })
        } catch {
          /* fail-open: a sidecar failure doesn't block */
        }
      }

      if (landedPages === 0) {
        access.onProgress?.({
          stage: 'done',
          total: 0,
          summary: t('aiStageAllFailed', { n: total }),
        })
        return fail(
          t('aiFailGenDeck'),
          `All ${total} pages failed to generate; retry or check the AI model configuration.`,
        )
      }

      // Send completion progress event
      access.onProgress?.({
        stage: 'done',
        total: landedPages,
        summary: t('aiStageDoneSummary', { n: landedPages }),
      })

      const note = buildProgressNote(state)
      const progressTail = note ? `\n${note}` : ''
      // Faithfully report all unfinished pages (both HTML generation failures and conversion/landing failures; all already retried once)
      const stillFailed: number[] = []
      for (let i = 0; i < total; i++) if (!doneFlags[i]) stillFailed.push(i + 1)
      const briefErr = (s?: string) => (s ? (s.length > 80 ? `${s.slice(0, 80)}…` : s) : '')
      const okMsg = `Self-driven generation produced ${landedPages}/${total} pages (HTML written page by page, displayed as generated; failed pages were auto-retried).`
      const failDetail = stillFailed
        .map((n) => `page ${n}${pageErrors[n - 1] ? ` (${briefErr(pageErrors[n - 1])})` : ''}`)
        .join(', ')
      const failMsg = stillFailed.length
        ? ` ⚠️ ${failDetail} still failed after retry (these pages are missing from the deck; later pages shifted forward). To fill them in, call generate_deck again with briefs for just those pages and insert_mode:"append", and tell the user the page is at the end.`
        : ' Fine-tune with the set_element_* / add_* tools.'
      // Pages that went through the plain-text fallback: the page exists in the deck but all design is lost; the model must be told explicitly to redo it in place, not silently treat it as success
      const degradedMsg = degraded.length
        ? ` ⚠️ ${[...degraded]
            .sort((a, b) => a - b)
            .map(
              (i) =>
                `page ${i + 1} (canvas slideIndex=${baseOffset + doneFlags.slice(0, i).filter(Boolean).length})`,
            )
            .join(
              ', ',
            )} degraded to a plain-text fallback page after conversion failure (all layout and styling lost): immediately redo these pages in place with regenerate_slide following the original brief, then reply to the user.`
        : ''
      // Content audit: placeholder text / near-empty pages must not be delivered as finished work
      const auditMsg = auditWarns.length
        ? ` ⚠️ Content audit: ${auditWarns.join('; ')}. Do not tell the user the deck is done — redo each flagged page in place with regenerate_slide using real content (from the attachments/context), then reply.`
        : ''
      return {
        output:
          okMsg + failMsg + degradedMsg + auditMsg + imageFailNote(deckImageFails) + progressTail,
        mutated: true,
        summary: t('aiSumDeckGenerated', { done: landedPages, total }),
      }
    }

    case 'add_slide': {
      const src = Number(call.input.sourceIndex)
      if (!slides[src])
        return fail(t('aiFailNewSlide'), `sourceIndex out of range (0-${slides.length - 1})`)
      const r = await window.slidesApi.addSlide({
        sourceIndex: src,
        clearText: call.input.clearText !== false,
        fitWidthPx: access.fitWidthPx,
      })
      if (!r) return fail(t('aiFailNewSlide'), 'Creation failed')
      access.applyDeck(r.slides, r.index)
      return {
        output: `Created page ${r.index + 1} (${r.slides.length} pages total). ✅ Use slideIndex=${r.index} when filling content into this new page (not 1, unless it happens to be 1). To add another page after it, use sourceIndex=${r.slides.length - 1} (current last page) so it appends at the end.`,
        mutated: true,
        summary: t('aiSumNewSlide', { n: r.index + 1 }),
      }
    }

    case 'add_text_box':
    case 'add_shape': {
      const idx = Number(call.input.slideIndex)
      if (!slides[idx])
        return fail(t('aiFailNewElement'), `slideIndex out of range (0-${slides.length - 1})`)
      const scratchBlock = blockScratchBuild(call.name, slides, state)
      if (scratchBlock) return scratchBlock
      const isShape = call.name === 'add_shape'
      const paragraphs = toEditParagraphs(call.input.paragraphs)
      if (!isShape && !paragraphs)
        return fail(t('aiFailNewTextbox'), 'paragraphs must be a non-empty array')
      const kind = isShape ? String(call.input.kind) : 'textbox'
      if (isShape && !/^[a-zA-Z][a-zA-Z0-9]*$/.test(kind)) {
        return fail(t('aiFailNewShape'), `Invalid shape name: ${kind}`)
      }
      const r = await window.slidesApi.addElement({
        slideIndex: idx,
        kind,
        xPx: Number(call.input.x),
        yPx: Number(call.input.y),
        wPx: Number(call.input.w),
        hPx: Number(call.input.h),
        fitWidthPx: access.fitWidthPx,
        ...(paragraphs ? { paragraphs } : {}),
        ...(isShape && call.input.fillColor ? { fillColor: String(call.input.fillColor) } : {}),
      })
      if (!r) return fail(t('aiFailNewElement'), 'Insertion failed')
      access.applySlide(idx, r.slide)
      return {
        output: `Created a new ${isShape ? 'shape' : 'text box'} on page ${idx + 1}, element id=${r.sourceId}.`,
        mutated: true,
        summary: isShape
          ? t('aiSumNewShape', { n: idx + 1 })
          : t('aiSumNewTextbox', { n: idx + 1 }),
      }
    }

    case 'add_chart': {
      const idx = Number(call.input.slideIndex)
      const slide = slides[idx]
      if (!slide) return fail(t('aiFailChart'), `slideIndex out of range (0-${slides.length - 1})`)
      const categories = Array.isArray(call.input.categories)
        ? call.input.categories.map(String)
        : []
      const seriesRaw = Array.isArray(call.input.series) ? call.input.series : []
      const series = seriesRaw
        .map((s) => ({
          name: String((s as { name?: unknown }).name ?? ''),
          values: Array.isArray((s as { values?: unknown }).values)
            ? ((s as { values: unknown[] }).values.map(Number) as number[])
            : [],
        }))
        .filter((s) => s.values.length > 0)
      if (categories.length === 0 || series.length === 0) {
        return fail(t('aiFailChart'), 'Neither categories nor series may be empty')
      }
      const gateErr = dataSourceGateError(call, state)
      if (gateErr) return fail(t('aiFailChart'), gateErr)
      const defW = Math.round(slide.widthPx * 0.62)
      const defH = Math.round(slide.heightPx * 0.62)
      const w = Number(call.input.w) || defW
      const h = Number(call.input.h) || defH
      const r = await window.slidesApi.addChart({
        slideIndex: idx,
        kind: String(call.input.kind) as
          'bar' | 'barStacked' | 'line' | 'area' | 'pie' | 'doughnut',
        ...(call.input.title ? { title: String(call.input.title) } : {}),
        categories,
        series,
        xPx:
          Number.isFinite(Number(call.input.x)) && call.input.x != null
            ? Number(call.input.x)
            : Math.round((slide.widthPx - w) / 2),
        yPx:
          Number.isFinite(Number(call.input.y)) && call.input.y != null
            ? Number(call.input.y)
            : Math.round((slide.heightPx - h) / 2),
        wPx: w,
        hPx: h,
        fitWidthPx: access.fitWidthPx,
      })
      if (!r) return fail(t('aiFailChart'), 'Insertion failed (check kind and data)')
      access.applySlide(idx, r.slide)
      const sampleNote = call.input.dataSource === 'sample' ? SAMPLE_DATA_NOTE : ''
      return {
        output: `Inserted a ${String(call.input.kind)} chart on page ${idx + 1}, element id=${r.sourceId}.${sampleNote}`,
        mutated: true,
        summary: t('aiSumChart', { n: idx + 1 }),
      }
    }

    case 'add_smartart': {
      const scratchBlockSA = blockScratchBuild(call.name, slides, state)
      if (scratchBlockSA) return scratchBlockSA
      const idx = Number(call.input.slideIndex)
      const slide = slides[idx]
      if (!slide)
        return fail(t('aiFailSmartart'), `slideIndex out of range (0-${slides.length - 1})`)
      const items = Array.isArray(call.input.items)
        ? call.input.items.map(String).filter(Boolean)
        : []
      if (items.length < 2) return fail(t('aiFailSmartart'), 'items requires at least 2 entries')
      const defW = Math.round(slide.widthPx * 0.7)
      const defH = Math.round(slide.heightPx * 0.5)
      const w = Number(call.input.w) || defW
      const h = Number(call.input.h) || defH
      const r = await window.slidesApi.addSmartArt({
        slideIndex: idx,
        layout: String(call.input.layout) as AddSmartArtOp['layout'],
        items,
        xPx:
          Number.isFinite(Number(call.input.x)) && call.input.x != null
            ? Number(call.input.x)
            : Math.round((slide.widthPx - w) / 2),
        yPx:
          Number.isFinite(Number(call.input.y)) && call.input.y != null
            ? Number(call.input.y)
            : Math.round((slide.heightPx - h) / 2),
        wPx: w,
        hPx: h,
        fitWidthPx: access.fitWidthPx,
      })
      if (!r) return fail(t('aiFailSmartart'), 'Insertion failed (check layout)')
      access.applySlide(idx, r.slide)
      return {
        output: `Inserted a ${String(call.input.layout)} diagram (${items.length} nodes) on page ${idx + 1}, element id=${r.sourceId}.`,
        mutated: true,
        summary: t('aiSumSmartart', { n: idx + 1 }),
      }
    }

    case 'add_table': {
      const idx = Number(call.input.slideIndex)
      const slide = slides[idx]
      if (!slide) return fail(t('aiFailTable'), `slideIndex out of range (0-${slides.length - 1})`)
      const rows = Number(call.input.rows)
      const cols = Number(call.input.cols)
      if (
        !Number.isInteger(rows) ||
        !Number.isInteger(cols) ||
        rows < 1 ||
        cols < 1 ||
        rows > 30 ||
        cols > 12
      ) {
        return fail(t('aiFailTable'), 'Invalid rows (1-30) / cols (1-12)')
      }
      const defW = Math.round(slide.widthPx * 0.7)
      const defH = Math.round(Math.min(slide.heightPx * 0.6, rows * 40 + 20))
      const w = Number(call.input.w) || defW
      const h = Number(call.input.h) || defH
      const r = await window.slidesApi.addTable({
        slideIndex: idx,
        rows,
        cols,
        xPx:
          Number.isFinite(Number(call.input.x)) && call.input.x != null
            ? Number(call.input.x)
            : Math.round((slide.widthPx - w) / 2),
        yPx:
          Number.isFinite(Number(call.input.y)) && call.input.y != null
            ? Number(call.input.y)
            : Math.round((slide.heightPx - h) / 2),
        wPx: w,
        hPx: h,
        fitWidthPx: access.fitWidthPx,
      })
      if (!r) return fail(t('aiFailTable'), 'Insertion failed')
      let updated = r.slide
      // Fill cells one by one (cells optional; out-of-range parts ignored)
      const cells = Array.isArray(call.input.cells) ? (call.input.cells as unknown[][]) : []
      let filled = 0
      for (let ri = 0; ri < Math.min(cells.length, rows); ri++) {
        const rowCells = Array.isArray(cells[ri]) ? cells[ri]! : []
        for (let ci = 0; ci < Math.min(rowCells.length, cols); ci++) {
          const text = String(rowCells[ci] ?? '')
          if (!text) continue
          const u = await window.slidesApi.editTableCell({
            slideIndex: idx,
            sourceId: r.sourceId,
            row: ri,
            col: ci,
            paragraphs: [{ runs: [{ text }] }],
          })
          if (u) {
            updated = u
            filled++
          }
        }
      }
      access.applySlide(idx, updated)
      return {
        output: `Inserted a ${rows}×${cols} table on page ${idx + 1}, element id=${r.sourceId}${filled ? `, filled ${filled} cell(s) with text` : ''}.`,
        mutated: true,
        summary: t('aiSumTable', { n: idx + 1 }),
      }
    }

    case 'edit_table_cell': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailEditTable'), `slideIndex out of range (0-${slides.length - 1})`)
      const paragraphs = toEditParagraphs(call.input.paragraphs)
      if (!paragraphs) return fail(t('aiFailEditTable'), 'paragraphs must be a non-empty array')
      const row = Number(call.input.row)
      const col = Number(call.input.col)
      const updated = await window.slidesApi.editTableCell({
        slideIndex: idx,
        sourceId,
        row,
        col,
        paragraphs,
      })
      if (!updated)
        return fail(
          t('aiFailEditTable'),
          `Table ${sourceId} not found or cell (${row},${col}) out of range`,
        )
      access.applySlide(idx, updated)
      return {
        output: `Replaced the text of cell (${row},${col}) in table ${sourceId} on page ${idx + 1}.`,
        mutated: true,
        summary: t('aiSumTableCell', { n: idx + 1 }),
      }
    }

    case 'edit_table_structure': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailTableStructure'), `slideIndex out of range (0-${slides.length - 1})`)
      const kind = String(call.input.kind) as
        'insert-row' | 'delete-row' | 'insert-col' | 'delete-col'
      if (!['insert-row', 'delete-row', 'insert-col', 'delete-col'].includes(kind)) {
        return fail(t('aiFailTableStructure'), 'Invalid kind')
      }
      const r = await window.slidesApi.tableStructure({
        slideIndex: idx,
        sourceId,
        kind,
        index: Number(call.input.index),
        ...(call.input.before ? { before: true } : {}),
      })
      if (!r)
        return fail(
          t('aiFailTableStructure'),
          `Operation failed (table ${sourceId} does not exist, index out of range, or the last row/column cannot be deleted)`,
        )
      access.applySlide(idx, r.slide)
      return {
        output: `Applied ${kind} (index=${Number(call.input.index)}) to table ${sourceId} on page ${idx + 1}. The table id may have been updated to ${r.sourceId}.`,
        mutated: true,
        summary: t('aiSumTableStructure', {
          n: idx + 1,
          op: t(
            kind.startsWith('insert')
              ? kind.endsWith('row')
                ? 'aiOpInsertRow'
                : 'aiOpInsertCol'
              : kind.endsWith('row')
                ? 'aiOpDeleteRow'
                : 'aiOpDeleteCol',
          ),
        }),
      }
    }

    case 'edit_table_style': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailTableStyle'), `slideIndex out of range (0-${slides.length - 1})`)
      const op: import('../../shared/ipc').EditTableStyleOp = { slideIndex: idx, sourceId }
      if (call.input.styleName != null) op.styleName = String(call.input.styleName)
      if (call.input.firstRow != null) op.firstRow = Boolean(call.input.firstRow)
      if (call.input.bandRow != null) op.bandRow = Boolean(call.input.bandRow)
      if (call.input.shadingColor != null) op.shadingColor = String(call.input.shadingColor)
      if (call.input.borderColor != null) op.borderColor = String(call.input.borderColor)
      if (call.input.borderWidthPt != null) op.borderWidthPt = Number(call.input.borderWidthPt)
      if (call.input.borderPreset != null)
        op.borderPreset = String(call.input.borderPreset) as 'all' | 'none'
      const updated = await window.slidesApi.editTableStyle(op)
      if (!updated)
        return fail(
          t('aiFailTableStyle'),
          `Operation failed (table ${sourceId} does not exist or is not of type table)`,
        )
      access.applySlide(idx, updated.slide)
      return {
        output: `Updated the style of table ${sourceId} on page ${idx + 1}.`,
        mutated: true,
        summary: t('aiSumTableStyle', { n: idx + 1 }),
      }
    }

    case 'edit_chart': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailChartEdit'), `slideIndex out of range (0-${slides.length - 1})`)
      const op: import('../../shared/ipc').EditChartOp = { slideIndex: idx, sourceId }
      if (call.input.kind != null)
        op.kind = String(call.input.kind) as import('../../shared/ipc').EditChartOp['kind']
      if (Array.isArray(call.input.categories))
        op.categories = (call.input.categories as unknown[]).map(String)
      if (Array.isArray(call.input.series)) {
        const gateErr = dataSourceGateError(call, state)
        if (gateErr) return fail(t('aiFailChartEdit'), gateErr)
        op.series = (call.input.series as Array<{ name: unknown; values: unknown[] }>).map((s) => ({
          name: String(s.name ?? ''),
          values: (Array.isArray(s.values) ? s.values : []).map(Number),
        }))
      }
      if (call.input.colorScheme != null) op.colorScheme = String(call.input.colorScheme)
      if (call.input.title != null) op.title = String(call.input.title)
      if (call.input.legendPos != null)
        op.legendPos = String(
          call.input.legendPos,
        ) as import('../../shared/ipc').EditChartOp['legendPos']
      if (typeof call.input.dataLabels === 'boolean') op.dataLabels = call.input.dataLabels
      if (typeof call.input.gridlines === 'boolean') op.gridlines = call.input.gridlines
      if (call.input.switchRowCol === true) op.switchRowCol = true
      const updated = await window.slidesApi.editChart(op)
      if (!updated)
        return fail(
          t('aiFailChartEdit'),
          `Operation failed (element ${sourceId} does not exist or is not a chart)`,
        )
      access.applySlide(idx, updated.slide)
      const sampleNote = op.series && call.input.dataSource === 'sample' ? SAMPLE_DATA_NOTE : ''
      return {
        output: `Updated chart ${sourceId} on page ${idx + 1}.${sampleNote}`,
        mutated: true,
        summary: t('aiSumChartEdit', { n: idx + 1 }),
      }
    }

    case 'set_slide_background': {
      const idx = Number(call.input.slideIndex)
      const color = String(call.input.color ?? '')
      if (idx !== -1 && !slides[idx])
        return fail(t('aiFailBackground'), `slideIndex out of range (0-${slides.length - 1} or -1)`)
      if (!/^#?[0-9a-fA-F]{6}$/.test(color))
        return fail(t('aiFailBackground'), 'color must be #RRGGBB')
      const r = await window.slidesApi.editBackground({
        slideIndex: idx,
        color: color.startsWith('#') ? color : `#${color}`,
        fitWidthPx: access.fitWidthPx,
      })
      if (!r) return fail(t('aiFailBackground'), 'Setting failed')
      access.applyDeck(r)
      return {
        output:
          idx === -1
            ? `Set the background of all ${r.length} pages to ${color}.`
            : `Set the background of page ${idx + 1} to ${color}.`,
        mutated: true,
        summary: idx === -1 ? t('aiSumBackgroundAll') : t('aiSumBackground', { n: idx + 1 }),
      }
    }

    case 'delete_element': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailDeleteElement'), `slideIndex out of range (0-${slides.length - 1})`)
      // Deletion is top-level only: for group members guide to ungroup instead of a misleading "not found"
      const target = resolveEditTarget(slides[idx]!, sourceId)
      if (target && ('nested' in target || target.groupId)) {
        const gid = 'nested' in target ? undefined : target.groupId
        return fail(
          t('aiFailDeleteElement'),
          `Element ${sourceId} is inside a group${gid ? ` (${gid})` : ''}; call ungroup_element on the group first and then delete it, or delete the whole group`,
        )
      }
      const updated = await window.slidesApi.deleteElement({ slideIndex: idx, sourceId })
      if (!updated)
        return fail(
          t('aiFailDeleteElement'),
          `Element ${sourceId} not found on page ${idx + 1} (ids change after regenerate/ungroup/save; call read_slide for fresh ids)`,
        )
      access.applySlide(idx, updated)
      return {
        output: `Deleted element ${sourceId} from page ${idx + 1}.`,
        mutated: true,
        summary: t('aiSumDeleteElement', { n: idx + 1 }),
      }
    }

    case 'ungroup_element': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      const slide = slides[idx]
      if (!slide)
        return fail(t('aiFailUngroup'), `slideIndex out of range (0-${slides.length - 1})`)
      const node = slide.nodes.find((n) => n.sourceId === sourceId)
      if (!node) {
        return fail(
          t('aiFailUngroup'),
          findNodeById(slide.nodes, sourceId)
            ? `${sourceId} is inside another group; ungroup the outer group first`
            : `Element ${sourceId} not found on page ${idx + 1}`,
        )
      }
      if (node.type !== 'group')
        return fail(t('aiFailUngroup'), `${sourceId} is not a group (type: ${node.type})`)
      if (node.decoration)
        return fail(t('aiFailUngroup'), `${sourceId} is a layout decoration, read-only`)
      const updated = await window.slidesApi.ungroupElement({ slideIndex: idx, sourceId })
      if (!updated) return fail(t('aiFailUngroup'), 'Ungroup failed')
      access.applySlide(idx, updated)
      // Ungrouping rewrites the page and re-ids every element; echo the fresh list so no extra read_slide is needed
      const fresh = collectNodeInfos(updated.nodes)
        .map((n) => `${n.id} | ${n.type}${n.text ? ` | ${preview(n.text)}` : ''}`)
        .join('\n')
      return {
        output: `Ungrouped ${sourceId} on page ${idx + 1} into ${node.children.length} top-level elements. All element ids on this page changed; current elements:\n${fresh}`,
        mutated: true,
        summary: t('aiSumUngroup', { n: idx + 1 }),
      }
    }

    case 'save_style_template': {
      const name = String(call.input.name ?? '').trim()
      if (!name) return fail(t('aiFailSaveTemplate'), 'name must not be empty')
      if (!access.saveStyleTemplate)
        return fail(
          t('aiFailSaveTemplate'),
          'The current environment does not support template saving',
        )
      const styleSkillToSave = state?.lastStyleSkill ?? ''
      const topicToSave = state?.lastTopic ?? ''
      if (!styleSkillToSave)
        return fail(
          t('aiFailSaveTemplate'),
          'The current deck has no Style Skill to save (generate a presentation with generate_deck first)',
        )
      const r = await access.saveStyleTemplate(name, {
        topic: topicToSave,
        styleSkill: styleSkillToSave,
        createdAt: new Date().toISOString(),
      })
      if (!r.ok) return fail(t('aiFailSaveTemplate'), r.error ?? 'Save failed')
      return {
        output: `Saved the style "${name}" as a template; next time pass style_template:"${name}" to reuse it directly.`,
        mutated: false,
        summary: t('aiSumSaveTemplate', { name }),
      }
    }

    case 'list_style_templates': {
      if (!access.listStyleTemplates)
        return fail(
          t('aiFailListTemplates'),
          'The current environment does not support template listing',
        )
      const templates = await access.listStyleTemplates()
      if (templates.length === 0) {
        return {
          output:
            'No saved style templates yet. After generating a deck, call save_style_template(name) to save the current style.',
          mutated: false,
          summary: t('aiSumTemplatesEmpty'),
        }
      }
      const lines = templates.map(
        (t) => `- ${t.name} (topic: ${t.topic || 'unknown'}, saved ${t.createdAt.slice(0, 10)})`,
      )
      return {
        output: `Saved style templates (${templates.length}):\n${lines.join('\n')}\n\nPass style_template:"<template name>" to generate_deck to reuse that style directly.`,
        mutated: false,
        summary: t('aiSumListTemplates', { count: templates.length }),
      }
    }

    default:
      return fail(call.name, `Unknown tool: ${call.name}`)
  }
}
