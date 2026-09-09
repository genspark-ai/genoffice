import type { AgentSkill, ToolDisplay } from '@genoffice/agent-core'
import type {
  GroupRenderNode,
  PictureRenderNode,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
} from '@genoffice/pptx-render'
import type { AgentToolCall, AgentToolDef } from '../../shared/ipc'
import { OP_GROUPS, opGuide, opGuideCatalog, opSignatureIndex } from '../../shared/op-docs'
import { auditSlideLayout, formatAudit } from './layout-audit'
import { runLayoutScript, type LayoutScriptElement } from './layout-script'
import { t } from '../i18n/locale'
import systemPrompt from './prompts/system.md?raw'

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
  /** Land generated pages: each pageMarkers entry redeems a one-slide pptx, merged into / replacing the current deck. Returns total page count or an error.
   *  mode="insert_at" inserts a single page at position insertAt (later pages shift) — used to re-insert failed pages at their original position.
   *  On pipeline failure it automatically falls back to element-level mode; fallbackReason explains why (ok is still true).
   *  deckName = presentation name derived from user input, used as the file name when the new draft is saved (instead of "Untitled-timestamp"). */
  landGeneratedPages?(
    pageMarkers: string[],
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
  /** Redo one slide in place: land the marker's page as a replacement for slide slideIndex (other slides untouched; undoable with ⌘Z). */
  regenerateSlide?(
    slideIndex: number,
    marker: string,
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
  /** live predicate (gsk login && cloud-tools toggle, or a BYOK media key); false hides generate_image */
  imageGenAvailable?(): boolean
  /** same for analyze_media */
  mediaAnalysisAvailable?(): boolean
  /**
   * Cloud single-page generation (gsk slide_generate), used by generate_deck's self-driven
   * pipeline: given the unified style + this page's brief/layout/images, the cloud service
   * writes the HTML and converts it to a one-slide pptx. Returns a marker string that goes
   * into a landGeneratedPages pageMarkers slot.
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
   * Local single-page generation (used when cloud is unavailable, e.g. BYOK without gsk):
   * same inputs and marker contract as generatePageCloud, but the page is produced entirely
   * locally — one LLM request writes a structured slide spec and the main process builds it
   * directly into a one-slide pptx (no HTML intermediate).
   */
  generatePageLocal?(args: {
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
  }): Promise<{ ok: boolean; marker?: string; error?: string; imageFailures?: string[] }>
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

const TOOLS: AgentToolDef[] = [
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
    name: 'execute_slide_script',
    description:
      "[Preferred tool for editing a slide's existing elements] Runs your JS edit script against one page; a single script covers: position/size/alignment/distribution/relative nudges/text/style/fill/stroke." +
      ' At run time the script automatically receives the real geometry and text of every element on the page (els) — **no read_slide needed first**; read-write combined, compute from els inside the script.' +
      ' The whole script is one atomic transaction: geometry in one batch, the rest in script order, one undo step; if any operation fails, everything rolls back and the page is unchanged — fix the script and resend it whole. A layout audit (overlap/out-of-bounds/text overflow) is returned at the end.' +
      ' Far more reliable than individual set_element_* calls — coordinate math happens at execution site, not from memory. If the audit reports problems, call this tool again immediately to fix.\n' +
      'Script environment (constrained synchronous JS-like DSL; no external APIs or ambient globals):\n' +
      '- els: array, each item {id,type,text,x,y,w,h,rotation,fontSizePt?,fill?,textColor?,strokeColor?,inGroup?,groupId?,locked?} (pixels, origin top-left; fill/textColor/strokeColor are current colors in #RRGGBB, read-only — write via setFill/setStyle/setStroke; inGroup+groupId=directly editable group child (all primitives work, coordinates absolute as shown); inGroup without groupId=nested in a sub-group, read-only — apply_ops ungroupElement on the outer group first; locked=layout decoration, read-only)\n' +
      '- canvas: {w,h} canvas size (px)\n' +
      "- setBox(id, {x?,y?,w?,h?,rotation?}): set an element's target box, pass only fields to change\n" +
      '- moveBy(id, dx, dy): relative move (left = negative dx, up = negative dy)\n' +
      '- resizeBy(id, dw, dh): relative resize\n' +
      '- setText(id, textOrParagraphs): replace text entirely; pass a string (split into paragraphs by \\n) or a paragraph array (same format as apply_ops setText paragraphs)\n' +
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
      'AI image generation/editing. Text-to-image, or pass referenceImageUrls for image editing; returns an image URL. NEW imagery: insert with insert_web_image. Editing an EXISTING slide picture (background removal/upscaling/etc.): swap it in place with replace_image — do not insert a duplicate. Use for custom illustrations/icons/backgrounds, style-consistent imagery; for real photos/screenshots still use image_search.',
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
            'Optional, defaults to the configured model. Genspark only — specify for special purposes: fal-bria-rmbg=background removal, fal-ai/recraft-clarity-upscale=upscale, flux-pro/outpaint=outpaint, fal-ai/image-editing/text-removal=remove text watermark',
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
      'Analyze media content: understand images/audio/video (video and audio need Genspark or Gemini as the media provider). Pass media URLs (or local file paths) and analysis requirements; returns analysis text. Video supports extracting key points, structure, and time ranges — good for turning user material into usable deck content.',
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
      'Download an image URL obtained from image_search or generate_image and insert it into a page (pixel coordinates). w×h is a layout frame, not a stretch target: the image keeps its aspect ratio, fills the frame, and the overflow is center-cropped (object-fit: cover) — pick the frame for the layout freely.',
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
    name: 'replace_image',
    description:
      'Swap a picture\'s source image for a URL (from image_search or generate_image) in place — position, size, z-order, border and effects all survive. This is the tool for "change/AI-edit this image" flows: e.g. run generate_image with referenceImageUrls for background removal/upscaling/editing, then replace_image with the returned URL. A new image with a different aspect ratio is never stretched: it fills the frame and is center-cropped (object-fit: cover). keepCrop keeps the existing crop window and is only correct when the new image has the same pixel geometry as the old one (e.g. background removal output).',
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
      '[Redo/redesign an existing page] Regenerates the page from your brief and replaces it in place (other pages untouched, undoable).' +
      ' Use when the user says "redo this page / redesign it / try another layout / make it prettier"; don\'t dismantle the page element by element with native tools.' +
      " Flow: first read_slide to get the page's current content, then check neighboring pages in the deck outline to grasp the deck's style;" +
      ' write a detailed brief — what to keep (copy real text/data into the brief verbatim), what to change, and the target layout; the deck style is applied automatically.' +
      ' If the page needs images, image_search first and pass real URLs in image_urls.' +
      ' If generation fails, it is usually a temporary error: do NOT loop retrying — make the concrete changes in place with execute_slide_script instead (or tell the user to try again in a few minutes).',
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
    name: 'apply_ops',
    description:
      '[Canonical edit surface] Apply a list of canonical edit ops as ONE transaction — atomic by default: any failure rolls everything back, nothing is half-applied. Set dry_run:true to validate the plan without touching the deck (rehearse risky batches). This is THE tool for every edit without a dedicated tool (fill, stroke, transform, delete, z-order, grouping, crop, opacity, effects, links, table cells, structure and styling, page background, speaker notes, page delete/move/duplicate, transitions, sections, theme) and for multi-page or many-element batches; a single op is a perfectly fine batch. For one-page layout math prefer execute_slide_script.\n' +
      'Addressing: every op takes target:{slide, el?} — slide = 0-based index or durable "s_<n>"; el = an element id from the outline/read_slide (e_* ids are durable). Group children: put the child id in target.el and add group:"<group id>".\n' +
      'Units are document-space EMU. read_slide reports px and its exact "1 px = N EMU" factor — convert with that N (9525 only on a standard 16:9 deck; other page sizes differ). Font sizes are pt.\n' +
      'Full op reference (the same executor every editing surface uses), one signature per op; ? marks optional fields:\n' +
      opSignatureIndex() +
      '\naddChart additionally takes dataSource:"user"|"document"|"search"|"sample" (figure provenance, checked before the batch runs; "search" needs a web_search in this conversation).' +
      '\nFor field tables, runnable JSON examples and common mistakes call load_guide with the group name(s) above before a batch you have not done before. ' +
      "A failing op's error also returns its exact one-line signature, and dry_run rehearses the whole batch without touching the deck. An unknown op name returns the full vocabulary.",
    inputSchema: {
      type: 'object',
      properties: {
        ops: {
          type: 'array',
          items: { type: 'object' },
          description: 'The op list, applied in order as one transaction (at most 50)',
        },
        dry_run: { type: 'boolean', description: 'Validate the plan only; the deck is untouched' },
        isolation: {
          type: 'string',
          enum: ['atomic', 'per_op'],
          description: 'atomic (default): all-or-nothing. per_op: independent ops, failures skip.',
        },
      },
      required: ['ops'],
    },
  },
  {
    name: 'load_guide',
    description:
      'Load the full documentation of one or more op groups into context: field tables with types and units, runnable JSON examples, common mistakes. Call it before an apply_ops batch that uses ops you have not used in this conversation; several groups can be loaded at once.\n' +
      'Available groups:\n' +
      opGuideCatalog(),
    inputSchema: {
      type: 'object',
      properties: {
        groups: {
          type: 'array',
          items: { type: 'string', enum: [...OP_GROUPS] },
          description: 'Group names to load, e.g. ["element","slide"]',
        },
      },
      required: ['groups'],
    },
  },
]

/** Collect readable text of nodes (including nested group children); returns a list of [sourceId, type, text] */
/** Find one node by id in the node tree (including groups). */
function findNodeById(nodes: RenderNode[], id: string): RenderNode | undefined {
  for (const n of nodes) {
    if (n.sourceId === id || n.durableId === id) return n
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
  const matches = (n: RenderNode) => n.sourceId === sourceId || n.durableId === sourceId
  for (const n of slide.nodes) {
    if (matches(n)) return { node: n }
    if (n.type === 'group') {
      const g = n as GroupRenderNode
      const child = g.children.find(matches)
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

/** Editable element ids of a slide (locked layout decorations excluded), with short text hints. */
function availableIdList(slide: RenderSlide, withText = false): string {
  return collectNodeInfos(slide.nodes)
    .filter((n) => !n.locked)
    .map((n) => (withText && n.text?.trim() ? `${n.id} ("${preview(n.text, 12)}")` : n.id))
    .join(', ')
}

/**
 * Shared not-found / nested-in-subgroup error text for element-targeting tools. Guided: the
 * not-found branch lists the ids that DO exist — models self-correct from the list, but
 * blind-retry a bare "not found".
 */
function targetError(
  target: EditTarget | null,
  sourceId: string,
  pageNo: number,
  slide?: RenderSlide,
): string | null {
  if (!target) {
    const avail = slide ? ` Elements on this page: [${availableIdList(slide)}].` : ''
    return `Element ${sourceId} not found on page ${pageNo}.${avail} (e_* ids are durable across edits/saves; call read_slide when in doubt)`
  }
  if ('nested' in target)
    return `Element ${sourceId} is nested inside a sub-group; ungroup the outer group first (apply_ops ungroupElement), or edit the sub-group as a whole`
  return null
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
      // Durable id when the element's bytes carry one — survives regenerate/
      // ungroup/save, so the AI can keep addressing across turns
      id: n.durableId ?? n.sourceId,
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
      out.push(
        ...collectNodeInfos(g.children, abs.x, abs.y, {
          id: n.durableId ?? n.sourceId,
          topLevel: !parent,
        }),
      )
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
  if (selectedIds.length > 0) {
    const currentSlide = slides[current]
    const selectedRefs = selectedIds.map((id) => {
      const node = currentSlide ? findNodeById(currentSlide.nodes, id) : undefined
      return node?.durableId ?? node?.sourceId ?? id
    })
    lines.push(`User selected elements: ${selectedRefs.join(', ')}`)
  }
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
          ? 'nested in a sub-group (read-only; apply_ops ungroupElement on the outer group to edit)'
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
  // Report the real px→EMU factor: render px carry the viewport scale, so ×9525 only
  // holds for decks whose baseline width is exactly the fit width (standard 16:9 at 1280).
  const pxToEmu = +(9525 / slide.scale).toFixed(2)
  return `Canvas ${slide.widthPx}×${slide.heightPx}px (1 px = ${pxToEmu} EMU)\n${parts.join('\n---\n') || '(no elements on this page)'}${colorNote}`
}

/** tools that need a media provider: Genspark login + cloud tools, or a BYOK media key in Settings */
function hiddenMediaTools(access: DeckAccess): Set<string> {
  const hidden = new Set<string>()
  if (access.imageGenAvailable?.() === false) hidden.add('generate_image')
  if (access.mediaAnalysisAvailable?.() === false) hidden.add('analyze_media')
  return hidden
}

function mediaToolsOffNote(hidden: Set<string>): string {
  if (hidden.size === 0) return ''
  const plural = hidden.size > 1
  return `\n\nNote: ${[...hidden].join(' and ')} ${plural ? 'are' : 'is'} currently unavailable (no image/media provider: signed out of Genspark or cloud tools off, and no media API key in Settings). Do not call or promise ${plural ? 'them' : 'it'}; for imagery use image_search + insert_web_image instead.`
}

export function createSlidesSkill(access: DeckAccess): AgentSkill {
  // The HTML pipeline was already used in this conversation → later calls without an explicit mode default to append.
  // Safety net for when the AI ignores the "pass all pages at once" constraint: separate calls no longer overwrite each other (P0-1).
  const state: SkillState = { htmlGenerated: false }
  return {
    id: 'slides',
    // live like tools: the off-note overrides the prose that still mentions the hidden tools
    get systemPrompt() {
      return systemPrompt + mediaToolsOffNote(hiddenMediaTools(access))
    },
    // live view: the predicates are re-read before every model request
    get tools() {
      const hidden = hiddenMediaTools(access)
      return hidden.size ? TOOLS.filter((t) => !hidden.has(t.name)) : TOOLS
    },
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
 * pipeline and assembles a whole deck element by element with addElement/addSmartArt ops —
 * such hand-built pages look crude and the layout falls apart (root cause of screenshot issues).
 * "From-scratch" detection: this session hasn't used the HTML pipeline (!htmlGenerated) AND the
 * deck has almost no real content (≤ 2 non-decoration elements with text, i.e. blank/initial
 * template). If so, reject and steer toward generate_deck. Adding a single element to an
 * existing rich deck / fine-tuning after the HTML pipeline are unaffected.
 */
function blockScratchBuild(
  kind: 'element' | 'smartart',
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
  const label = kind === 'smartart' ? t('aiLabelInsertSmartart') : t('aiFailNewElement')
  return {
    output:
      "For blank/from-scratch scenarios don't hand-assemble pages element by element with addElement/addSmartArt/addTable/addChart ops (crude layout). " +
      'Use the generation pipeline instead: new whole deck → generate_deck; new pages for an existing deck → generate_deck(pages, insert_mode:"append"). ' +
      'Write it beautifully in HTML/CSS and the system converts it into editable elements. Use insert ops only when the deck already has polished content and one element needs refining.',
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
  return dataSourceGateErrorForInput(call.input, state)
}

/** Same gate over a plain argument object (an apply_ops op carries dataSource as an extra field). */
function dataSourceGateErrorForInput(
  input: Record<string, unknown>,
  state: SkillState | undefined,
): string | null {
  const src = String(input.dataSource ?? '')
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

/** Append the missing-image report to the tool output: the model learns which pages lack images and how to fix them, instead of silently treating it as success. */
function imageFailNote(fails?: { page: number; url: string }[]): string {
  if (!fails?.length) return ''
  const detail = fails.map((f) => `page ${f.page} (${f.url})`).join(', ')
  return `\n⚠️ Missing images: ${detail} failed to download/convert; those image slots are blank on the page. Re-run image_search with more generic English keywords, pick a working image, patch it onto the page with insert_web_image (slideIndex = page number - 1), then reply to the user.`
}

/**
 * Activity-chip label per op for single-op apply_ops runs, so the panel still
 * reads "Edit text" / "Move element" instead of a bare count now that the
 * dedicated tools are gone. Multi-op batches fall back to the count summary.
 */
const OP_LABEL_KEYS: Record<string, string> = {
  setText: 'aiOpEditText',
  setFont: 'aiOpEditFormat',
  setParagraphFormat: 'aiOpEditFormat',
  setTransform: 'aiOpMoveElement',
  setFill: 'aiOpSetFill',
  setStroke: 'aiOpSetStroke',
  deleteElement: 'aiOpDeleteElement',
  ungroupElement: 'aiOpUngroup',
  setPictureSrcRect: 'aiOpCropImage',
  setPictureOpacity: 'aiOpPictureOpacity',
  setNotes: 'aiOpSpeakerNotes',
  deleteSlide: 'aiOpDeleteSlide',
  setBackground: 'aiOpBackground',
  setTableCell: 'aiOpEditTable',
  tableStructure: 'aiOpTableStructure',
  tableMerge: 'aiOpTableStructure',
  setTableStyle: 'aiOpTableStyle',
  duplicateSlide: 'aiOpNewSlide',
  addElement: 'aiOpNewShape',
  addChart: 'aiOpInsertChart',
  addSmartArt: 'aiOpInsertSmartart',
  addTable: 'aiOpInsertTable',
}

/** Localized summary for an apply_ops run: the op's label when exactly one op ran, else the count. */
function applyOpsSummary(ops: unknown[], count: number): string {
  if (count === 1 && ops.length === 1) {
    const op = ops[0] as { op?: unknown; kind?: unknown } | null
    const name = typeof op?.op === 'string' ? op.op : ''
    const key =
      name === 'addElement' && op?.kind === 'textbox' ? 'aiOpNewTextbox' : OP_LABEL_KEYS[name]
    if (key) return t(key as Parameters<typeof t>[0])
  }
  return t('aiSumApplyOps', { count })
}

/** Insert ops subject to the anti-scratch-build guard when the deck is still blank. */
const SCRATCH_GUARDED_OPS = new Set(['addElement', 'addSmartArt', 'addTable', 'addChart'])

interface OpPreflight {
  /** ops to send over the IPC (AI-layer fields stripped) */
  ops: unknown[]
  /** an addChart declared its figures illustrative — the output must say so */
  sampleData: boolean
}

/**
 * apply_ops runs the same policy gates the dedicated tools run before an op
 * reaches the executor: the anti-scratch-build guard for inserts on a blank
 * deck, and figure provenance for chart data. `dataSource` is an AI-layer
 * field the registry does not know; it is checked here and stripped before the
 * ops cross the IPC. Malformed entries pass through untouched so the executor
 * returns its own guided error for them.
 */
function preflightOps(
  opsIn: unknown[],
  slides: RenderSlide[],
  state: SkillState | undefined,
): OpPreflight | ReturnType<typeof fail> {
  const ops: unknown[] = []
  let sampleData = false
  for (const [i, raw] of opsIn.entries()) {
    if (
      typeof raw !== 'object' ||
      raw === null ||
      typeof (raw as { op?: unknown }).op !== 'string'
    ) {
      ops.push(raw)
      continue
    }
    const op = { ...(raw as Record<string, unknown>) }
    const name = op.op as string
    if (SCRATCH_GUARDED_OPS.has(name)) {
      const blocked = blockScratchBuild(
        name === 'addSmartArt' ? 'smartart' : 'element',
        slides,
        state,
      )
      if (blocked) return { ...blocked, output: `ops[${i}] ${name}: ${blocked.output}` }
    }
    if (name === 'addChart') {
      const gateErr = dataSourceGateErrorForInput(op, state)
      if (gateErr) return fail(t('aiFailApplyOps'), `ops[${i}] ${name}: ${gateErr}`)
      if (op.dataSource === 'sample') sampleData = true
      delete op.dataSource
    }
    ops.push(op)
  }
  return { ops, sampleData }
}

/**
 * Layout audit of the pages a transaction touched (numeric slide prefixes of
 * the journal echo), capped so a deck-wide batch does not flood the reply.
 * Empty when every audited page passes.
 */
function auditTouchedPages(
  records: Array<{ target?: string }> | undefined,
  slides: RenderSlide[],
): string {
  const pages = new Set<number>()
  for (const rec of records ?? []) {
    const prefix = rec.target?.split('/')[0] ?? ''
    if (/^\d+$/.test(prefix)) {
      const idx = Number(prefix)
      if (slides[idx]) pages.add(idx)
    }
    if (pages.size >= 4) break
  }
  const failing = [...pages]
    .sort((a, b) => a - b)
    .map((idx) => ({ page: idx + 1, issues: auditSlideLayout(slides[idx]!) }))
    .filter((a) => a.issues.length > 0)
  if (failing.length === 0) return ''
  return (
    `\n<layout-audit>⚠️ Found issue(s) on ${failing.length} page(s):\n` +
    failing.map((a) => `page ${a.page}:\n${a.issues.map((s) => `- ${s}`).join('\n')}`).join('\n') +
    '\n→ Fix issues your edit caused before replying: execute_slide_script on the affected page (it reads live geometry) or another apply_ops with setTransform; at most 2 fix rounds. Issues that already existed and that you did not touch are for your judgment only — do not report them to the user, and never quote element ids in the reply.\n</layout-audit>'
  )
}

async function executeTool(
  access: DeckAccess,
  call: AgentToolCall,
  state?: SkillState,
  signal?: AbortSignal,
) {
  const slides = access.getSlides()
  switch (call.name) {
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
      // ── Dispatch: the whole script is ONE transaction — the collected primitives go
      //   to the main process in a single IPC and apply atomically through the op
      //   executor (any failure rolls the deck back there and returns a guided error).
      const applied = await window.slidesApi.applyEditScript({
        slideIndex: idx,
        fitWidthPx: access.fitWidthPx,
        boxes: r.ops,
        edits: r.edits,
      })
      if (!applied || 'error' in applied) {
        return fail(
          t('aiFailScript'),
          `${applied && 'error' in applied ? applied.error : t('aiErrUnknown')}\n` +
            `The page is unchanged (the script is atomic). Elements on page ${idx + 1}: ` +
            `[${availableIdList(slide, true)}]. Fix the script and resend it whole.${returnedStr}${logsStr}`,
        )
      }
      access.applySlide(idx, applied.slide)
      const counts = { text: 0, style: 0, fill: 0, stroke: 0 }
      for (const e of r.edits) counts[e.kind] += 1
      const parts: string[] = []
      if (r.ops.length > 0) parts.push(`layout ${r.ops.length} element(s)`)
      if (counts.text > 0) parts.push(`text ${counts.text} item(s)`)
      if (counts.style > 0) parts.push(`style ${counts.style} item(s)`)
      if (counts.fill > 0) parts.push(`fill ${counts.fill} item(s)`)
      if (counts.stroke > 0) parts.push(`stroke ${counts.stroke} item(s)`)
      const issues = auditSlideLayout(applied.slide)
      return {
        output: `Applied the edit script to page ${idx + 1}: ${parts.join(', ')}.${returnedStr}${logsStr}${formatAudit(issues)}`,
        mutated: true,
        summary: t('aiSumScript', { n: idx + 1, count: r.ops.length + r.edits.length }),
      }
    }

    case 'web_search': {
      const query = String(call.input.query ?? '').trim()
      if (!query) return fail(t('aiFailWebSearch'), 'query must not be empty')
      const r = await window.slidesApi.webSearch(query, Number(call.input.maxResults) || 6)
      // a backend failure must not read as "no results" — the model would fabricate conclusions
      if (r.method === 'error') {
        return fail(
          t('aiFailWebSearch'),
          `web search failed (service error, not an empty result — you may retry): ${r.error ?? 'unknown error'}`,
        )
      }
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
      // a backend failure must not read as an empty gallery — the model would fabricate image choices
      if (r.method === 'error') {
        return fail(
          t('aiFailImageSearch'),
          `image search failed (service error, not an empty result — you may retry): ${r.error ?? 'unknown error'}`,
        )
      }
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
      // file:// = a BYOK-generated image in the local store (the main process only resolves its own files)
      if (!/^(https?|file):\/\//.test(url)) return fail(t('aiFailInsertImage'), 'Invalid url')
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

    case 'replace_image': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      const failKey = 'aiFailReplaceImage' as const
      const slide = slides[idx]
      if (!slide) return fail(t(failKey), `slideIndex out of range (0-${slides.length - 1})`)
      const target = resolveEditTarget(slide, sourceId)
      const terr = targetError(target, sourceId, idx + 1, slides[idx])
      if (terr || !target || 'nested' in target) return fail(t(failKey), terr!)
      if (target.node.type !== 'picture')
        return fail(t(failKey), `Element ${sourceId} is not a picture (type: ${target.node.type})`)
      if (target.groupId)
        return fail(
          t(failKey),
          `Element ${sourceId} is inside a group; this tool only supports top-level pictures — ungroup it first (apply_ops ungroupElement)`,
        )

      const url = String(call.input.url ?? '')
      if (!/^(https?|file):\/\//.test(url)) return fail(t(failKey), 'Invalid url')
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
      const regenUseCloud =
        !!access.generatePageCloud && !!(await access.isCloudPageGenEnabled?.().catch(() => false))
      if (!access.regenerateSlide || (!regenUseCloud && !access.generatePageLocal))
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
      // One retry then give up (same semantics as generate_deck pages). Both paths return a
      // marker pointing at a one-slide pptx temp file; landing is shared.
      const backoff = access.retryBackoffMs ?? 2000
      let marker: string | null = null
      let genImageFails: string[] = []
      let lastErr = ''
      const regenArgs = {
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
      }
      const regenGen = regenUseCloud ? access.generatePageCloud! : access.generatePageLocal!
      for (let attempt = 0; attempt < 2 && !marker; attempt++) {
        if (attempt > 0 && backoff > 0) await new Promise((r) => setTimeout(r, backoff))
        const res = await regenGen(regenArgs)
        if (res.ok && res.marker) {
          marker = res.marker
          if ('imageFailures' in res && Array.isArray(res.imageFailures))
            genImageFails = res.imageFailures
        } else lastErr = res.error ?? t('aiErrUnknown')
      }
      if (!marker)
        return fail(
          t('aiFailRegen'),
          `Page generation failed (2 attempts): ${lastErr}. This is usually a temporary service error — do not keep calling regenerate_slide in a loop. Instead, make the requested changes in place with execute_slide_script / set_element_* (group children are editable too), or tell the user to retry in a few minutes. The page was not modified.`,
        )
      const r = await access.regenerateSlide(idx, marker)
      if (!r.ok)
        return fail(
          t('aiFailRegen'),
          `${r.error || 'Redo failed'} (the page was not modified; retry once, or edit it in place with execute_slide_script)`,
        )
      if (state) state.htmlGenerated = true
      const allImageFails = [
        ...(r.imageFailures ?? []),
        ...genImageFails.map((url) => ({ page: idx + 1, url })),
      ]
      return {
        output:
          `Redid page ${idx + 1} in place from the brief (other pages untouched; the user can undo). Fine-tune afterwards with execute_slide_script / set_element_* tools.` +
          imageFailNote(allImageFails.length ? allImageFails : undefined),
        mutated: true,
        summary: t('aiSumRegen', { n: idx + 1 }),
      }
    }

    case 'generate_deck': {
      // ── Self-driven pipeline:
      //   1) Plan: use pages if passed; with topic, the tool plans the outline via LLM (batched recursion over threshold) — fixes missing pages at the input side.
      //   2) Generate: batched concurrent page generation (one retry per page), **each batch lands immediately → frontend shows pages one by one**.
      //      Cloud (gsk slide_generate) when available; otherwise fully local — the LLM (app AI
      //      transport, works with BYOK) writes a slide spec that is built directly into a pptx.
      const useCloud =
        !!access.generatePageCloud && !!(await access.isCloudPageGenEnabled?.().catch(() => false))
      if (!useCloud && !access.generatePageLocal)
        return fail(
          t('aiFailGenDeck'),
          'No page generation pipeline is available in this environment',
        )
      if (!access.landGeneratedPages)
        return fail(
          t('aiFailGenDeck'),
          'The current environment does not support the page landing pipeline',
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
      // Cloud (gsk slide_generate) and local (LLM spec → pptx-engine build) both produce a
      // one-slide pptx temp file; genOne returns its marker and landing reads the bytes.
      // Land strictly in page order: nextToLand pointer; a page lands only when its marker is ready, keeping page order intact.
      const markerByIndex: (string | null)[] = new Array(total).fill(null)
      // Per-page completion flags (aligned with pages; same reference as state.pageDone, used by buildContext progress injection)
      const doneFlags: boolean[] = state?.pageDone ?? new Array(total).fill(false)
      const genFailed: number[] = [] // Page indexes (0-based) whose page generation failed
      const landFailed: number[] = [] // Page indexes (0-based) whose HTML generated but conversion/landing failed
      const degraded: number[] = [] // Page indexes (0-based) that "landed" via the plain-text fallback — must be reported, otherwise dead pages appear silently
      const deckImageFails: { page: number; url: string }[] = [] // Image download/conversion failures (page numbers are deck-global 1-based)
      const pageErrors: (string | undefined)[] = new Array(total).fill(undefined) // Last failure reason per page
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
        const pageArgs = {
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
        }
        // Both paths return a marker pointing at a one-slide pptx temp file. One retry, then the
        // page is skipped for now (locally-failed pages get one more chance in the retry round)
        // and the rest of the deck keeps generating.
        const gen = useCloud ? access.generatePageCloud! : access.generatePageLocal!
        for (let attempt = 0; attempt < 2; attempt++) {
          if (cancelled()) return null
          if (attempt > 0 && BACKOFF_MS > 0) await new Promise((r) => setTimeout(r, BACKOFF_MS))
          const res = await gen(pageArgs)
          if (res.ok && res.marker) {
            pageErrors[pageIndex - 1] = undefined
            if ('imageFailures' in res && Array.isArray(res.imageFailures))
              deckImageFails.push(...res.imageFailures.map((url) => ({ page: pageIndex, url })))
            return res.marker
          }
          lastErr = res.error ?? t('aiErrUnknown')
        }
        pageErrors[pageIndex - 1] = lastErr
        return null
      }

      // Land in page order: starting from nextToLand, land as many as possible (stop at a gap and wait for it to generate).
      const flushLanded = async () => {
        while (nextToLand < total && markerByIndex[nextToLand] !== null) {
          if (cancelled()) return
          const marker = markerByIndex[nextToLand] as string
          if (marker.length > 0) {
            const m: 'replace' | 'append' = firstDone ? 'append' : insertMode
            const r = await access.landGeneratedPages!([marker], m, deckName)
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
          batchIdxs.map(async (idx) => ({ idx, marker: await genOne(pages[idx], idx + 1) })),
        )
        if (cancelled()) break
        for (const { idx, marker } of results) {
          if (marker && marker.length > 0) {
            markerByIndex[idx] = marker
          } else {
            markerByIndex[idx] = '' // Empty-string placeholder; doesn't block subsequent landings
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

      // ── One retry round for failed pages, re-inserted at their original page position with
      //   insert_at (target position = existing-page offset + pages completed before this one).
      //   Landing-failed pages re-land (cheap: the one-slide pptx already exists). Cloud
      //   generation-failed pages already spent their single retry and stay skipped; local
      //   generation-failed pages get one more generation attempt here (LLM calls are the
      //   user's own quota, and a JSON spec retry is cheap).
      if (!cancelled()) {
        const retryIdxs = [...new Set([...(useCloud ? [] : genFailed), ...landFailed])].sort(
          (a, b) => a - b,
        )
        for (const idx of retryIdxs) {
          if (cancelled()) break
          let marker = markerByIndex[idx]
          if (!marker && !useCloud) marker = await genOne(pages[idx]!, idx + 1)
          if (!marker) {
            pageProgressItems[idx] = {
              ...pageProgressItems[idx]!,
              status: 'error',
              error: pageErrors[idx],
            }
            continue
          }
          const isFirstLand = !firstDone
          const r = isFirstLand
            ? await access.landGeneratedPages!([marker], insertMode, deckName)
            : await access.landGeneratedPages!(
                [marker],
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
      return {
        output: okMsg + failMsg + degradedMsg + imageFailNote(deckImageFails) + progressTail,
        mutated: true,
        summary: t('aiSumDeckGenerated', { done: landedPages, total }),
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

    case 'apply_ops': {
      const opsIn = Array.isArray(call.input.ops) ? (call.input.ops as unknown[]) : null
      if (!opsIn || opsIn.length === 0)
        return fail(t('aiFailApplyOps'), 'ops must be a non-empty array')
      const pre = preflightOps(opsIn, slides, state)
      if (!('ops' in pre)) return pre
      const r = await window.slidesApi.applyTxn?.({
        ops: pre.ops,
        ...(call.input.dry_run === true ? { dryRun: true } : {}),
        ...(call.input.isolation === 'per_op' ? { isolation: 'per_op' as const } : {}),
      })
      if (!r) return fail(t('aiFailApplyOps'), 'The transaction could not be executed')
      const failLines = (r.failures ?? []).map((f) => `- ops[${f.index}]: ${f.error}`).join('\n')
      if (r.dryRun) {
        const planStr = (r.plan ?? []).join('\n')
        return {
          output:
            `Dry run — the deck was NOT modified.\n` +
            (planStr ? `Valid plan:\n${planStr}\n` : '') +
            (failLines ? `Rejected ops:\n${failLines}\n` : 'All ops validated.\n') +
            `Resend with dry_run omitted to apply.`,
          mutated: false,
          summary: t('aiSumApplyOpsDryRun'),
        }
      }
      if (!r.applied) {
        return fail(
          t('aiFailApplyOps'),
          `Nothing was applied${call.input.isolation === 'per_op' ? '' : ' (atomic)'}:\n${failLines || 'unknown failure'}`,
        )
      }
      // Ops can delete slides — clamp the current index into the rebuilt deck
      // (page deletion may leave the current index past the end)
      access.applyDeck(r.slides!, Math.max(0, Math.min(access.getCurrent(), r.slides!.length - 1)))
      const created = (r.records ?? []).flatMap((rec) => rec.created ?? [])
      const doneStr = (r.records ?? [])
        .map((rec) => `${rec.op}${rec.target ? ` @${rec.target}` : ''}`)
        .join(', ')
      return {
        output:
          `Applied ${r.records?.length ?? 0} op(s): ${doneStr}.` +
          (created.length ? ` New element ids: ${created.join(', ')}.` : '') +
          (failLines ? `\nSkipped (per_op):\n${failLines}` : '') +
          (pre.sampleData ? SAMPLE_DATA_NOTE : '') +
          auditTouchedPages(r.records, r.slides!),
        mutated: true,
        summary: applyOpsSummary(pre.ops, r.records?.length ?? 0),
      }
    }

    case 'load_guide': {
      const raw = Array.isArray(call.input.groups) ? call.input.groups.map(String) : []
      if (raw.length === 0) return fail(t('aiFailLoadGuide'), 'groups must be a non-empty array')
      const unknown = raw.filter((g) => !opGuide(g))
      if (unknown.length > 0) {
        return fail(
          t('aiFailLoadGuide'),
          `Unknown guide group(s): ${unknown.join(', ')}. Available: ${OP_GROUPS.join(', ')}.`,
        )
      }
      const groups = [...new Set(raw)]
      return {
        output: groups.map((g) => opGuide(g)!).join('\n\n---\n\n'),
        mutated: false,
        summary: t('aiSumLoadGuide', { names: groups.join(', ') }),
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
