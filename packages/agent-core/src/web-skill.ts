/**
 * The web + instructions AgentSkill, shared by every app.
 *
 * Each editor has its own preload bridge (`window.desktop`, `window.desktopApi`,
 * `window.slidesApi`, `window.pdfApi`), so the transport is injected rather
 * than imported — that is the only thing that differs between the four, and
 * duplicating the tool definitions four times is how they would otherwise
 * drift apart.
 *
 * Tools:
 *  - `browse_page`  render a URL in the built-in browser and read it
 *  - `extract_pages` pull several URLs as markdown via Tavily (no rendering)
 *  - `load_skill`   fetch the body of one of the user's own skills
 *  - `remember` / `forget`  the agent's own notes on how this user works
 *  - `view_page`    render the open document and look at it
 *
 * Web/image search stays in each app's existing search skill; this module adds
 * the reading half, which is what "browse the web" actually needs.
 */
import { LOAD_SKILL_TOOL, type AppSurface } from './instructions'
import { FORGET_TOOL, MEMORY_TOOL_GUIDANCE, REMEMBER_TOOL } from './memory'

import type { AgentSkill } from './skill'
import type { AgentToolDef } from './types'

/** render the open document so the model can judge layout, not just structure */
export const VIEW_PAGE_TOOL = 'view_page'

export interface BrowsePageBridgeResult {
  ok: boolean
  error?: string
  page?: {
    url: string
    title: string
    text: string
    truncated: boolean
    links?: Array<{ text: string; href: string }>
  }
}

export interface ExtractPagesBridgeResult {
  ok: boolean
  error?: string
  pages?: Array<{ url: string; title: string; content: string }>
  failed?: string[]
}

/** What the renderer must supply; each app maps these onto its own preload. */
export interface WebSkillBridge {
  browsePage(url: string, opts: { includeLinks: boolean }): Promise<BrowsePageBridgeResult>
  extractPages(urls: string[], advanced: boolean): Promise<ExtractPagesBridgeResult>
  /** body of a user skill, already scope-filtered for this surface */
  loadSkill(id: string): Promise<string>
  /** record a preference; false when the text was not worth storing */
  remember(text: string): Promise<boolean>
  /** drop a recorded preference by its exact text; false when nothing matched */
  forget(text: string): Promise<boolean>
  /**
   * Render what the view is currently showing and return it as PNG images, so
   * the model can judge spacing, overflow and balance rather than inferring
   * them from the structure. Absent in apps that have nothing to render.
   *
   * Deliberately takes no page number: capture reads the live DOM, so a
   * specific page would first have to be scrolled or selected into view, and
   * the renderer adapters have no handle on the editor's navigation state. The
   * tool says so rather than quietly returning whichever page happened to be
   * on screen.
   */
  viewPage?(): Promise<ViewPageBridgeResult>
}

export interface ViewPageBridgeResult {
  ok: boolean
  error?: string
  /** raw base64 PNGs, no data: prefix */
  images?: string[]
  /** what was actually captured, e.g. 'slide 3 of 12' */
  label?: string
}

export interface WebSkillOptions {
  bridge: WebSkillBridge
  surface: AppSurface
  /** true once the user has at least one skill in scope; gates the load_skill tool */
  hasUserSkills: () => boolean
  /** prompt text contributed by rules + the skill catalogue, rebuilt each turn */
  instructionsPrompt: () => string
}

const BROWSE_PROMPT = `## Browsing
- \`browse_page\` opens a URL in the built-in browser and returns the rendered text, so it works on pages that only render with JavaScript. Use it when a search snippet is not enough and you need what the page actually says.
- \`extract_pages\` is the cheaper bulk option: it returns several URLs as markdown without rendering them. Prefer it when you already have the links and just need the text.
- Quote or cite what you read; never present a page's claims as your own knowledge.

## Memory
- \`${REMEMBER_TOOL}\` stores one short preference so it survives into later conversations. ${MEMORY_TOOL_GUIDANCE}
- \`${FORGET_TOOL}\` removes one, by its exact recorded wording. Use it when the user says a preference no longer holds.
- Anything already recorded appears under 'What you remember about this user'. Do not re-record what is listed there.

## Looking at the page
- \`${VIEW_PAGE_TOOL}\` renders the open document and returns it as an image, so you can see the layout instead of inferring it from the structure. Use it for questions about spacing, overflow, alignment and balance — and again after a layout change, to check the result rather than assuming it worked.
- It captures what is on screen right now, and takes no page number. To look at a particular page or slide, bring it into view first with the document tools, then call this.
- Reading the content is still the job of the document tools; this is for judging how it looks.`

function browseTools(includeLoadSkill: boolean, includeViewPage: boolean): AgentToolDef[] {
  const tools: AgentToolDef[] = [
    {
      name: 'browse_page',
      description:
        'Open a URL in the built-in browser and return the rendered page text. Handles JavaScript-rendered pages. Use for reading one page in depth.',
      inputSchema: {
        type: 'object',
        properties: {
          url: { type: 'string', description: 'Absolute http(s) URL' },
          includeLinks: {
            type: 'boolean',
            description: 'Also return the in-page links; default false',
          },
        },
        required: ['url'],
      },
    },
    {
      name: 'extract_pages',
      description:
        'Fetch one or more URLs as markdown text without rendering them. Faster than browse_page for several pages at once.',
      inputSchema: {
        type: 'object',
        properties: {
          urls: {
            type: 'array',
            items: { type: 'string' },
            description: 'Absolute http(s) URLs, at most 20',
          },
          advanced: {
            type: 'boolean',
            description: 'Slower, more thorough extraction; default false',
          },
        },
        required: ['urls'],
      },
    },
  ]
  tools.push(
    {
      name: REMEMBER_TOOL,
      description: `Remember one durable preference about how this user wants you to work. ${MEMORY_TOOL_GUIDANCE}`,
      inputSchema: {
        type: 'object',
        properties: {
          text: { type: 'string', description: 'One sentence, in the third person' },
        },
        required: ['text'],
      },
    },
    {
      name: FORGET_TOOL,
      description: 'Forget a preference you recorded earlier, matched on its exact wording.',
      inputSchema: {
        type: 'object',
        properties: {
          text: { type: 'string', description: 'The recorded wording, copied exactly' },
        },
        required: ['text'],
      },
    },
  )
  if (includeViewPage) {
    tools.push({
      name: VIEW_PAGE_TOOL,
      description:
        'Render what the document is currently showing and look at it. Use this to judge layout — spacing, overflow, alignment, balance — and to check a layout change actually landed. Captures the current view only; scroll or select the page you want first.',
      inputSchema: { type: 'object', properties: {} },
    })
  }
  if (includeLoadSkill) {
    tools.push({
      name: LOAD_SKILL_TOOL,
      description:
        "Read the full text of one of the user's skills, listed under 'User skills'. Call this before acting when a task matches a skill.",
      inputSchema: {
        type: 'object',
        properties: { id: { type: 'string', description: 'The skill id from the list' } },
        required: ['id'],
      },
    })
  }
  return tools
}

export function createWebSkill(options: WebSkillOptions): AgentSkill {
  const { bridge, hasUserSkills, instructionsPrompt } = options
  return {
    id: 'web',
    systemPrompt: BROWSE_PROMPT,
    // lazy: composeSkills and the loop both read this per request, so a skill
    // the user adds mid-session becomes callable on the next turn
    get tools(): AgentToolDef[] {
      return browseTools(hasUserSkills(), typeof bridge.viewPage === 'function')
    },
    buildContext: () => instructionsPrompt(),
    executeTool: async (call) => {
      if (call.name === 'browse_page') {
        const url = String(call.input.url ?? '').trim()
        if (!url) return { output: 'url must not be empty', isError: true, summary: 'browse_page' }
        const result = await bridge.browsePage(url, {
          includeLinks: call.input.includeLinks === true,
        })
        if (!result.ok || !result.page) {
          return {
            output: result.error ?? 'Browse failed',
            isError: true,
            summary: `browse ${url}`,
          }
        }
        const { page } = result
        const parts = [`# ${page.title}`, page.url, '', page.text]
        if (page.truncated) parts.push('\n[truncated]')
        if (page.links?.length) {
          parts.push('\n## Links', ...page.links.map((l) => `- ${l.text} — ${l.href}`))
        }
        return { output: parts.join('\n'), mutated: false, summary: `Read ${page.title || url}` }
      }

      if (call.name === 'extract_pages') {
        const raw = Array.isArray(call.input.urls) ? call.input.urls : []
        const urls = raw.map((u) => String(u)).filter(Boolean)
        if (!urls.length) {
          return { output: 'urls must not be empty', isError: true, summary: 'extract_pages' }
        }
        const result = await bridge.extractPages(urls, call.input.advanced === true)
        if (!result.ok) {
          return {
            output: result.error ?? 'Extract failed',
            isError: true,
            summary: 'extract_pages',
          }
        }
        const pages = result.pages ?? []
        const body = pages.map((p) => `# ${p.title}\n${p.url}\n\n${p.content}`).join('\n\n---\n\n')
        const failed = result.failed?.length ? `\n\nFailed: ${result.failed.join(', ')}` : ''
        return {
          output: (body || '(no content)') + failed,
          mutated: false,
          summary: `Extracted ${pages.length} page(s)`,
        }
      }

      if (call.name === VIEW_PAGE_TOOL) {
        if (!bridge.viewPage) {
          return { output: 'This app cannot render pages.', isError: true, summary: 'view_page' }
        }
        const result = await bridge.viewPage()
        if (!result.ok || !result.images?.length) {
          return {
            output: result.error ?? 'Could not render the page',
            isError: true,
            summary: 'view_page',
          }
        }
        const label = result.label ?? 'the page'
        return {
          // the pictures ride a user turn the loop appends; this text is only
          // the tool's own acknowledgement, which is all a tool result can hold
          output: `Rendered ${label}. The image follows.`,
          mutated: false,
          summary: `Viewed ${label}`,
          images: result.images.map((base64) => ({ base64, mime: 'image/png' })),
        }
      }

      if (call.name === REMEMBER_TOOL || call.name === FORGET_TOOL) {
        const text = String(call.input.text ?? '').trim()
        if (!text) {
          return { output: 'text must not be empty', isError: true, summary: call.name }
        }
        const remembering = call.name === REMEMBER_TOOL
        const ok = remembering ? await bridge.remember(text) : await bridge.forget(text)
        // not an error: the model asked for something reasonable and the store
        // declined (too long, or nothing matched). Say so and let it move on.
        return {
          output: ok
            ? remembering
              ? 'Remembered.'
              : 'Forgotten.'
            : remembering
              ? 'Not stored — the text was empty or unusable.'
              : 'Nothing recorded matches that wording exactly.',
          mutated: false,
          summary: remembering ? `Remembered: ${text}` : `Forgot: ${text}`,
        }
      }

      if (call.name === LOAD_SKILL_TOOL) {
        const id = String(call.input.id ?? '').trim()
        const body = await bridge.loadSkill(id)
        return { output: body, mutated: false, summary: `Loaded skill ${id}` }
      }

      return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
    },
  }
}
