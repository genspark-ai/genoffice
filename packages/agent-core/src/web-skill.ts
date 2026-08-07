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
 *
 * Web/image search stays in each app's existing search skill; this module adds
 * the reading half, which is what "browse the web" actually needs.
 */
import { LOAD_SKILL_TOOL, type AppSurface } from './instructions'
import type { AgentSkill } from './skill'
import type { AgentToolDef } from './types'

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
- Quote or cite what you read; never present a page's claims as your own knowledge.`

function browseTools(includeLoadSkill: boolean): AgentToolDef[] {
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
      return browseTools(hasUserSkills())
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

      if (call.name === LOAD_SKILL_TOOL) {
        const id = String(call.input.id ?? '').trim()
        const body = await bridge.loadSkill(id)
        return { output: body, mutated: false, summary: `Loaded skill ${id}` }
      }

      return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
    },
  }
}
