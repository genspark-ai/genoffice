import type { AgentToolDef } from './types'
import type { AgentSkill } from './skill'

/**
 * The user's indexed files (the local full-text index: save folder, folder
 * roots, recents and starred files) as a searchable, readable body. The
 * per-turn listing is a bounded window onto that corpus (starred files first,
 * then recent ones); search reaches the whole index.
 *
 * The skill is storage-agnostic on purpose: deps inject the three main-process
 * operations (list / search / read), so it carries no electron or index
 * dependency and is testable with mocks. The shell's knowledge-base gateway
 * (apps/shell/src/main/knowledge-base.ts) is the intended backing; the docs
 * panel is the first consumer.
 *
 * Citation contract: when an answer draws on corpus content it must cite the
 * source as [file name](filenav:///abs/path) — the same primitive the editors'
 * in-document citations use (see @genoffice/ui file-nav). verifyResponse
 * rejects answers citing paths outside the corpus, the way fabricated
 * block numbers are rejected elsewhere. Since search reaches beyond the
 * listed window, paths this skill's own search tool returned this session are
 * whitelisted too — the prompt already tells the model to cite only links from
 * the list or a search result.
 */

/** characters of corpus listing injected before it degrades to names only */
const KB_CONTEXT_MAX_CHARS = 12_000

export interface KbFileInfo {
  /** absolute path — also the citation target */
  path: string
  /** owning folder name ("contracts" in /Users/me/contracts/lease.docx) */
  folder: string
  name: string
  ext: string
  sizeBytes: number
}

export interface KbSearchHit {
  path: string
  name: string
  ext: string
  folder: string
  /** plain text around the first match (marks stripped) */
  snippet: string
}

export interface KbReadPage {
  text: string
  /** total parsed characters — page on with offset when more remains */
  totalChars: number
  offset: number
}

export interface KnowledgeBaseDeps {
  /** the prompt's listing window as of this turn (a synchronous snapshot; the panel refreshes it on the corpus IPC) */
  listFiles(): KbFileInfo[]
  search(query: string, limit: number): Promise<KbSearchHit[]>
  /** one page of parsed text; path must belong to the corpus (the main process checks index membership) */
  read(path: string, offset: number): Promise<KbReadPage>
  /** paths outside the corpus an answer may still cite (files a sibling skill wrote this session) */
  alsoCitable?(): readonly string[]
}

/** keep in sync with FILE_NAV_SCHEME in @genoffice/ui (this package must not
 *  depend on @genoffice/ui: the ui barrel reaches react) */
const FILE_NAV_SCHEME = 'filenav://'

export function formatKbSize(bytes: number): string {
  if (bytes < 1024) return `${bytes}B`
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)}KB`
  return `${(bytes / 1024 / 1024).toFixed(1)}MB`
}

const KB_SYSTEM_PROMPT = `## Knowledge base
The user's local files form a knowledge base: the full-text index covers the
save folder, the added folder roots, and recent/starred files. Each turn's
context lists a window of it (index | folder | file | type | size, with the
absolute path) — the list is not the whole corpus, search reaches all of it.
- Use it only when the request refers to those files or their contents; drafting,
  rewriting or formatting the open document from the request alone never needs it.
- When the request touches the knowledge base, search it first (search_knowledge_base)
  and read the promising files (read_knowledge_file, paged by offset) before
  answering. Never answer corpus questions from file names alone.
- A full-text search matches literal words only: if it returns nothing, retry
  with synonyms or the other language before
  concluding the corpus does not cover the topic.
- Long files are paged: the result reports totalChars and the next offset; keep
  reading with offset while it matters.
- Cite what you used: right after a statement drawn from a file, write
  [file name](cite link) — copy the cite link exactly as the list or a search
  result gives it (it is already encoded). Cite ONLY links that appear there;
  never invent or edit one.
- Say so plainly when the knowledge base does not contain the answer.`

/**
 * Same encoding as @genoffice/ui file-nav (agent-core must not depend on the
 * renderer package): everything the inline-link grammar rejects — spaces, %,
 * CJK, parens, Windows backslashes — is percent-encoded, slashes stay.
 */
export function kbCiteHref(path: string): string {
  const encoded = encodeURIComponent(path)
    .replace(/%2F/g, '/')
    .replace(/\(/g, '%28')
    .replace(/\)/g, '%29')
  return FILE_NAV_SCHEME + encoded
}

const KB_EMPTY =
  'The knowledge base is empty (nothing indexed yet). Do not search it again; work from the request and the open document.'

/** how many search-result paths the citation whitelist remembers per session */
const KB_CITABLE_SEARCH_PATHS_MAX = 2000

export function createKnowledgeBaseSkill(deps: KnowledgeBaseDeps): AgentSkill {
  // Paths this skill's search returned, so verifyResponse can whitelist what
  // the model was actually shown — search reaches the whole index, far beyond
  // the listed window, so listFiles alone cannot vouch for a valid citation.
  const citableSearchPaths = new Set<string>()
  const readTool: AgentToolDef = {
    name: 'read_knowledge_file',
    description:
      'Read a page of a knowledge-base file’s parsed text (docx/pdf/xlsx/pptx/text are parsed locally). ' +
      'Address the file by its index in the knowledge-base list, or by its absolute path from a search result. ' +
      'Read offset=0 first; continue with the returned offset while totalChars says more remains.',
    inputSchema: {
      type: 'object',
      properties: {
        index: {
          type: 'integer',
          description: 'file index from the knowledge-base list (0-based)',
        },
        path: {
          type: 'string',
          description:
            'absolute path from a search result — use this when the file is not in the listed window',
        },
        offset: { type: 'integer', description: 'character offset to read from; default 0' },
      },
      required: [],
    },
  }
  const searchTool: AgentToolDef = {
    name: 'search_knowledge_base',
    description:
      'Full-text search over the knowledge base (the user’s indexed local files), for requests that refer to those files; not for writing or formatting from the request alone. ' +
      'Returns the matching files with a snippet around the first hit. Literal-word matching: ' +
      'retry with synonyms or the other language when it returns nothing.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'search words' },
        limit: { type: 'integer', description: 'max results, default 8, at most 20' },
      },
      required: ['query'],
    },
  }
  const empty = () => deps.listFiles().length === 0
  return {
    id: 'knowledge-base',
    // nothing starred: no prompt section and no tools, so a model that searches
    // on reflex (BYOK/local endpoints did, repeatedly) has nothing to call
    get systemPrompt() {
      return empty() ? '' : KB_SYSTEM_PROMPT
    },
    get tools() {
      return empty() ? [] : [readTool, searchTool]
    },
    buildContext: () => {
      const files = deps.listFiles()
      if (files.length === 0) return ''
      const plain = files.map(
        (f, i) => `${i} | ${f.folder} | ${f.name} | .${f.ext} | ${formatKbSize(f.sizeBytes)}`,
      )
      const paths = files.map((f, i) => `${i} = ${f.path}\n    cite: ${kbCiteHref(f.path)}`)
      const withPaths = `${plain.join('\n')}\n\nAbsolute paths and cite links:\n${paths.join('\n')}`
      // a large corpus degrades to names only: the model searches first anyway,
      // and search results carry the real paths to cite
      const body =
        withPaths.length <= KB_CONTEXT_MAX_CHARS
          ? withPaths
          : `${plain.join('\n')}\n\n(The corpus is large; cite links come with search results.)`
      return `Knowledge base — a window of the user's indexed files (index | folder | file | type | size):\n${body}`
    },
    executeTool: async (call) => {
      if (deps.listFiles().length === 0) {
        return { output: KB_EMPTY, isError: true, summary: 'knowledge base empty' }
      }
      if (call.name === 'search_knowledge_base') {
        const query = typeof call.input.query === 'string' ? call.input.query : ''
        if (!query.trim()) {
          return { output: 'query must not be empty', isError: true, summary: 'search ""' }
        }
        const limitRaw = Number(call.input.limit)
        const limit = Number.isFinite(limitRaw)
          ? Math.min(Math.max(1, Math.floor(limitRaw)), 20)
          : 8
        const hits = await deps.search(query, limit)
        if (hits.length === 0)
          return { output: 'no match. Retry with synonyms or the other language.', summary: query }
        // whitelist these paths for citations: the model is allowed to cite
        // exactly what a search result showed it, listed window or not
        for (const h of hits) {
          citableSearchPaths.add(h.path)
          if (citableSearchPaths.size > KB_CITABLE_SEARCH_PATHS_MAX) {
            const oldest = citableSearchPaths.values().next().value
            if (oldest !== undefined) citableSearchPaths.delete(oldest)
          }
        }
        const lines = hits.map(
          (h) =>
            `${h.folder}/${h.name}\n  path: ${h.path}\n  cite: ${kbCiteHref(h.path)}\n  snippet: ${h.snippet}`,
        )
        return { output: lines.join('\n'), summary: `${query} → ${hits.length}` }
      }
      if (call.name === 'read_knowledge_file') {
        const path = typeof call.input.path === 'string' ? call.input.path.trim() : ''
        const offsetRaw = Number(call.input.offset)
        const offset = Number.isFinite(offsetRaw) ? offsetRaw : 0
        // a search hit outside the listed window reads by path; the main-side
        // kb:read gate (index membership) is what confines this to the corpus
        if (path) {
          const page = await deps.read(path, offset)
          const more = page.offset + page.text.length < page.totalChars
          return {
            output:
              `${path} — characters ${page.offset}..${page.offset + page.text.length} of ${page.totalChars}` +
              (more ? ' (continue with offset)' : ' (end)') +
              `\n\n${page.text}`,
            summary: `${path}@${page.offset}`,
          }
        }
        const files = deps.listFiles()
        const index = Number(call.input.index)
        const file = Number.isInteger(index) ? files[index] : undefined
        if (!file) {
          return {
            output: `no knowledge-base file at index ${call.input.index}; pass a path from a search result instead`,
            isError: true,
            summary: `read #${call.input.index}`,
          }
        }
        const page = await deps.read(file.path, offset)
        const more = page.offset + page.text.length < page.totalChars
        return {
          output:
            `${file.folder}/${file.name} — characters ${page.offset}..${page.offset + page.text.length} of ${page.totalChars}` +
            (more ? ' (continue with offset)' : ' (end)') +
            `\n\n${page.text}`,
          summary: `${file.name}@${page.offset}`,
        }
      }
      return { output: `unknown tool: ${call.name}`, isError: true, summary: call.name }
    },
    verifyResponse: (response) => {
      const cited = [...response.matchAll(/\]\((filenav:\/\/[^)\s]+)\)/g)].map((m) => m[1]!)
      if (cited.length === 0) return null
      // hrefs arrive percent-encoded (the renderer's own generator encodes
      // spaces/CJK); compare decoded forms so a hand-written raw href passes too
      const known = new Set([
        ...deps.listFiles().map((f) => f.path),
        ...citableSearchPaths,
        ...(deps.alsoCitable?.() ?? []),
      ])
      const unknown = cited.filter((href) => {
        const raw = href.slice(FILE_NAV_SCHEME.length)
        let decoded = raw
        try {
          decoded = decodeURIComponent(raw)
        } catch {
          // keep the raw text
        }
        return !known.has(decoded)
      })
      if (unknown.length === 0) return null
      return (
        `These citations are not knowledge-base paths: ${unknown.join(', ')}. ` +
        'Cite only the cite links given verbatim in the knowledge-base list or a search result, or remove the citation.'
      )
    },
  }
}
