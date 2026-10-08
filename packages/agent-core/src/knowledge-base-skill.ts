import type { AgentToolDef } from './types'
import type { AgentSkill } from './skill'

/**
 * The reader's starred files (the star corpus) as a searchable, readable body.
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
 * block numbers are rejected elsewhere.
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
  /** the corpus as of this turn (a synchronous snapshot; the panel refreshes it on the starred-list IPC) */
  listFiles(): KbFileInfo[]
  search(query: string, limit: number): Promise<KbSearchHit[]>
  /** one page of parsed text; path must be a corpus file */
  read(path: string, offset: number): Promise<KbReadPage>
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
The user's starred files ("收藏") form a knowledge base; each turn's context
lists them (index | folder | file | type | size, with the absolute path).
- When the request touches the knowledge base, search it first (search_knowledge_base)
  and read the promising files (read_knowledge_file, paged by offset) before
  answering. Never answer corpus questions from file names alone.
- A full-text search matches literal words only: if it returns nothing, retry
  with synonyms or the other language (e.g. 实习期 → 试用期 → probation) before
  concluding the corpus does not cover the topic.
- Long files are paged: the result reports totalChars and the next offset; keep
  reading with offset while it matters.
- Cite what you used: right after a statement drawn from a file, write
  [file name](filenav:///absolute/path) — copy the path exactly from the list
  or a search result. Cite ONLY paths that appear there; never invent one.
- Say so plainly when the knowledge base does not contain the answer.`

export function createKnowledgeBaseSkill(deps: KnowledgeBaseDeps): AgentSkill {
  const readTool: AgentToolDef = {
    name: 'read_knowledge_file',
    description:
      'Read a page of a knowledge-base file’s parsed text (docx/pdf/xlsx/pptx/text are parsed locally). ' +
      'Read offset=0 first; continue with the returned offset while totalChars says more remains.',
    inputSchema: {
      type: 'object',
      properties: {
        index: {
          type: 'integer',
          description: 'file index from the knowledge-base list (0-based)',
        },
        offset: { type: 'integer', description: 'character offset to read from; default 0' },
      },
      required: ['index'],
    },
  }
  return {
    id: 'knowledge-base',
    systemPrompt: KB_SYSTEM_PROMPT,
    tools: [
      readTool,
      {
        name: 'search_knowledge_base',
        description:
          'Full-text search over the knowledge base (the user’s starred files). ' +
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
      },
    ],
    buildContext: () => {
      const files = deps.listFiles()
      if (files.length === 0) return ''
      const plain = files.map(
        (f, i) => `${i} | ${f.folder} | ${f.name} | .${f.ext} | ${formatKbSize(f.sizeBytes)}`,
      )
      const paths = files.map((f, i) => `${i} = ${f.path}`)
      const withPaths = `${plain.join('\n')}\n\nAbsolute paths (for citations):\n${paths.join('\n')}`
      // a large corpus degrades to names only: the model searches first anyway,
      // and search results carry the real paths to cite
      const body =
        withPaths.length <= KB_CONTEXT_MAX_CHARS
          ? withPaths
          : `${plain.join('\n')}\n\n(The corpus is large; paths come with search results.)`
      return `Knowledge base — the user's starred files (index | folder | file | type | size):\n${body}`
    },
    executeTool: async (call) => {
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
        const lines = hits.map(
          (h) => `${h.folder}/${h.name}\n  path: ${h.path}\n  snippet: ${h.snippet}`,
        )
        return { output: lines.join('\n'), summary: `${query} → ${hits.length}` }
      }
      if (call.name === 'read_knowledge_file') {
        const files = deps.listFiles()
        const index = Number(call.input.index)
        const file = Number.isInteger(index) ? files[index] : undefined
        if (!file) {
          return {
            output: `no knowledge-base file at index ${call.input.index}`,
            isError: true,
            summary: `read #${call.input.index}`,
          }
        }
        const offsetRaw = Number(call.input.offset)
        const page = await deps.read(file.path, Number.isFinite(offsetRaw) ? offsetRaw : 0)
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
      const known = new Set(deps.listFiles().map((f) => f.path))
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
        'Cite only [file name](filenav:///absolute/path) links whose path appears verbatim in the knowledge-base list or a search result, or remove the citation.'
      )
    },
  }
}
