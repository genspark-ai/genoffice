/**
 * Workspace Q&A (local-first RAG).
 *
 * Indexes the user's saved documents (the shared save folder, e.g.
 * <Documents>/GenOffice) into local Ollama embeddings and answers semantic
 * searches against that index. Everything runs in the main process and files
 * never leave the machine. Requires an embedding model installed in Ollama
 * (auto-detected from /api/tags; e.g. `nomic-embed-text`).
 */
import { readFileSync, readdirSync, statSync, writeFileSync, mkdirSync } from 'node:fs'
import { join, extname } from 'node:path'
import { ipcMain } from 'electron'
import {
  chunkText,
  defaultAiSettings,
  embedWithOllama,
  listOllamaModels,
  OLLAMA_DEFAULT_BASE_URL,
  ollamaApiRoot,
  pickEmbeddingModel,
  resolveAiSettings,
  searchChunks,
  type AiSettings,
  type LegacyAiSettings,
  type WorkspaceIndexResult,
  type WorkspaceSearchResult,
} from '@genoffice/ai-provider'
import { parseFileToText } from '@genoffice/file-parse'

export interface WorkspaceDeps {
  settingsPath: () => string
  userDataPath: (...parts: string[]) => string
  saveDir: () => string
}

const SUPPORTED_EXTS = new Set(['.md', '.txt', '.docx', '.pdf', '.pptx', '.xlsx'])
const MAX_FILE_BYTES = 10 * 1024 * 1024
const CHUNK_SIZE = 800
const CHUNK_OVERLAP = 120

interface IndexChunk {
  file: string
  text: string
  vector: number[]
}

interface StoredIndex {
  model: string
  root: string
  files: Record<string, { mtimeMs: number; size: number }>
  chunks: IndexChunk[]
}

function readJsonFile<T>(path: string, fallback: T): T {
  try {
    return JSON.parse(readFileSync(path, 'utf8')) as T
  } catch {
    return fallback
  }
}

function collectWorkspaceFiles(dir: string): Record<string, { mtimeMs: number; size: number }> {
  const found: Record<string, { mtimeMs: number; size: number }> = {}
  const visit = (current: string): void => {
    let entries
    try {
      entries = readdirSync(current, { withFileTypes: true })
    } catch {
      return
    }
    for (const entry of entries) {
      if (entry.name.startsWith('.')) continue
      const full = join(current, entry.name)
      if (entry.isDirectory()) {
        visit(full)
        continue
      }
      if (!entry.isFile() || !SUPPORTED_EXTS.has(extname(entry.name).toLowerCase())) continue
      try {
        const stat = statSync(full)
        if (stat.size > MAX_FILE_BYTES) continue
        found[full] = { mtimeMs: stat.mtimeMs, size: stat.size }
      } catch {
        // unreadable file — skip
      }
    }
  }
  visit(dir)
  return found
}

export function registerWorkspaceIpc(deps: WorkspaceDeps): void {
  const indexPath = () => join(deps.userDataPath('ai-workspace'), 'index.json')
  let indexCache: StoredIndex | null = null
  let building: Promise<StoredIndex> | null = null

  const readStoredIndex = (): StoredIndex | null =>
    indexCache ?? readJsonFile<StoredIndex | null>(indexPath(), null)
  const writeStoredIndex = (index: StoredIndex): void => {
    mkdirSync(join(deps.userDataPath('ai-workspace')), { recursive: true })
    writeFileSync(indexPath(), JSON.stringify(index))
    indexCache = index
  }

  /** probe /api/tags to find an installed embedding model; null when none exists */
  const resolveEmbedEndpoint = async (): Promise<{ model: string; root: string } | null> => {
    const stored = readJsonFile<Partial<AiSettings> & LegacyAiSettings>(deps.settingsPath(), {})
    const ai = resolveAiSettings(stored, defaultAiSettings())
    const baseUrl = ai.providers.ollama?.baseUrl || OLLAMA_DEFAULT_BASE_URL
    const root = ollamaApiRoot(baseUrl)
    const tags = await listOllamaModels(root)
    const model = pickEmbeddingModel((tags.models ?? []).map((m) => m.name))
    return model ? { model, root } : null
  }

  const buildIndex = async (): Promise<StoredIndex> => {
    const endpoint = await resolveEmbedEndpoint()
    if (!endpoint) {
      throw new Error('No embedding model installed in Ollama. Run: ollama pull nomic-embed-text')
    }
    const { model, root } = endpoint
    const dir = deps.saveDir()
    const files = collectWorkspaceFiles(dir)
    const prev = readStoredIndex()
    const filesMeta: StoredIndex['files'] = {}
    const chunks: IndexChunk[] = []
    const reusedPaths = new Set<string>()
    if (prev && prev.model === model && prev.root === root) {
      for (const [path, meta] of Object.entries(prev.files)) {
        const cur = files[path]
        if (cur && cur.mtimeMs === meta.mtimeMs && cur.size === meta.size) {
          reusedPaths.add(path)
          filesMeta[path] = meta
          for (const c of prev.chunks) if (c.file === path) chunks.push(c)
        }
      }
    }
    for (const [path, meta] of Object.entries(files)) {
      if (reusedPaths.has(path)) continue
      filesMeta[path] = meta
      try {
        const parsed = await parseFileToText(path)
        if (!parsed.ok || parsed.kind !== 'text' || parsed.text == null) continue
        const textChunks = chunkText(parsed.text, CHUNK_SIZE, CHUNK_OVERLAP)
        if (textChunks.length === 0) continue
        const vectors = await embedWithOllama(root, model, textChunks)
        textChunks.forEach((text, i) => {
          if (vectors[i]) chunks.push({ file: path, text, vector: vectors[i] })
        })
      } catch {
        // unreadable / unsupported file — skip, never fail the whole index
      }
    }
    const index: StoredIndex = { model, root, files: filesMeta, chunks }
    writeStoredIndex(index)
    return index
  }

  const ensureIndex = (): Promise<StoredIndex> => {
    const existing = readStoredIndex()
    if (existing && existing.chunks.length > 0) return Promise.resolve(existing)
    building ??= buildIndex().finally(() => {
      building = null
    })
    return building
  }

  ipcMain.handle('workspace:index', async (): Promise<WorkspaceIndexResult> => {
    try {
      const index = await ensureIndex()
      return {
        ok: true,
        files: Object.keys(index.files).length,
        chunks: index.chunks.length,
        model: index.model,
      }
    } catch (err) {
      return { ok: false, error: err instanceof Error ? err.message : String(err) }
    }
  })

  ipcMain.handle(
    'workspace:search',
    async (_event, query: unknown, k: unknown): Promise<WorkspaceSearchResult> => {
      try {
        const q = typeof query === 'string' ? query.trim() : ''
        if (!q) return { ok: false, error: 'empty query' }
        const index = await ensureIndex()
        if (index.chunks.length === 0) return { ok: true, query: q, results: [] }
        const limit = Math.min(Math.max(Number(k) || 5, 1), 10)
        const [qVec] = await embedWithOllama(index.root, index.model, [q])
        if (!qVec) return { ok: true, query: q, results: [] }
        return { ok: true, query: q, results: searchChunks(index.chunks, qVec, limit) }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )
}
