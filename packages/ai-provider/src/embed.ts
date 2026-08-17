/**
 * Local embeddings + retrieval primitives for Workspace Q&A.
 * Uses Ollama's native /api/embed endpoint (not the OpenAI-compatible
 * /v1 route) so nothing ever leaves the machine.
 */

/** strip a /v1 suffix and trailing slashes so the native API root is used */
export function ollamaApiRoot(baseUrl?: string): string {
  return (baseUrl ?? 'http://localhost:11434')
    .replace(/\/v1\/?$/, '')
    .replace(/\/+$/, '')
}

export interface EmbedRequest {
  model: string
  input: string[] | string
}

/** batched embedding via POST {root}/api/embed */
export async function embedWithOllama(
  baseUrl: string | undefined,
  model: string,
  texts: string[],
): Promise<number[][]> {
  const root = ollamaApiRoot(baseUrl)
  const resp = await fetch(`${root}/api/embed`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ model, input: texts } satisfies EmbedRequest),
    signal: AbortSignal.timeout(120_000),
  })
  if (!resp.ok) {
    throw new Error(`embedding request failed: HTTP ${resp.status}`)
  }
  const data = (await resp.json()) as { embeddings?: number[][] }
  if (!Array.isArray(data.embeddings) || data.embeddings.length !== texts.length) {
    throw new Error('embedding response missing embeddings')
  }
  return data.embeddings
}

/** heuristically prefer a model that can embed; null when none is installed */
export function pickEmbeddingModel(models: readonly string[]): string | null {
  const candidate = models.find((m) =>
    /embed|nomic|bge|mxbai|snowflake|gte|e5|all-minilm|minilm/i.test(m),
  )
  return candidate ?? null
}

/**
 * Split text into overlapping chunks on paragraph/sentence boundaries.
 * Keeps ~`size` chars per chunk with `overlap` chars of context carried over.
 */
export function chunkText(
  text: string,
  size = 800,
  overlap = 120,
): string[] {
  const normalized = text.replace(/\r\n/g, '\n').trim()
  if (normalized.length === 0) return []
  const chunks: string[] = []
  let start = 0
  while (start < normalized.length) {
    let end = Math.min(start + size, normalized.length)
    if (end < normalized.length) {
      // prefer a paragraph / sentence boundary near the cut
      const window = normalized.slice(start, end)
      const breakAt = Math.max(
        window.lastIndexOf('\n\n'),
        window.lastIndexOf('. '),
        window.lastIndexOf('。'),
        window.lastIndexOf('! '),
        window.lastIndexOf('? '),
      )
      if (breakAt > size / 2) end = start + breakAt + 1
    }
    chunks.push(normalized.slice(start, end).trim())
    if (end >= normalized.length) break
    start = Math.max(end - Math.min(overlap, end - start), start + 1)
  }
  return chunks.filter((c) => c.length > 0)
}

export function cosineSimilarity(a: readonly number[], b: readonly number[]): number {
  if (a.length !== b.length || a.length === 0) return 0
  let dot = 0
  let na = 0
  let nb = 0
  for (let i = 0; i < a.length; i++) {
    const av = a[i]
    const bv = b[i]
    if (av === undefined || bv === undefined) continue
    dot += av * bv
    na += av * av
    nb += bv * bv
  }
  if (na === 0 || nb === 0) return 0
  return dot / (Math.sqrt(na) * Math.sqrt(nb))
}

export interface WorkspaceIndexHit {
  file: string
  snippet: string
  score: number
}

export interface WorkspaceIndexResult {
  ok: boolean
  files?: number
  chunks?: number
  model?: string
  error?: string
}

export interface WorkspaceSearchResult {
  ok: boolean
  query?: string
  results?: WorkspaceIndexHit[]
  error?: string
}

/** top-k chunks by cosine similarity against the query vector */
export function searchChunks(
  chunks: ReadonlyArray<{ file: string; text: string; vector: readonly number[] }>,
  queryVector: readonly number[],
  k: number,
): WorkspaceIndexHit[] {
  return chunks
    .map((c) => ({ file: c.file, snippet: c.text, score: cosineSimilarity(queryVector, c.vector) }))
    .sort((a, b) => b.score - a.score)
    .slice(0, k)
    .filter((h) => h.score > 0)
}
