import { readFile, stat } from 'node:fs/promises'
import { extname } from 'node:path'
import { parseFileToText } from '@genoffice/file-parse'

/** larger files are indexed by name and folder only */
export const MAX_EXTRACT_BYTES = 64 * 1024 * 1024
/** body text kept for snippets and indexed; longer files are cut */
export const MAX_BODY_CHARS = 1_000_000

export type Extracted =
  { kind: 'text'; text: string } | { kind: 'name-only' } | { kind: 'error'; error: string }

const MARKUP_EXTS = new Set(['html', 'htm'])

function stripMarkup(html: string): string {
  return html
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
}

export async function extractText(path: string): Promise<Extracted> {
  try {
    const st = await stat(path)
    if (st.size > MAX_EXTRACT_BYTES) return { kind: 'name-only' }
    const ext = extname(path).slice(1).toLowerCase()
    if (MARKUP_EXTS.has(ext)) {
      return {
        kind: 'text',
        text: stripMarkup(await readFile(path, 'utf-8')).slice(0, MAX_BODY_CHARS),
      }
    }
    const parsed = await parseFileToText(path)
    if (parsed.kind === 'unsupported') return { kind: 'name-only' }
    if (!parsed.ok || parsed.kind !== 'text')
      return { kind: 'error', error: parsed.error ?? 'parse failed' }
    return { kind: 'text', text: (parsed.text ?? '').slice(0, MAX_BODY_CHARS) }
  } catch (e) {
    return { kind: 'error', error: e instanceof Error ? e.message : String(e) }
  }
}
