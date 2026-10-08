import { randomUUID } from 'node:crypto'
import { renameSync, unlinkSync, writeFileSync } from 'node:fs'

/**
 * userData/open-documents.json: the files the running shell has open in tabs,
 * for other processes — the genoffice CLI refuses to rewrite a document in place
 * while the editor shows it, since the editor would either warn at its next
 * save (docs, sheets) or silently overwrite the change (slides). Rewritten on
 * every tab change, emptied at startup, removed on quit; the pid lets a reader
 * ignore a registry left behind by a crash.
 */
export const OPEN_DOCUMENTS_FILE = 'open-documents.json'

export interface OpenDocumentsRegistry {
  pid: number
  updatedAt: string
  paths: string[]
}

const reportedWriteErrors = new Set<string>()

export function publishOpenDocuments(path: string, paths: readonly string[]): boolean {
  const registry: OpenDocumentsRegistry = {
    pid: process.pid,
    updatedAt: new Date().toISOString(),
    paths: [...new Set(paths)].sort(),
  }
  const tempPath = `${path}.${process.pid}.${randomUUID()}.tmp`
  try {
    writeFileSync(tempPath, JSON.stringify(registry), 'utf8')
    renameSync(tempPath, path)
    return true
  } catch (err) {
    try {
      unlinkSync(tempPath)
    } catch {}
    const message = err instanceof Error ? err.message : String(err)
    const key = (err as NodeJS.ErrnoException)?.code ?? message
    if (!reportedWriteErrors.has(key)) {
      reportedWriteErrors.add(key)
      console.warn('[shell] open-documents registry write failed:', message)
    }
    return false
  }
}

export function clearOpenDocuments(path: string): void {
  try {
    unlinkSync(path)
  } catch {}
}
