import { existsSync } from 'node:fs'
import { join } from 'node:path'

/**
 * Where a merged document is written.
 *
 * The renderer picks the name (it comes from the record), so the main process
 * keeps only its last part: a name must not carry a folder into the directory
 * the user chose, and two records of the same name must not overwrite each
 * other. genoffice#2097.
 */

/** the name a merged document is written under, as a .docx file name */
export function mergeOutputFileName(name: string): string {
  const base = name
    .split(/[\\/]/)
    .pop()!
    .replace(/[\\/:*?"<>|]/g, '_')
    .trim()
  const safe = base === '' || base === '.' || base === '..' ? 'document' : base
  return safe.toLowerCase().endsWith('.docx') ? safe : `${safe}.docx`
}

/** a free path inside dir for that name: name.docx, name-2.docx, name-3.docx… */
export function mergeOutputPath(dir: string, name: string): string {
  const fileName = mergeOutputFileName(name)
  const dot = fileName.lastIndexOf('.')
  const base = dot > 0 ? fileName.slice(0, dot) : fileName
  const ext = dot > 0 ? fileName.slice(dot) : ''
  let candidate = join(dir, fileName)
  for (let i = 2; existsSync(candidate); i += 1) {
    candidate = join(dir, `${base}-${i}${ext}`)
  }
  return candidate
}
