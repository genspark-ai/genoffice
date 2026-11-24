import { dirname, isAbsolute, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

/**
 * The local file a link's source names: a path as Word writes it in a field,
 * or the `file:` target of a linked picture. One that is relative is read
 * against the folder of the document that holds the link. Null for anything
 * that is not a local file: a web address, or a file on another machine.
 */
export function linkSourcePath(source: string, docPath: string | null): string | null {
  const raw = source.trim()
  if (!raw || isRemoteFile(raw)) return null
  if (/^file:/i.test(raw)) {
    try {
      return fileURLToPath(raw.replace(/\\/g, '/'))
    } catch {
      return null
    }
  }
  if (/^[a-z][a-z0-9+.-]+:/i.test(raw) && !/^[a-z]:[\\/]/i.test(raw)) return null
  if (isAbsolute(raw) || /^[a-z]:[\\/]/i.test(raw)) return raw
  return docPath ? resolve(dirname(docPath), raw) : null
}

/** a UNC path, or a `file:` URL that names a host: reading one sends the user's credentials to that host */
function isRemoteFile(raw: string): boolean {
  const url = /^file:(.*)$/i.exec(raw.replace(/\\/g, '/'))
  if (!url) return /^[\\/]{2}/.test(raw)
  const slashes = /^\/*/.exec(url[1])![0].length
  if (slashes !== 2) return slashes > 3
  const host = url[1].slice(2).split('/')[0]
  return host !== '' && host.toLowerCase() !== 'localhost'
}
