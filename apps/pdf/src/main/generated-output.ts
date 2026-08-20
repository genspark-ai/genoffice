import { existsSync } from 'node:fs'
import { join } from 'node:path'

/** Pick a safe, unused PDF path inside the configured GenOffice save directory. */
export function uniqueGeneratedPdfPath(
  dir: string,
  suggestedName: string,
  pathExists: (path: string) => boolean = existsSync,
): string {
  // Control characters are intentionally rejected from generated file names.
  // eslint-disable-next-line no-control-regex
  const invalidFileNameCharacters = /[/\\:*?"<>|\u0000-\u001f]/g
  // Take the last path segment on either separator so traversal never escapes the
  // directory, but keep ':' as an ordinary (sanitized) character: path.basename on
  // Windows would read 'a:b?.pdf' as a drive-relative path and drop the stem.
  let fileName = (String(suggestedName || 'merged.pdf').split(/[\\/]/).pop() ?? '')
    .replace(invalidFileNameCharacters, '_')
    .trim()
  if (!fileName || fileName === '.' || fileName === '..') fileName = 'merged.pdf'
  if (!/\.pdf$/i.test(fileName)) fileName += '.pdf'

  const stem = fileName.slice(0, -4)
  let candidate = join(dir, fileName)
  for (let i = 2; pathExists(candidate); i++) candidate = join(dir, `${stem}-${i}.pdf`)
  return candidate
}
