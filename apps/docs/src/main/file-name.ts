import { existsSync } from 'node:fs'
import { join } from 'node:path'

/**
 * The bare file name behind a caller-supplied suggestion.
 *
 * Renderers and MCP tools hand suggested names into the new-file writers
 * (docs:save-new's defaultName among them), and those writes have no dialog
 * or per-path grant behind them — the name is the only thing deciding where
 * the bytes land. Stripping every separator keeps a `..` segment, an absolute
 * path or a Windows-style drive path from steering the write out of the
 * destination directory; the unique-path loop then works purely inside it.
 */
/** what a write is called when the suggestion carried no name of its own */
export const UNNAMED_FILE = 'untitled'

export function bareFileName(fileName: string): string {
  const bare = fileName.split(/[\\/]+/).pop() ?? ''
  // a bare `.` or `..` segment is navigation, not a name
  return bare === '.' || bare === '..' ? '' : bare
}

/** first free path for fileName inside dir: name.ext, name-2.ext, name-3.ext…
 *  fileName arrives from renderers and MCP callers as a *suggestion*, so it is
 *  reduced to its bare name first (file-name.ts): a `..` segment or an
 *  absolute path must not steer the write out of dir — docs:save-new's
 *  suggested name had no dialog or grant behind it, and every new-file write
 *  funnels through here. */
export function uniquePathIn(dir: string, fileName: string): string {
  // a suggestion that names nothing (`report/`, `..`) would make the candidate the directory itself
  const bare = bareFileName(fileName) || UNNAMED_FILE
  const dot = bare.lastIndexOf('.')
  const base = dot > 0 ? bare.slice(0, dot) : bare
  const ext = dot > 0 ? bare.slice(dot) : ''
  let candidate = join(dir, bare)
  for (let i = 2; existsSync(candidate); i++) candidate = join(dir, `${base}-${i}${ext}`)
  return candidate
}
