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
export function bareFileName(fileName: string): string {
  const bare = fileName.split(/[\\/]+/).pop() ?? ''
  // a bare `.` or `..` segment is navigation, not a name
  return bare === '.' || bare === '..' ? '' : bare
}
