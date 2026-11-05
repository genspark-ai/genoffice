/// A picked file's `file:///` spelling for an external link's rels Target:
/// forward slashes, percent-encoded segments — what the sidecar's path resolver
/// decodes back (Excel stores the same shape).
///
/// A Windows drive letter keeps its colon. Excel writes
/// `file:///C:/Data/src.xlsx`, and the sidecar's resolver checks for the drive
/// letter before percent-decoding, so `C%3A` would be read as a POSIX path and
/// the link would show as broken right after Change Source.
export function externalFileTarget(path: string): string {
  const universal = path.replaceAll('\\', '/')
  // A Windows path starts with its drive; every other shape starts with the
  // slash `file:///` already supplies.
  const drive = /^[A-Za-z]:/.exec(universal)?.[0]
  const rest = drive === undefined ? universal : universal.slice(drive.length)
  const trimmed = rest.startsWith('/') ? rest.slice(1) : rest
  const encoded = trimmed.split('/').map(encodeURIComponent).join('/')
  return drive === undefined ? `file:///${encoded}` : `file:///${drive}/${encoded}`
}
