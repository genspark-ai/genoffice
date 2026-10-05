/**
 * File name of a native path, for both processes.
 *
 * The renderer runs sandboxed (no nodeIntegration), so it cannot use
 * `basename` from node:path — and node's own basename only strips the
 * separator of the host platform. Splitting on '/' alone returns the whole
 * path on Windows, where a session path looks like `C:\Users\me\deck.pptx`;
 * that string then becomes an export file name, and `writeFile` rejects it.
 */
export function baseName(filePath: string): string {
  const parts = filePath.split(/[/\\]/)
  return parts[parts.length - 1] ?? ''
}
