/// Imported from its own electron-utils subpath rather than the package
/// barrel: the barrel drags Electron-only modules into headless bundles, and
/// tests that mock the barrel would otherwise have to mock this helper too.
export { atomicWriteFile } from '@genoffice/electron-utils/atomic-write'

/** 'PK\x03\x04' local-file-header check — cheap docx/zip sanity test. */
export function looksLikeZip(bytes: Buffer): boolean {
  return bytes.length >= 4 && bytes.readUInt32LE(0) === 0x04034b50
}
