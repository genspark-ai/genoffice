/// Imported from its own electron-utils subpath rather than the package
/// barrel: the barrel drags Electron-only modules into headless bundles, and
/// tests that mock the barrel would otherwise have to mock this helper too.
export { atomicWriteFile } from '@genoffice/electron-utils/atomic-write'
