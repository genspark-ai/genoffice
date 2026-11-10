/// Re-exported from its own electron-utils subpath on purpose: the CLI bundles
/// this module (packages/cli/src/formats/xlsx.ts), and the package barrel pulls
/// in renderer/main-only modules that import 'electron' at top level, which a
/// plain Node process cannot resolve.
export { atomicWriteFile } from '@genoffice/electron-utils/atomic-write'
