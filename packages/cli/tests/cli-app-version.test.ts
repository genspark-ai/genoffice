import { describe, expect, it } from 'vitest'
import {
  CLI_BUNDLE,
  CLI_VERSION_BANNER_PREFIX,
  CLI_VERSION_ENV,
  cliVersionBanner,
  findElectronRequires,
  readBundledCliVersion,
  resolveCliVersion,
} from '../build.mjs'
import packageJson from '../package.json'

describe('packaged CLI version', () => {
  it('defaults to the CLI package version when no app version is injected', () => {
    expect(resolveCliVersion({}, packageJson.version)).toBe(packageJson.version)
    expect(resolveCliVersion({ [CLI_VERSION_ENV]: '' }, packageJson.version)).toBe(
      packageJson.version,
    )
    expect(resolveCliVersion({ [CLI_VERSION_ENV]: '   ' }, packageJson.version)).toBe(
      packageJson.version,
    )
  })

  it('reports the injected app version the release build baked in', () => {
    expect(resolveCliVersion({ [CLI_VERSION_ENV]: '0.5.149' }, packageJson.version)).toBe('0.5.149')
    expect(resolveCliVersion({ [CLI_VERSION_ENV]: ' 0.6.0 ' }, packageJson.version)).toBe('0.6.0')
  })

  it('round-trips the version through the bundle banner', () => {
    const bundle = [
      "const __cliImportMetaUrl = require('node:url').pathToFileURL(__filename).href;",
      cliVersionBanner('0.5.149'),
      'module.exports = {}',
    ].join('\n')
    expect(readBundledCliVersion(bundle)).toBe('0.5.149')
    expect(readBundledCliVersion(`var x = ${JSON.stringify(CLI_VERSION_BANNER_PREFIX)}`)).toBeNull()
    expect(readBundledCliVersion(`${CLI_VERSION_BANNER_PREFIX}not json;`)).toBeNull()
  })

  it('escapes nothing into the banner that would break the bundle', () => {
    expect(cliVersionBanner('1.0.0')).toBe('const __cliAppVersion = "1.0.0";')
  })

  it('builds into the bundle packaging ships', () => {
    expect(CLI_BUNDLE.endsWith('genoffice.cjs')).toBe(true)
  })
})

describe('electron-free CLI bundle', () => {
  it('names the bundled modules that require electron, once each', () => {
    const bundle = [
      '// ../../packages/electron-utils/src/remote-image.ts',
      'var init_remote_image = __esm({ "x"() { fetch2 = require("undici"); } });',
      '// ../../packages/electron-utils/src/drop-open.ts',
      'var init_drop_open = __esm({ "y"() { import_electron = require("electron"); } });',
      "    import_electron_again = require('electron');",
      '// node_modules/some-dep/index.js',
      'var z = require("electron");',
    ].join('\n')
    expect(findElectronRequires(bundle)).toEqual([
      '../../packages/electron-utils/src/drop-open.ts',
      'node_modules/some-dep/index.js',
    ])
  })

  it('accepts a bundle that never requires electron', () => {
    expect(findElectronRequires('// src/cli.ts\nvar e = "electron"; require("node:fs")')).toEqual(
      [],
    )
  })
})
