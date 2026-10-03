import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { resolve } from 'node:path'
import { runInNewContext } from 'node:vm'
import { describe, expect, it } from 'vitest'

const require = createRequire(import.meta.url)
const shellRoot = resolve(import.meta.dirname, '..')
const shellPackage = JSON.parse(readFileSync(resolve(shellRoot, 'package.json'), 'utf-8')) as {
  version: string
}

const UPDATE_URL = 'https://updates.example.com/genoffice'
const PUBLISHER_NAME = 'Example Software, Inc.'
const WIN_SIGN_SCRIPT_REL = 'scripts/win-sign.cjs'

interface PublishEntry {
  provider: string
  url?: string
  publisherName?: string | string[] | null
}

interface LoadedConfig {
  publish?: PublishEntry[]
  win?: { signtoolOptions?: { publisherName?: string | string[] | null } }
  beforePack: (context: { electronPlatformName: string }) => Promise<void>
}

/**
 * Loads the real packaging config with the given release env. The stand-in fs
 * reports every packaging input as present, except scripts/win-sign.cjs unless
 * `winSignScriptPresent` says otherwise, so each case exercises only the
 * verification gate.
 */
function loadConfig(opts: {
  env?: Record<string, string>
  winSignScriptPresent?: boolean
}): LoadedConfig {
  const configModule = { exports: {} as Record<string, unknown> }
  runInNewContext(readFileSync(resolve(shellRoot, 'electron-builder.cjs'), 'utf8'), {
    module: configModule,
    __dirname: shellRoot,
    process: {
      platform: 'linux',
      arch: 'x64',
      env: opts.env ?? {},
      execPath: '/usr/bin/node',
    },
    require: (id: string) => {
      if (id === './package.json') return shellPackage
      if (id === 'node:child_process') {
        return { execFileSync: () => Buffer.from('') }
      }
      if (id === 'node:fs') {
        return {
          ...require(id),
          existsSync: (path: string) => {
            if (String(path).endsWith(WIN_SIGN_SCRIPT_REL)) {
              return opts.winSignScriptPresent === true
            }
            return true
          },
          readFileSync: (path: string, encoding: string) => {
            // Keep the beforePack preconditions green: a CLI bundle that already
            // carries the app version, and a valid generated third-party notice.
            if (String(path).endsWith('genoffice.cjs')) {
              return `const __cliAppVersion = ${JSON.stringify(shellPackage.version)};\n`
            }
            if (String(path).endsWith('THIRD-PARTY-NOTICES.txt')) {
              return '@embedpdf/pdfium\nCopyright 2014 PDFium Authors\nApache License\n'
            }
            return require('node:fs').readFileSync(path, encoding as BufferEncoding)
          },
        }
      }
      return require(id)
    },
  })
  return configModule.exports as LoadedConfig
}

describe('Windows update signature verification', () => {
  it('refuses a release build whose update feed declares no publisher name', async () => {
    const config = loadConfig({ env: { GENOFFICE_UPDATE_URL: UPDATE_URL } })
    await expect(config.beforePack({ electronPlatformName: 'win32' })).rejects.toThrow(
      /GENOFFICE_WIN_PUBLISHER_NAME/,
    )
  })

  it('bakes the publisher name into the publish config electron-updater reads', () => {
    const config = loadConfig({
      env: { GENOFFICE_UPDATE_URL: UPDATE_URL, GENOFFICE_WIN_PUBLISHER_NAME: PUBLISHER_NAME },
    })
    expect(config.publish?.[0]).toMatchObject({
      provider: 'generic',
      url: UPDATE_URL,
      // Array form is what electron-builder's publish schema accepts; a bare
      // string makes the whole configuration fail its schema check.
      publisherName: [PUBLISHER_NAME],
    })
  })

  it('packages a Windows release that declares the publisher name', async () => {
    const config = loadConfig({
      env: { GENOFFICE_UPDATE_URL: UPDATE_URL, GENOFFICE_WIN_PUBLISHER_NAME: PUBLISHER_NAME },
    })
    await expect(config.beforePack({ electronPlatformName: 'win32' })).resolves.toBeUndefined()
  })

  it('leaves a local build with no update feed packable', async () => {
    const config = loadConfig({})
    expect(config.publish).toBeUndefined()
    await expect(config.beforePack({ electronPlatformName: 'win32' })).resolves.toBeUndefined()
  })

  it('refuses to sign a release without a publisher name to verify against', () => {
    expect(() =>
      loadConfig({
        env: { GENOFFICE_UPDATE_URL: UPDATE_URL, GENOFFICE_WIN_SIGN_MODE: 'production' },
        winSignScriptPresent: true,
      }),
    ).toThrow(/GENOFFICE_WIN_PUBLISHER_NAME/)
  })

  it('names the missing signing script instead of failing during signing', () => {
    expect(() =>
      loadConfig({
        env: {
          GENOFFICE_UPDATE_URL: UPDATE_URL,
          GENOFFICE_WIN_SIGN_MODE: 'production',
          GENOFFICE_WIN_PUBLISHER_NAME: PUBLISHER_NAME,
        },
      }),
    ).toThrow(new RegExp(WIN_SIGN_SCRIPT_REL.replace('.', '\\.')))
  })

  it('passes the publisher name to the sign hook so both agree', () => {
    const config = loadConfig({
      env: {
        GENOFFICE_UPDATE_URL: UPDATE_URL,
        GENOFFICE_WIN_SIGN_MODE: 'production',
        GENOFFICE_WIN_PUBLISHER_NAME: PUBLISHER_NAME,
      },
      winSignScriptPresent: true,
    })
    expect(config.win?.signtoolOptions?.publisherName).toBe(PUBLISHER_NAME)
    expect(config.publish?.[0].publisherName).toEqual([PUBLISHER_NAME])
  })
})
