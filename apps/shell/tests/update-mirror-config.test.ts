import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { resolve } from 'node:path'
import { runInNewContext } from 'node:vm'
import { describe, expect, it } from 'vitest'

const require = createRequire(import.meta.url)
const shellRoot = resolve(import.meta.dirname, '..')

interface LoadedConfig {
  publish?: unknown
  extraMetadata?: {
    version?: string
    genofficeUpdateMirror?: { baseUrl: string }
  }
}

/**
 * Loads the real electron-builder config in a vm sandbox with the given env,
 * exactly like the other packaging tests (file-associations, packaged-cli-version).
 */
function loadConfig(env: Record<string, string>): LoadedConfig {
  const configModule = { exports: {} as Record<string, unknown> }
  runInNewContext(readFileSync(resolve(shellRoot, 'electron-builder.cjs'), 'utf8'), {
    module: configModule,
    __dirname: shellRoot,
    URL,
    process: { platform: 'linux', arch: 'x64', env: { ...env }, execPath: '/usr/bin/node' },
    require: (id: string) => {
      if (id === 'electron/package.json') return { version: '0.0.0-test' }
      return require(id)
    },
  })
  return configModule.exports as LoadedConfig
}

describe('update mirror packaging config', () => {
  it('bakes GENOFFICE_UPDATE_MIRROR_URL into extraMetadata for the runtime failover', () => {
    const config = loadConfig({
      GENOFFICE_UPDATE_URL: 'https://cdn.example.com/mac',
      GENOFFICE_UPDATE_MIRROR_URL: 'https://mirror.example.com/releases',
    })
    expect(config.publish).toEqual([
      { provider: 'generic', url: 'https://cdn.example.com/mac', channel: 'latest' },
    ])
    expect(config.extraMetadata?.genofficeUpdateMirror).toEqual({
      baseUrl: 'https://mirror.example.com/releases',
    })
  })

  it('leaves no mirror metadata when the variable is unset', () => {
    const config = loadConfig({ GENOFFICE_UPDATE_URL: 'https://cdn.example.com/mac' })
    expect(config.extraMetadata).toBeUndefined()
  })

  it('rejects a non-HTTPS mirror URL', () => {
    expect(() =>
      loadConfig({
        GENOFFICE_UPDATE_URL: 'https://cdn.example.com/mac',
        GENOFFICE_UPDATE_MIRROR_URL: 'http://mirror.example.com/releases',
      }),
    ).toThrow(/GENOFFICE_UPDATE_MIRROR_URL must be an HTTPS base URL/)
  })
})
