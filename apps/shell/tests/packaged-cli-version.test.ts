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
const CLI_BUNDLE_REL = '../../packages/cli/dist/genoffice.cjs'

// ELF e_machine values the linux packaging guard reads out of the sidecar.
const ELF_X86_64 = 0x3e
const ELF_AARCH64 = 0xb7

interface Rebuild {
  script: string
  appVersion: string | undefined
}

/** A 20-byte ELF header (e_ident + e_machine) for the given architecture. */
function elfHeader(machine: number): Buffer {
  const header = Buffer.alloc(20)
  header.write('\x7fELF', 0, 'latin1')
  header.writeUInt16LE(machine, 18)
  return header
}

/**
 * Loads the real packaging config and drives its beforePack hook with a stand-in
 * CLI bundle: `rebuiltWith` is the version the fake rebuild bakes in, so the hook
 * sees exactly what a real GENOFFICE_APP_VERSION rebuild would leave behind.
 */
function loadConfig(opts: {
  baked: string
  extraMetadataVersion?: string
  rebuildApplies?: boolean
  /** e_machine the fake linux sidecar reports (defaults to x86_64). */
  sidecarMachine?: number
  /** Extra env for the packaged config, e.g. GENOFFICE_LINUX_ARM64=1. */
  env?: Record<string, string>
  /** Whether the fake linux sidecar exists at all (false → missing source). */
  sidecarExists?: boolean
}): {
  config: {
    beforePack: (context: { electronPlatformName: string }) => Promise<void>
    linux?: { target?: Array<{ target: string; arch: string[] }> }
  }
  rebuilds: Rebuild[]
} {
  const rebuilds: Rebuild[] = []
  let baked = opts.baked
  const configModule = { exports: {} as Record<string, unknown> }
  runInNewContext(readFileSync(resolve(shellRoot, 'electron-builder.cjs'), 'utf8'), {
    module: configModule,
    __dirname: shellRoot,
    process: {
      platform: 'linux',
      arch: 'x64',
      env: { ...opts.env },
      execPath: '/usr/bin/node',
    },
    require: (id: string) => {
      if (id === './package.json') return shellPackage
      if (id === 'node:child_process') {
        return {
          execFileSync: (_file: string, args: string[], runOpts: { env?: NodeJS.ProcessEnv }) => {
            rebuilds.push({ script: args[0], appVersion: runOpts.env?.GENOFFICE_APP_VERSION })
            if (opts.rebuildApplies !== false) baked = runOpts.env?.GENOFFICE_APP_VERSION ?? baked
            return Buffer.from('')
          },
        }
      }
      if (id === 'node:fs') {
        return {
          ...require(id),
          existsSync: (path: string) =>
            String(path).endsWith('xlsx-sidecar') ? opts.sidecarExists !== false : true,
          readFileSync: (path: string, encoding: string) => {
            if (String(path).endsWith('genoffice.cjs')) {
              return `const __cliAppVersion = ${JSON.stringify(baked)};\n`
            }
            // Packaging also gates on a generated third-party notice. Serve a
            // valid one so the CLI version check is what these tests exercise.
            if (String(path).endsWith('THIRD-PARTY-NOTICES.txt')) {
              return '@embedpdf/pdfium\nCopyright 2014 PDFium Authors\nApache License\n'
            }
            return require('node:fs').readFileSync(path, encoding as BufferEncoding)
          },
          // The linux packaging guard reads the sidecar's ELF header to learn
          // its architecture: serve the header of the requested one.
          openSync: () => 42,
          readSync: (_fd: number, buffer: Buffer, offset: number, length: number) => {
            const header = elfHeader(opts.sidecarMachine ?? ELF_X86_64)
            const count = Math.min(length, header.length)
            header.copy(buffer, offset, 0, count)
            return count
          },
          closeSync: () => undefined,
        }
      }
      return require(id)
    },
  })
  const config = configModule.exports as {
    beforePack: (context: { electronPlatformName: string }) => Promise<void>
    extraMetadata?: { version?: string }
    linux?: { target?: Array<{ target: string; arch: string[] }> }
  }
  if (opts.extraMetadataVersion) {
    config.extraMetadata = { ...(config.extraMetadata ?? {}), version: opts.extraMetadataVersion }
  }
  return { config, rebuilds }
}

describe('packaged CLI version agreement', () => {
  it('rebuilds the CLI bundle with the app version when the two disagree', async () => {
    const { config, rebuilds } = loadConfig({ baked: '9.9.9' })
    await config.beforePack({ electronPlatformName: 'linux' })
    expect(rebuilds).toHaveLength(1)
    expect(rebuilds[0].script.replace(/\\/g, '/').endsWith('packages/cli/build.mjs')).toBe(true)
    expect(rebuilds[0].appVersion).toBe(shellPackage.version)
  })

  it('uses the release version electron-builder injects into extraMetadata', async () => {
    const { config, rebuilds } = loadConfig({
      baked: shellPackage.version,
      extraMetadataVersion: '0.5.149',
    })
    await config.beforePack({ electronPlatformName: 'linux' })
    expect(rebuilds).toHaveLength(1)
    expect(rebuilds[0].appVersion).toBe('0.5.149')
  })

  it('leaves an already matching bundle alone', async () => {
    const { config, rebuilds } = loadConfig({ baked: shellPackage.version })
    await config.beforePack({ electronPlatformName: 'linux' })
    expect(rebuilds).toEqual([])
  })

  it('fails the package when the bundle still disagrees after the rebuild', async () => {
    const { config, rebuilds } = loadConfig({ baked: '9.9.9', rebuildApplies: false })
    await expect(config.beforePack({ electronPlatformName: 'linux' })).rejects.toThrow(
      /but the app ships/,
    )
    expect(rebuilds).toHaveLength(1)
  })

  it('packages the CLI bundle whose version the check reads', () => {
    const { config } = loadConfig({ baked: shellPackage.version })
    const shipped = (config as unknown as { extraResources: Array<{ from: string }> })
      .extraResources
    expect(shipped.some((entry) => entry.from === CLI_BUNDLE_REL)).toBe(true)
  })
})

describe('linux sidecar architecture guard', () => {
  it('packages an x86_64 sidecar into the x64 linux targets', async () => {
    const { config } = loadConfig({ baked: shellPackage.version, sidecarMachine: ELF_X86_64 })
    await config.beforePack({ electronPlatformName: 'linux' })
    const targets = config.linux?.target ?? []
    expect(targets.length).toBeGreaterThan(0)
    expect(targets.every((entry) => entry.arch[0] === 'x64')).toBe(true)
  })

  it('packages an aarch64 sidecar into the arm64 linux targets', async () => {
    const { config } = loadConfig({
      baked: shellPackage.version,
      sidecarMachine: ELF_AARCH64,
      env: { GENOFFICE_LINUX_ARM64: '1' },
    })
    await config.beforePack({ electronPlatformName: 'linux' })
    const targets = config.linux?.target ?? []
    expect(targets.length).toBeGreaterThan(0)
    expect(targets.every((entry) => entry.arch[0] === 'arm64')).toBe(true)
  })

  it('refuses an aarch64 sidecar in an x64 package', async () => {
    const { config } = loadConfig({ baked: shellPackage.version, sidecarMachine: ELF_AARCH64 })
    await expect(config.beforePack({ electronPlatformName: 'linux' })).rejects.toThrow(
      /linux sidecar is aarch64 but this package ships x86_64/,
    )
  })

  it('refuses an x86_64 sidecar in an arm64 package', async () => {
    const { config } = loadConfig({
      baked: shellPackage.version,
      sidecarMachine: ELF_X86_64,
      env: { GENOFFICE_LINUX_ARM64: '1' },
    })
    await expect(config.beforePack({ electronPlatformName: 'linux' })).rejects.toThrow(
      /linux sidecar is x86_64 but this package ships aarch64/,
    )
  })

  it('names the x64 sidecar candidates when the x64 source is missing', async () => {
    const { config } = loadConfig({ baked: shellPackage.version, sidecarExists: false })
    const message = await config.beforePack({ electronPlatformName: 'linux' }).then(
      () => '',
      (err: unknown) => (err as Error).message,
    )
    expect(message).toContain('x86_64-unknown-linux-gnu')
    expect(message).not.toContain('aarch64-unknown-linux-gnu')
  })

  it('names the arm64 sidecar candidates when the arm64 source is missing', async () => {
    const { config } = loadConfig({
      baked: shellPackage.version,
      sidecarExists: false,
      env: { GENOFFICE_LINUX_ARM64: '1' },
    })
    const message = await config.beforePack({ electronPlatformName: 'linux' }).then(
      () => '',
      (err: unknown) => (err as Error).message,
    )
    expect(message).toContain('aarch64-unknown-linux-gnu')
    expect(message).not.toContain('x86_64-unknown-linux-gnu')
  })
})
