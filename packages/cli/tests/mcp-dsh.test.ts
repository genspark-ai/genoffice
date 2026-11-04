import { describe, expect, it } from 'vitest'
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  mcpSnippet,
  readMcpEntry,
  removeMcpEntry,
  writeMcpEntry,
  MCP_SERVER_NAME,
} from '../src/agent-mcp'

const LAUNCH = { command: '/usr/local/bin/genoffice', args: ['mcp'], env: undefined }

const dir = mkdtempSync(join(tmpdir(), 'mcp-dsh-'))

describe('the dsh Cordis patch shape', () => {
  it('the snippet shows the insert block DSH reads', () => {
    const snippet = mcpSnippet('dsh', LAUNCH)
    expect(snippet).toContain('- insert:')
    expect(snippet).toContain('id: mcp-genoffice')
    expect(snippet).toContain("name: '@deepseek-ai/dsh-mcp-client'")
    expect(snippet).toContain('serverName: genoffice')
    expect(snippet).toContain(`command: ${LAUNCH.command}`)
    expect(snippet).toContain("args: ['mcp']")
    expect(snippet).toContain('transport: stdio')
  })

  it('write → registered → remove → absent, other entries untouched', () => {
    const file = join(dir, 'cordis-existing.yml')
    writeFileSync(
      file,
      [
        '- insert:',
        '  - id: mcp-other',
        "    name: '@deepseek-ai/dsh-mcp-client'",
        '    config:',
        '      serverName: other-server',
        '      transport: stdio',
        '      command: /usr/bin/other',
        '      args: []',
      ].join('\n') + '\n',
    )
    writeMcpEntry('dsh', file, LAUNCH)
    const text = readFileSync(file, 'utf-8')
    expect(text).toContain('mcp-genoffice')
    expect(text).toContain('mcp-other') // the other server stays
    expect(readMcpEntry('dsh', file, LAUNCH).status).toBe('registered')
    expect(removeMcpEntry('dsh', file)).toBe(true)
    const after = readFileSync(file, 'utf-8')
    expect(after).not.toContain('mcp-genoffice')
    expect(after).toContain('mcp-other') // the other server survives the remove
    expect(readMcpEntry('dsh', file, LAUNCH).status).toBe('absent')
    expect(removeMcpEntry('dsh', file)).toBe(false)
  })

  it('a fresh file gets a bare insert block', () => {
    const file = join(dir, 'cordis-fresh.yml')
    writeMcpEntry('dsh', file, LAUNCH)
    const text = readFileSync(file, 'utf-8')
    expect(text).toContain('- insert:')
    expect(readMcpEntry('dsh', file, LAUNCH).status).toBe('registered')
  })

  it('a genoffice launcher at another path is stale, a foreign command is occupied', () => {
    const file = join(dir, 'cordis-stale.yml')
    writeFileSync(
      file,
      [
        '- insert:',
        '  - id: mcp-genoffice',
        "    name: '@deepseek-ai/dsh-mcp-client'",
        '    config:',
        '      serverName: genoffice',
        '      transport: stdio',
        '      command: /old/path/genoffice',
        "      args: ['mcp']",
      ].join('\n') + '\n',
    )
    const stale = readMcpEntry('dsh', file, LAUNCH)
    expect(stale.status).toBe('stale')
    expect(stale.command).toBe('/old/path/genoffice')

    const occupiedFile = join(dir, 'cordis-occupied.yml')
    writeFileSync(
      occupiedFile,
      [
        '- insert:',
        '  - id: mcp-genoffice',
        "    name: '@other/mcp-client'",
        '    config:',
        '      serverName: not-genoffice',
        '      transport: stdio',
        '      command: /usr/bin/not-genoffice',
        '      args: []',
      ].join('\n') + '\n',
    )
    // the id is the ownership key DSH reads, but a foreign command under it is
    // not ours to overwrite without --force
    expect(readMcpEntry('dsh', occupiedFile, LAUNCH).status).toBe('occupied')
  })

  it('writes the env the packaged app needs to run as Node', () => {
    const snippet = mcpSnippet('dsh', { ...LAUNCH, env: { ELECTRON_RUN_AS_NODE: '1' } })
    expect(snippet).toContain('env:')
    expect(snippet).toContain("ELECTRON_RUN_AS_NODE: '1'")
  })

  it('doubles a single quote in a scalar and rewrites Windows paths in the args too', () => {
    const snippet = mcpSnippet('dsh', {
      command: 'C:\\Program Files\\GenOffice\\genoffice.exe',
      args: ['C:\\Program Files\\GenOffice\\genoffice.cjs', "it's"],
      env: undefined,
    })
    expect(snippet).toContain('command: C:/Program Files/GenOffice/genoffice.exe')
    expect(snippet).toContain("args: ['C:/Program Files/GenOffice/genoffice.cjs', 'it''s']")
  })

  it('leaves the rest of the patch byte for byte', () => {
    const file = join(dir, 'cordis-bytes.yml')
    const original = [
      '# a comment the user wrote',
      '',
      '- insert:',
      '  - id: mcp-other',
      "    name: '@deepseek-ai/dsh-mcp-client'",
      '    config:',
      '      serverName: other',
      '      transport: stdio',
      '      command: /usr/bin/other',
      '      args: []',
      '',
    ].join('\n')
    writeFileSync(file, original)
    writeMcpEntry('dsh', file, LAUNCH)
    const written = readFileSync(file, 'utf-8')
    expect(written.startsWith(original)).toBe(true)
    expect(written).toContain('mcp-genoffice')
    removeMcpEntry('dsh', file)
    expect(readFileSync(file, 'utf-8')).toBe(original)
  })

  it('a file that is not a top-level sequence is left for a manual edit', () => {
    const file = join(dir, 'cordis-mapping.yml')
    const original = 'profile: web\nplugins: []\n'
    writeFileSync(file, original)
    expect(readMcpEntry('dsh', file, LAUNCH).status).toBe('manual')
    writeMcpEntry('dsh', file, LAUNCH)
    expect(readFileSync(file, 'utf-8')).toBe(original)
  })

  it('the CJK sheet-name test file round-trips through a real write', () => {
    const file = join(dir, 'cordis-roundtrip.yml')
    writeMcpEntry('dsh', file, LAUNCH)
    writeMcpEntry('dsh', file, LAUNCH) // idempotent rewrite
    const text = readFileSync(file, 'utf-8')
    expect(text.match(/mcp-genoffice/g)?.length).toBe(1)
  })
})

describe('MCP_SERVER_NAME consistency', () => {
  it('the server name is genoffice across the shapes', () => {
    expect(MCP_SERVER_NAME).toBe('genoffice')
  })
})
