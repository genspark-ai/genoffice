import { describe, expect, it } from 'vitest'
import { mkdtemp, mkdir, rm, utimes, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import type { AgentToolDef } from '@genoffice/agent-core'
import {
  buildCodexAppServerPrompt,
  codexAppServerLaunchArgs,
  codexAppServerOutputSchema,
  parseCodexAppServerTurn,
  resolveCodexCliPath,
} from '../src/codex-app-server'

const tools: AgentToolDef[] = [
  {
    name: 'replace_text',
    description: 'Replace document text',
    inputSchema: {
      type: 'object',
      properties: { before: { type: 'string' }, after: { type: 'string' } },
      required: ['before', 'after'],
    },
  },
]

describe('Codex app-server bridge', () => {
  it('automatically selects the newest complete desktop-managed Codex install', async () => {
    const root = await mkdtemp(join(tmpdir(), 'genoffice-codex-discovery-'))
    try {
      const older = join(root, 'old-release')
      const newer = join(root, 'new-release')
      await mkdir(older)
      await mkdir(newer)
      for (const directory of [older, newer]) {
        await writeFile(join(directory, 'codex.exe'), '')
        await writeFile(join(directory, 'codex-code-mode-host.exe'), '')
      }
      await utimes(join(older, 'codex.exe'), new Date(1_000), new Date(1_000))
      await utimes(join(newer, 'codex.exe'), new Date(2_000), new Date(2_000))
      await utimes(older, new Date(1_000), new Date(1_000))
      await utimes(newer, new Date(2_000), new Date(2_000))

      await expect(
        resolveCodexCliPath('', { platform: 'win32', managedInstallRoot: root, env: {} }),
      ).resolves.toBe(join(newer, 'codex.exe'))
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  it('migrates a stale saved hash path to the current managed install', async () => {
    const root = await mkdtemp(join(tmpdir(), 'genoffice-codex-migration-'))
    try {
      const current = join(root, 'current-release')
      await mkdir(current)
      await writeFile(join(current, 'codex.exe'), '')
      await writeFile(join(current, 'codex-code-mode-host.exe'), '')

      await expect(
        resolveCodexCliPath(join(root, 'removed-release', 'codex.exe'), {
          platform: 'win32',
          managedInstallRoot: root,
          env: {},
        }),
      ).resolves.toBe(join(current, 'codex.exe'))
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  it('starts the official stdio app-server instead of assuming an HTTP endpoint', () => {
    expect(codexAppServerLaunchArgs()).toEqual(['app-server', '--listen', 'stdio://'])
  })

  it('builds a strict one-turn output schema with only known tool names', () => {
    const schema = codexAppServerOutputSchema(tools) as {
      properties: { toolCalls: { items: { properties: { name: { enum: string[] } } } } }
    }
    expect(schema.properties.toolCalls.items.properties.name.enum).toEqual(['replace_text'])
    expect(codexAppServerOutputSchema([])).toMatchObject({
      properties: { toolCalls: { maxItems: 0 } },
    })
  })

  it('serializes new conversation events, tool results and schemas', () => {
    const prompt = buildCodexAppServerPrompt(
      'Edit only when requested.',
      [
        { role: 'user', text: 'Change A to B.' },
        {
          role: 'assistant',
          text: '',
          toolCalls: [{ id: 'call-1', name: 'replace_text', input: { before: 'A', after: 'B' } }],
        },
        { role: 'tool', results: [{ id: 'call-1', name: 'replace_text', output: 'done' }] },
      ],
      tools,
      8192,
    )
    expect(prompt).toContain('Edit only when requested.')
    expect(prompt).toContain('replace_text')
    expect(prompt).toContain('"output":"done"')
    expect(prompt).toContain('8192 output tokens')
  })

  it('parses text and JSON-encoded GenOffice tool arguments', () => {
    expect(
      parseCodexAppServerTurn(
        JSON.stringify({
          text: 'I will make that change.',
          toolCalls: [
            {
              id: 'call-1',
              name: 'replace_text',
              inputJson: JSON.stringify({ before: 'A', after: 'B' }),
            },
          ],
        }),
        tools,
      ),
    ).toEqual({
      text: 'I will make that change.',
      toolCalls: [{ id: 'call-1', name: 'replace_text', input: { before: 'A', after: 'B' } }],
    })
  })

  it('rejects unknown tools and empty final turns', () => {
    expect(() =>
      parseCodexAppServerTurn(
        JSON.stringify({ text: '', toolCalls: [{ id: 'x', name: 'shell', inputJson: '{}' }] }),
        tools,
      ),
    ).toThrow('unknown tool')
    expect(() =>
      parseCodexAppServerTurn(JSON.stringify({ text: '', toolCalls: [] }), tools),
    ).toThrow('no content')
  })
})
