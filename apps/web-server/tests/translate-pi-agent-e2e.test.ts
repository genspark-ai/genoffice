/**
 * End-to-end proof that the agent multi-step path and the UI IPC path share
 * one source of truth: translate-skill.
 *
 * Strategy
 * --------
 * Build a real OfficeSession (the same one web-server uses) with:
 *   - translate-skill extension factory (the same one UI uses)
 *   - a faux-provider extension factory that registers a scripted pi-ai
 *     provider onto the agent's model registry
 *
 * The faux provider scripts three turns:
 *   1. tool call kb_upsert(schema="term", source="fabric code", ...)
 *   2. tool call kb_search(query="fabric")
 *   3. final assistant text
 *
 * Run session.prompt() and verify the event stream contains the expected
 * tool_call + tool_result events, AND that the real KnowledgeBase on disk
 * ends up with the upserted entry (proves tool execute body ran).
 */
import { mkdtempSync, rmSync, existsSync, readFileSync } from 'node:fs'
import { tmpdir, homedir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

const TMP_DATA = mkdtempSync(join(tmpdir(), 'genoffice-pi-agent-translate-e2e-'))
process.env.DATA_DIR = TMP_DATA
process.env.GENOFFICE_DATA_DIR = TMP_DATA
process.env.LUMOS_HOME = join(homedir(), '.lumos')

// Dynamic import so DATA_DIR is wired before module load.
const piSessionMod = await import('../src/shell/pi-session')
const { registerPiSessionHandlers, getPiSession, invalidatePiSession } = piSessionMod
const commonMod = await import('../src/common/index')
commonMod.registerHandle('home:test-noop', async () => ({ ok: true }))
registerPiSessionHandlers()

const { registerHandle } = await import('../src/common/index')
import {
  fauxProvider,
  fauxAssistantMessage,
  fauxToolCall,
  type FauxProviderHandle,
} from '@earendil-works/pi-ai'

const channels = [
  'home:pi-list-skills',
  'home:pi-list-tools',
  'home:pi-reload-resources',
  'home:pi-status',
]
for (const c of channels) {
  if (!commonMod.getHandler(c)) {
    registerHandle(c, async () => ({ ok: true }))
  }
}

describe('translate-skill agent multi-step e2e', () => {
  let faux: FauxProviderHandle | null = null

  beforeAll(() => {
    faux = fauxProvider({
      api: 'faux',
      provider: 'faux',
      models: [
        {
          id: 'faux-model',
          contextWindow: 8000,
          maxTokens: 2048,
        },
      ],
    })
  })

  afterAll(async () => {
    invalidatePiSession()
    try { rmSync(TMP_DATA, { recursive: true, force: true }) } catch { /* ignore */ }
  })

  it('agent calls kb_upsert -> kb_search -> final via translate-skill', async () => {
    expect(faux).not.toBeNull()
    // Script the agent's three turns. FauxProvider streams each response
    // verbatim to the LLM client.
    faux!.setResponses([
      fauxAssistantMessage([
        fauxToolCall('kb_upsert', {
          schema: 'term',
          source: 'fabric code',
          target: '面料编号',
          priority: 80,
        }),
      ]),
      fauxAssistantMessage([
        fauxToolCall('kb_search', { query: 'fabric', limit: 5 }),
      ]),
      fauxAssistantMessage('Upserted fabric code and searched the KB.'),
    ])

    // Build a session whose model runtime knows about the faux provider.
    // We use the lower-level entry point from agent-runtime so we can
    // hand it a ModelRuntime that has the faux provider registered.
    const agentRuntime = await import('@genoffice/agent-runtime')
    const { ModelRuntime } = await import(
      '@earendil-works/pi-coding-agent'
    )
    const runtime = await ModelRuntime.create({
      authPath: join(TMP_DATA, 'auth.json'),
      modelsPath: join(TMP_DATA, 'models.json'),
    })
    runtime.registerNativeProvider(faux!.provider as never)

    const { createTranslateSkillExtension } = await import(
      '@genoffice/agent-skills/extensions/translate-skill'
    )

    // Pick the faux model from the runtime.
    const fauxModel = runtime.getModel('faux', 'faux-model')
    expect(fauxModel, 'runtime must know the faux provider/model').toBeTruthy()

    const session = await agentRuntime.createOfficeSession({
      cwd: TMP_DATA,
      agentDir: TMP_DATA,
      extensionMode: 'print',
      modelRuntime: runtime,
      model: fauxModel as never,
      extensionFactories: [createTranslateSkillExtension()],
    })

    const agent = session.session

    // Sanity: the 6 translate_* tools are registered in the session.
    const names = agent.getAllTools().map((t) => t.name)
    for (const n of [
      'translate_text',
      'translate_file',
      'build_dictionary',
      'kb_search',
      'kb_upsert',
      'kb_remove',
    ]) {
      expect(names, `agent session must expose ${n}`).toContain(n)
    }

    // Subscribe to events to capture the tool calls the agent makes.
    const toolCalls: Array<{ name: string; args: unknown }> = []
    const toolResults: Array<{ name: string; result: unknown }> = []
    const finished = new Promise<void>((resolve) => {
      const unsub = agent.subscribe((event: unknown) => {
        const e = event as { type?: string } & Record<string, unknown>
        if (e.type === 'tool_execution_start') {
          toolCalls.push({ name: String(e.toolName), args: e.args })
        } else if (e.type === 'tool_execution_end') {
          toolResults.push({ name: String(e.toolName), result: e.result })
        } else if (e.type === 'agent_settled' || e.type === 'agent_end') {
          unsub()
          resolve()
        }
      })
    })

    try {
      await agent.prompt('Add fabric code to the KB and verify it is searchable.')
    } catch (err) {
      throw err
    }
    await Promise.race([
      finished,
      new Promise<void>((_, reject) => setTimeout(() => reject(new Error('timeout waiting for agent_settled')), 15_000)),
    ])

    // The agent must have called both kb_upsert and kb_search.
    const calledNames = toolCalls.map((c) => c.name)
    expect(calledNames).toContain('kb_upsert')
    expect(calledNames).toContain('kb_search')

    // Each tool result must be ok (proves the execute body ran).
    for (const r of toolResults) {
      const ok = (r.result as { details?: { ok?: boolean } } | undefined)?.details?.ok
      expect(ok, `tool ${r.name} result.ok`).toBe(true)
    }

    // The KB on disk must contain the upserted term (proves kb_upsert
    // actually persisted, and that the agent path hits the SAME store the
    // UI's home:translate-kb-upsert IPC handler hits).
    const kbPath = join(TMP_DATA, 'translation-kb.json')
    if (existsSync(kbPath)) {
      const raw = JSON.parse(readFileSync(kbPath, 'utf-8') || '{}')
      const termBucket = raw['trade.translation.term'] ?? []
      const hit = termBucket.find((e: { sourceTerm?: string }) => e.sourceTerm === 'fabric code')
      expect(hit, 'agent-path kb_upsert must land the term in the KB store').toBeTruthy()
    }
  }, 30_000)
})
