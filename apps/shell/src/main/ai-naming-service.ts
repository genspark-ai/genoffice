import { app, ipcMain } from 'electron'
import { join } from 'node:path'
import { readFileSync } from 'node:fs'
import type { AgentMessage } from '@genoffice/agent-core'
import { streamForProvider } from '@genoffice/ai-provider'
import { gskApiKey } from '@genoffice/ai-search'
import { buildNamingInstruction, excerptForNaming, sanitizeFileStem } from './ai-naming'
import { resolveNamingProvider } from './ai-naming-provider'

/** The channel the renderer's "Name with AI…" button calls. */
export const AI_NAMING_SUGGEST_CHANNEL = 'ai:suggest-file-name'

/** A name is a few words, not an essay; anything longer is a model that ignored the brief. */
const MAX_TOKENS = 64

/**
 * The AI settings, read the same way `registerAiIpc` reads them: a plain
 * `ai-settings.json` in userData. The naming service is registered from the same
 * process, so it can read the file itself rather than round-tripping the value
 * through an IPC hop that already exists for renderers.
 */
function readAiSettingsFile(): unknown {
  try {
    return JSON.parse(
      readFileSync(join(app.getPath('userData'), 'ai-settings.json'), 'utf8'),
    ) as unknown
  } catch {
    // no file yet, or unreadable: a naming attempt will simply be declined
    return {}
  }
}

/**
 * Ask the model for a name, and return it as a bare file stem.
 *
 * This runs one cheap turn and takes the first line. If the provider is
 * unreachable the caller keeps whatever name the document already has — a
 * failed rename must never cost the reader their file.
 */
async function askModel(content: string, lang: string): Promise<string> {
  const target = resolveNamingProvider(readAiSettingsFile(), gskApiKey())
  if (!target) return ''

  let out = ''
  const messages: AgentMessage[] = [
    { role: 'user', text: buildNamingInstruction(excerptForNaming(content), lang) },
  ]
  const controller = new AbortController()
  // a naming call that has not answered in 30s is not going to
  const timer = setTimeout(() => controller.abort(), 30_000)
  try {
    await streamForProvider(target.provider, target.config, '', messages, [], MAX_TOKENS, {
      onDelta: (text) => {
        out += text
      },
      onToolCall: () => {},
      signal: controller.signal,
    })
  } catch (err) {
    console.warn('[shell] file naming failed:', err)
    return ''
  } finally {
    clearTimeout(timer)
  }
  // a model that wrapped the name in a sentence still gets one usable line
  return sanitizeFileStem(out.split('\n').find((line) => line.trim()) ?? '')
}

/**
 * The naming bridge's only channel.
 *
 * There is no preference and no automatic path: a name is proposed because the
 * reader asked for one, and the caller puts it in front of them.
 */
export function registerAiNamingIpc(lang: () => string): void {
  ipcMain.handle(
    AI_NAMING_SUGGEST_CHANNEL,
    async (
      _event,
      input: { content?: unknown },
    ): Promise<{ ok: boolean; name?: string; reason?: string }> => {
      const content = typeof input?.content === 'string' ? input.content : ''
      if (!content.trim()) return { ok: false, reason: 'empty-content' }
      const name = await askModel(content, lang())
      // A model that returned nothing usable — no provider, unreachable,
      // refused, or a stem that sanitised away to empty — is a *failed* naming,
      // not a successful one handing back "". The caller's contract is that it
      // keeps whatever name the document already had, and returning ok:true
      // with an empty name would instead rename the file to nothing.
      if (!name) return { ok: false, reason: 'empty-name' }
      return { ok: true, name }
    },
  )
}
