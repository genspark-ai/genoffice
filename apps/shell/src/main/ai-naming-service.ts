import { app, ipcMain } from 'electron'
import { join } from 'node:path'
import type { AgentMessage } from '@genoffice/agent-core'
import type { AiProviderConfig, AiProviderId } from '@genoffice/ai-provider'
import { streamForProvider } from '@genoffice/ai-provider'
import { readFileSync } from 'node:fs'
import { writeAppSetting, readAppSettings } from './app-settings'
import {
  buildNamingInstruction,
  excerptForNaming,
  fileNamingPrefOn,
  FILE_NAMING_PREF_KEY,
  sanitizeFileStem,
} from './ai-naming'
import { decideNaming, type NamingTrigger } from './naming-policy'

/** Channel the renderer's naming button and the save path both call. */
export const AI_NAMING_CHANNELS = {
  /** ask the model for a name for the given source text */
  suggest: 'ai:suggest-file-name',
  /** the model-naming preference (Settings → General) */
  getEnabled: 'ai:get-file-naming',
  setEnabled: 'ai:set-file-naming',
} as const

const PREF_PATH = () => join(app.getPath('userData'), 'app-settings.json')

/** The provider ids this build knows; anything else in the file is ignored. */
const PROVIDER_IDS: readonly AiProviderId[] = ['genspark', 'codex', 'anthropic']
/** A name is a few words, not an essay; anything longer is a model that ignored the brief. */
const MAX_TOKENS = 64

/**
 * A file name is a small thing to ask a model for, so this runs one cheap
 * turn and takes the first line. If the provider is unreachable the caller
 * keeps whatever name the document already has — a failed rename must never
 * cost the reader their file.
 */
/**
 * The AI settings, read the same way `registerAiIpc` reads them: a plain
 * `ai-settings.json` in userData. The naming service is registered from the same
 * process, so it can read the file itself rather than round-tripping the value
 * through an IPC hop that already exists for renderers.
 */
function readAiSettingsFile(): Record<string, unknown> {
  try {
    const raw: unknown = JSON.parse(
      readFileSync(join(app.getPath('userData'), 'ai-settings.json'), 'utf8'),
    )
    return raw && typeof raw === 'object' && !Array.isArray(raw)
      ? (raw as Record<string, unknown>)
      : {}
  } catch {
    // no file yet, or unreadable: a naming attempt will simply be declined
    return {}
  }
}

/** Whether a model could answer at all: a provider selected and given a key. */
function hasUsableProvider(): boolean {
  return currentProviderConfig().config.apiKey !== undefined
}

function currentProviderConfig(): { provider: AiProviderId; config: AiProviderConfig } {
  const settings = readAiSettingsFile()
  const selected = typeof settings.provider === 'string' ? settings.provider : 'genspark'
  // the file is user-editable and may name a provider this build does not
  // have, so a miss falls back rather than passing a bad id downstream
  const provider = (PROVIDER_IDS as readonly string[]).includes(selected)
    ? (selected as AiProviderId)
    : 'genspark'
  const providers = (settings.providers ?? {}) as Partial<Record<AiProviderId, AiProviderConfig>>
  return { provider, config: providers[provider] ?? ({} as AiProviderConfig) }
}

async function askModel(content: string, lang: string): Promise<string> {
  const { provider, config } = currentProviderConfig()

  let out = ''
  const messages: AgentMessage[] = [
    { role: 'user', text: buildNamingInstruction(excerptForNaming(content), lang) },
  ]
  const controller = new AbortController()
  // a naming call that has not answered in 30s is not going to
  const timer = setTimeout(() => controller.abort(), 30_000)
  try {
    await streamForProvider(provider, config, '', messages, [], MAX_TOKENS, {
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

export function registerAiNamingIpc(lang: () => string, enabled: () => boolean): void {
  ipcMain.handle(
    AI_NAMING_CHANNELS.suggest,
    async (_event, input: { content?: unknown; trigger?: unknown; filePath?: unknown }) => {
      const content = typeof input?.content === 'string' ? input.content : ''
      const trigger: NamingTrigger = input?.trigger === 'manual' ? 'manual' : 'first-save'
      const filePath = typeof input?.filePath === 'string' ? input.filePath : null
      const decision = decideNaming({
        enabled: enabled(),
        trigger,
        filePath,
        content,
        hasProvider: hasUsableProvider(),
      })
      if (decision !== 'name') return { ok: false, reason: decision }
      const name = await askModel(content, lang())
      // A model that returned nothing usable — no key, unreachable, refused, or
      // a stem that sanitised away to empty — is a *failed* naming, not a
      // successful one handing back "". The caller's contract is that it keeps
      // whatever name the document already had, and returning ok:true with an
      // empty name would instead rename the file to nothing.
      if (!name) return { ok: false, reason: 'empty-name' }
      return { ok: true, name }
    },
  )

  ipcMain.handle(AI_NAMING_CHANNELS.getEnabled, (): boolean => enabled())
  ipcMain.handle(AI_NAMING_CHANNELS.setEnabled, (_event, on: unknown): boolean => {
    if (typeof on !== 'boolean') return enabled()
    // the shell's own settings writer: an atomic replace of one key in the
    // same file every other preference lives in, so a reader that looks for
    // the wrong key here is a reader that always finds nothing
    try {
      writeAppSetting(PREF_PATH(), FILE_NAMING_PREF_KEY, on)
    } catch (err) {
      // the in-memory value still applies this session; the next launch just
      // asks again, which is better than refusing the toggle
      console.warn('[shell] could not persist the file-naming preference:', err)
    }
    return on
  })
}

/**
 * Whether first-save naming is switched on, as stored.
 *
 * Exported next to the writer so the two cannot read different keys — which is
 * not a type error here, only a preference that saves and then always comes
 * back off.
 */
export function readFileNamingPref(): boolean {
  return fileNamingPrefOn(readAppSettings(PREF_PATH()))
}
