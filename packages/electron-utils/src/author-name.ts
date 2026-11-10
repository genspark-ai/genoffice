/// The name written on comments and tracked changes (Word: Options → General
/// → User name). Stored under `authorName` in userData/app-settings.json by
/// the shell's Settings → General; every editor main resolves through here so
/// a blank or unset value falls back to the OS account name the same way.
import { readFileSync } from 'node:fs'
import { userInfo } from 'node:os'
import { join } from 'node:path'

export const AUTHOR_NAME_KEY = 'authorName'
export const AUTHOR_NAME_MAX = 120
const FALLBACK = 'User'

/** trimmed and capped; '' when nothing usable */
export function sanitizeAuthorName(value: unknown): string {
  if (typeof value !== 'string') return ''
  return value.trim().replace(/\s+/g, ' ').slice(0, AUTHOR_NAME_MAX)
}

/** the configured name from app-settings.json, or '' when unset/unreadable */
export function readAuthorNameSetting(settingsPath: string): string {
  try {
    const raw: unknown = JSON.parse(readFileSync(settingsPath, 'utf8'))
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return ''
    return sanitizeAuthorName((raw as Record<string, unknown>)[AUTHOR_NAME_KEY])
  } catch {
    return ''
  }
}

/** OS account name; a generic label when the platform refuses to say */
export function systemAuthorName(): string {
  try {
    return sanitizeAuthorName(userInfo().username) || FALLBACK
  } catch {
    return FALLBACK
  }
}

export function resolveAuthorName(settingsPath: string): string {
  return readAuthorNameSetting(settingsPath) || systemAuthorName()
}

/** convenience for the Electron mains */
export function configuredAuthorName(app: { getPath(name: 'userData'): string }): string {
  return resolveAuthorName(join(app.getPath('userData'), 'app-settings.json'))
}
