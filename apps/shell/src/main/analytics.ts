import { randomUUID } from 'node:crypto'
import { readAppSettings, writeAppSetting } from './app-settings'

/**
 * Anonymous usage analytics via the GA4 Measurement Protocol.
 *
 * Official builds get the measurement id / api secret injected at package
 * time (electron-builder extraMetadata, fed by CI secrets — see
 * apps/shell/electron-builder.cjs). The credentials never live in the repo:
 * source builds and forks have no keys, createAnalytics() then returns a
 * no-op tracker and the app behaves identically minus the reporting.
 *
 * Privacy rules (also documented in README):
 * - events carry a random install id (client_id), never account identity
 * - no document content, file names, or paths are ever sent
 * - official builds report by default; the user can turn reporting off in Settings → General
 *   (ANALYTICS_ENABLED_KEY in userData/app-settings.json)
 * - sends are fire-and-forget with a short timeout; failures are swallowed
 */

/** app-settings.json key: false = explicit opt-out; absent means enabled */
export const ANALYTICS_ENABLED_KEY = 'analyticsEnabled'
/** app-settings.json key: random UUID identifying this install anonymously */
export const ANALYTICS_CLIENT_ID_KEY = 'analyticsClientId'

export interface AnalyticsKeys {
  measurementId: string
  apiSecret: string
}

export interface Analytics {
  /** true when packaged keys are present (events are still consent-gated) */
  readonly active: boolean
  /** fire-and-forget event; never throws, no-op without keys or after opt-out */
  track(name: string, params?: Record<string, string | number>): void
}

/**
 * Pull the injected GA4 credentials out of a parsed package.json. Returns
 * null unless both fields are non-empty strings — the caller then installs
 * the no-op tracker.
 */
export function extractAnalyticsKeys(pkg: unknown): AnalyticsKeys | null {
  if (!pkg || typeof pkg !== 'object') return null
  const raw = (pkg as Record<string, unknown>).genofficeAnalytics
  if (!raw || typeof raw !== 'object') return null
  const { measurementId, apiSecret } = raw as Record<string, unknown>
  if (typeof measurementId !== 'string' || !measurementId.trim()) return null
  if (typeof apiSecret !== 'string' || !apiSecret.trim()) return null
  return { measurementId: measurementId.trim(), apiSecret: apiSecret.trim() }
}

/** Source/dev runs stay keyless even if their local package metadata is edited. */
export function extractPackagedAnalyticsKeys(
  pkg: unknown,
  isPackaged: boolean,
): AnalyticsKeys | null {
  return isPackaged ? extractAnalyticsKeys(pkg) : null
}

/** Reporting is on by default; only an explicit persisted false disables it. */
export function analyticsEnabledFrom(settings: Record<string, unknown>): boolean {
  return settings[ANALYTICS_ENABLED_KEY] !== false
}

/**
 * Read the persistent anonymous install id, creating and saving one on first
 * use. Storage failures fall back to a per-session id rather than blocking.
 */
export function ensureAnalyticsClientId(settingsPath: string): string {
  const existing = readAppSettings(settingsPath)[ANALYTICS_CLIENT_ID_KEY]
  if (typeof existing === 'string' && existing.trim()) return existing
  const created = randomUUID()
  try {
    writeAppSetting(settingsPath, ANALYTICS_CLIENT_ID_KEY, created)
  } catch {
    // unwritable settings file: report under a session-scoped id instead
  }
  return created
}

/** GA4 event name rules: letters/digits/underscore, must start with a letter */
const EVENT_NAME_RE = /^[A-Za-z][A-Za-z0-9_]{0,39}$/

export function isValidEventName(name: unknown): name is string {
  return typeof name === 'string' && EVENT_NAME_RE.test(name)
}

/** GA4 param value limit is 100 chars; clamp instead of dropping the event */
function clampParams(params: Record<string, string | number>): Record<string, string | number> {
  const out: Record<string, string | number> = {}
  for (const [key, value] of Object.entries(params)) {
    if (!EVENT_NAME_RE.test(key)) continue
    out[key] = typeof value === 'string' ? value.slice(0, 100) : value
  }
  return out
}

export interface AnalyticsOptions {
  keys: AnalyticsKeys | null
  /** evaluated only for a valid event while reporting is enabled */
  getClientId: () => string
  /** re-checked per event so the Settings toggle applies immediately */
  isEnabled: () => boolean
  /** attached to every event; evaluated per event so live-changeable values
   * (UI language) stay fresh */
  baseParams?: () => Record<string, string | number>
  /** injectable for tests; defaults to global fetch */
  fetchFn?: typeof fetch
}

const SEND_TIMEOUT_MS = 5000

export function createAnalytics(options: AnalyticsOptions): Analytics {
  const { keys, getClientId, isEnabled, baseParams = () => ({}), fetchFn = fetch } = options
  if (!keys) return { active: false, track: () => {} }

  const endpoint =
    'https://www.google-analytics.com/mp/collect' +
    `?measurement_id=${encodeURIComponent(keys.measurementId)}` +
    `&api_secret=${encodeURIComponent(keys.apiSecret)}`
  // one session per process run; GA4 uses it to group events and count
  // engaged sessions (together with engagement_time_msec)
  const sessionId = Date.now().toString()

  return {
    active: true,
    track(name, params = {}) {
      if (!isValidEventName(name)) return
      try {
        if (!isEnabled()) return
        const body = JSON.stringify({
          client_id: getClientId(),
          events: [
            {
              name,
              params: {
                session_id: sessionId,
                engagement_time_msec: 100,
                ...clampParams(baseParams()),
                ...clampParams(params),
              },
            },
          ],
        })
        void fetchFn(endpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body,
          signal: AbortSignal.timeout(SEND_TIMEOUT_MS),
        }).catch(() => {})
      } catch {
        // analytics must never surface an error into the app
      }
    },
  }
}
