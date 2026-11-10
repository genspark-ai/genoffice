import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  ANALYTICS_CLIENT_ID_KEY,
  ANALYTICS_ENABLED_KEY,
  ANALYTICS_FIRST_LAUNCH_PENDING_KEY,
  INSTALL_FIRST_LAUNCH_EVENT,
  analyticsEnabledFrom,
  createAnalytics,
  ensureAnalyticsClientId,
  ensureAnalyticsClientState,
  extractAnalyticsKeys,
  extractPackagedAnalyticsKeys,
  isValidEventName,
  markAnalyticsFirstLaunchSent,
} from '../src/main/analytics'

/**
 * GA4 Measurement Protocol analytics (src/main/analytics.ts): keyless builds
 * must be a strict no-op, opted-out installs must not send, and network
 * failures must never propagate.
 */

const KEYS = { measurementId: 'G-TEST123', apiSecret: 'secret' }

function okFetch() {
  return vi.fn(() => Promise.resolve(new Response(null, { status: 204 })))
}

describe('extractAnalyticsKeys', () => {
  it('reads the injected block from package.json', () => {
    expect(extractAnalyticsKeys({ genofficeAnalytics: KEYS })).toEqual(KEYS)
  })

  it('trims whitespace around the values', () => {
    expect(
      extractAnalyticsKeys({
        genofficeAnalytics: { measurementId: ' G-1 ', apiSecret: ' s ' },
      }),
    ).toEqual({ measurementId: 'G-1', apiSecret: 's' })
  })

  it('returns null when the block is missing (source/fork builds)', () => {
    expect(extractAnalyticsKeys({ name: '@genoffice/shell' })).toBeNull()
    expect(extractAnalyticsKeys(null)).toBeNull()
    expect(extractAnalyticsKeys('nope')).toBeNull()
  })

  it('returns null when either credential is empty or not a string', () => {
    expect(
      extractAnalyticsKeys({ genofficeAnalytics: { measurementId: 'G-1', apiSecret: '' } }),
    ).toBeNull()
    expect(
      extractAnalyticsKeys({ genofficeAnalytics: { measurementId: 42, apiSecret: 's' } }),
    ).toBeNull()
    expect(extractAnalyticsKeys({ genofficeAnalytics: { measurementId: 'G-1' } })).toBeNull()
  })

  it('accepts metadata only for a packaged runtime', () => {
    expect(extractPackagedAnalyticsKeys({ genofficeAnalytics: KEYS }, true)).toEqual(KEYS)
    expect(extractPackagedAnalyticsKeys({ genofficeAnalytics: KEYS }, false)).toBeNull()
  })
})

describe('analyticsEnabledFrom', () => {
  it('defaults to on when the key is absent or malformed', () => {
    expect(analyticsEnabledFrom({})).toBe(true)
    expect(analyticsEnabledFrom({ [ANALYTICS_ENABLED_KEY]: 'no' })).toBe(true)
    expect(analyticsEnabledFrom({ [ANALYTICS_ENABLED_KEY]: 1 })).toBe(true)
  })

  it('disables only an explicit boolean opt-out', () => {
    expect(analyticsEnabledFrom({ [ANALYTICS_ENABLED_KEY]: false })).toBe(false)
    expect(analyticsEnabledFrom({ [ANALYTICS_ENABLED_KEY]: true })).toBe(true)
  })
})

describe('ensureAnalyticsClientId', () => {
  let dir: string
  let settingsPath: string

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'analytics-'))
    settingsPath = join(dir, 'app-settings.json')
  })

  afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  it('creates and persists a UUID on first call', () => {
    const id = ensureAnalyticsClientId(settingsPath)
    expect(id).toMatch(/^[0-9a-f-]{36}$/)
    const saved = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
    expect(saved[ANALYTICS_CLIENT_ID_KEY]).toBe(id)
  })

  it('returns the stored id on later calls', () => {
    const first = ensureAnalyticsClientId(settingsPath)
    expect(ensureAnalyticsClientId(settingsPath)).toBe(first)
  })

  it('keeps a pre-existing install id stable', () => {
    writeFileSync(
      settingsPath,
      JSON.stringify({ [ANALYTICS_CLIENT_ID_KEY]: 'existing-install-id' }),
    )
    expect(ensureAnalyticsClientId(settingsPath)).toBe('existing-install-id')
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({
      [ANALYTICS_CLIENT_ID_KEY]: 'existing-install-id',
    })
  })

  it('keeps unrelated settings intact', () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'ja' }))
    ensureAnalyticsClientId(settingsPath)
    const saved = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
    expect(saved.language).toBe('ja')
  })

  it('keeps a new install pending until its cohort-start event is delivered', () => {
    const created = ensureAnalyticsClientState(settingsPath)
    expect(created.clientId).toMatch(/^[0-9a-f-]{36}$/)
    expect(created.firstLaunchPending).toBe(true)
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))[ANALYTICS_FIRST_LAUNCH_PENDING_KEY]).toBe(
      true,
    )
    expect(ensureAnalyticsClientState(settingsPath)).toEqual(created)

    markAnalyticsFirstLaunchSent(settingsPath)
    expect(ensureAnalyticsClientState(settingsPath)).toEqual({
      clientId: created.clientId,
      firstLaunchPending: false,
    })
  })

  it('does not classify an id created before retention tracking as new', () => {
    writeFileSync(settingsPath, JSON.stringify({ [ANALYTICS_CLIENT_ID_KEY]: 'legacy-install-id' }))
    expect(ensureAnalyticsClientState(settingsPath)).toEqual({
      clientId: 'legacy-install-id',
      firstLaunchPending: false,
    })
  })
})

describe('isValidEventName', () => {
  it('accepts GA4-style names', () => {
    expect(isValidEventName('app_launch')).toBe(true)
    expect(isValidEventName('file_open')).toBe(true)
  })

  it('rejects invalid names', () => {
    expect(isValidEventName('')).toBe(false)
    expect(isValidEventName('9lives')).toBe(false)
    expect(isValidEventName('has space')).toBe(false)
    expect(isValidEventName('x'.repeat(41))).toBe(false)
    expect(isValidEventName(42)).toBe(false)
  })
})

describe('createAnalytics', () => {
  it('is a strict no-op without keys', () => {
    const fetchFn = okFetch()
    const getClientId = vi.fn(() => 'c')
    const shouldTrackFirstLaunch = vi.fn(() => true)
    const analytics = createAnalytics({
      keys: null,
      getClientId,
      isEnabled: () => true,
      shouldTrackFirstLaunch,
      fetchFn,
    })
    expect(analytics.active).toBe(false)
    analytics.track('app_launch')
    expect(fetchFn).not.toHaveBeenCalled()
    expect(getClientId).not.toHaveBeenCalled()
    expect(shouldTrackFirstLaunch).not.toHaveBeenCalled()
  })

  it('sends a Measurement Protocol payload with keys present', () => {
    const fetchFn = okFetch()
    const analytics = createAnalytics({
      keys: KEYS,
      getClientId: () => 'client-1',
      isEnabled: () => true,
      baseParams: () => ({ app_version: '1.0.0', platform: 'darwin' }),
      getCountryCode: () => 'us',
      fetchFn,
    })
    expect(analytics.active).toBe(true)
    analytics.track('file_open', { ext: 'docx' })

    expect(fetchFn).toHaveBeenCalledTimes(1)
    const [url, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe(
      'https://www.google-analytics.com/mp/collect?measurement_id=G-TEST123&api_secret=secret',
    )
    const payload = JSON.parse(init.body as string) as {
      client_id: string
      user_location?: { country_id: string }
      events: Array<{ name: string; params: Record<string, unknown> }>
    }
    expect(payload.client_id).toBe('client-1')
    expect(payload.user_location).toEqual({ country_id: 'US' })
    expect(payload.events).toHaveLength(1)
    expect(payload.events[0].name).toBe('file_open')
    expect(payload.events[0].params.ext).toBe('docx')
    expect(payload.events[0].params.app_version).toBe('1.0.0')
    expect(payload.events[0].params.platform).toBe('darwin')
    expect(payload.events[0].params.session_id).toMatch(/^\d+$/)
    expect(payload.events[0].params.engagement_time_msec).toBe(100)
  })

  it('keeps sending without geography when the OS country code is unavailable', () => {
    const fetchFn = okFetch()
    const analytics = createAnalytics({
      keys: KEYS,
      getClientId: () => 'c',
      isEnabled: () => true,
      getCountryCode: () => {
        throw new Error('locale unavailable')
      },
      fetchFn,
    })
    analytics.track('app_launch')

    const [, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit]
    const payload = JSON.parse(init.body as string) as Record<string, unknown>
    expect(payload.user_location).toBeUndefined()
    expect(payload.events).toHaveLength(1)
  })

  it('prepends one cohort-start event for a newly identified install', async () => {
    const fetchFn = okFetch()
    const shouldTrackFirstLaunch = vi.fn(() => true)
    const onFirstLaunchSent = vi.fn()
    const analytics = createAnalytics({
      keys: KEYS,
      getClientId: () => 'new-install-id',
      isEnabled: () => true,
      shouldTrackFirstLaunch,
      onFirstLaunchSent,
      baseParams: () => ({ app_version: '1.0.0' }),
      fetchFn,
    })

    analytics.track('app_launch')
    const firstPayload = JSON.parse((fetchFn.mock.calls[0][1] as RequestInit).body as string) as {
      events: Array<{ name: string; params: Record<string, unknown> }>
    }
    expect(firstPayload.events.map((event) => event.name)).toEqual([
      INSTALL_FIRST_LAUNCH_EVENT,
      'app_launch',
    ])
    expect(firstPayload.events[0].params.app_version).toBe('1.0.0')
    expect(firstPayload.events[0].params.session_id).toMatch(/^\d+$/)
    await vi.waitFor(() => expect(onFirstLaunchSent).toHaveBeenCalledTimes(1))

    analytics.track('file_open', { ext: 'docx' })
    const laterPayload = JSON.parse((fetchFn.mock.calls[1][1] as RequestInit).body as string) as {
      events: Array<{ name: string }>
    }
    expect(laterPayload.events.map((event) => event.name)).toEqual(['file_open'])
    expect(shouldTrackFirstLaunch).toHaveBeenCalledTimes(1)
  })

  it('retries the cohort-start event after a failed delivery', async () => {
    const fetchFn = vi
      .fn()
      .mockRejectedValueOnce(new Error('offline'))
      .mockResolvedValue(new Response(null, { status: 204 }))
    const onFirstLaunchSent = vi.fn()
    const analytics = createAnalytics({
      keys: KEYS,
      getClientId: () => 'new-install-id',
      isEnabled: () => true,
      shouldTrackFirstLaunch: () => true,
      onFirstLaunchSent,
      fetchFn,
    })

    analytics.track('app_launch')
    await vi.waitFor(() => expect(fetchFn).toHaveBeenCalledTimes(1))
    await Promise.resolve()
    analytics.track('file_open')

    const retryPayload = JSON.parse((fetchFn.mock.calls[1][1] as RequestInit).body as string) as {
      events: Array<{ name: string }>
    }
    expect(retryPayload.events.map((event) => event.name)).toEqual([
      INSTALL_FIRST_LAUNCH_EVENT,
      'file_open',
    ])
    await vi.waitFor(() => expect(onFirstLaunchSent).toHaveBeenCalledTimes(1))
  })

  it('honors the runtime gate and stops immediately after opt-out', () => {
    const fetchFn = okFetch()
    const getClientId = vi.fn(() => 'c')
    const getCountryCode = vi.fn(() => 'US')
    const shouldTrackFirstLaunch = vi.fn(() => false)
    let enabled = false
    const analytics = createAnalytics({
      keys: KEYS,
      getClientId,
      isEnabled: () => enabled,
      getCountryCode,
      shouldTrackFirstLaunch,
      fetchFn,
    })
    analytics.track('pre_consent')
    expect(fetchFn).not.toHaveBeenCalled()
    expect(getClientId).not.toHaveBeenCalled()
    expect(getCountryCode).not.toHaveBeenCalled()
    expect(shouldTrackFirstLaunch).not.toHaveBeenCalled()

    enabled = true
    analytics.track('post_consent')
    expect(fetchFn).toHaveBeenCalledTimes(1)
    expect(getClientId).toHaveBeenCalledTimes(1)
    expect(getCountryCode).toHaveBeenCalledTimes(1)
    expect(shouldTrackFirstLaunch).toHaveBeenCalledTimes(1)

    enabled = false
    analytics.track('after_opt_out')
    expect(fetchFn).toHaveBeenCalledTimes(1)
    expect(getClientId).toHaveBeenCalledTimes(1)
    expect(getCountryCode).toHaveBeenCalledTimes(1)
    expect(shouldTrackFirstLaunch).toHaveBeenCalledTimes(1)
  })

  it('creates and persists the client id only when an eligible consented event needs it', () => {
    const dir = mkdtempSync(join(tmpdir(), 'analytics-lazy-'))
    const settingsPath = join(dir, 'app-settings.json')
    const fetchFn = okFetch()
    let enabled = false
    try {
      const analytics = createAnalytics({
        keys: KEYS,
        getClientId: () => ensureAnalyticsClientId(settingsPath),
        isEnabled: () => enabled,
        fetchFn,
      })

      analytics.track('bad event')
      analytics.track('pre_consent')
      expect(existsSync(settingsPath)).toBe(false)
      expect(fetchFn).not.toHaveBeenCalled()

      enabled = true
      analytics.track('post_consent')
      expect(fetchFn).toHaveBeenCalledTimes(1)
      const saved = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
      expect(saved[ANALYTICS_CLIENT_ID_KEY]).toMatch(/^[0-9a-f-]{36}$/)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('re-evaluates base params per event (live language switches)', () => {
    const fetchFn = okFetch()
    let lang = 'en'
    const analytics = createAnalytics({
      keys: KEYS,
      getClientId: () => 'c',
      isEnabled: () => true,
      baseParams: () => ({ ui_lang: lang }),
      fetchFn,
    })
    analytics.track('app_launch')
    lang = 'ja'
    analytics.track('app_launch')
    const langOf = (call: unknown[]) =>
      (
        JSON.parse((call[1] as RequestInit).body as string) as {
          events: Array<{ params: Record<string, unknown> }>
        }
      ).events[0].params.ui_lang
    expect(langOf(fetchFn.mock.calls[0])).toBe('en')
    expect(langOf(fetchFn.mock.calls[1])).toBe('ja')
  })

  it('drops events with invalid names instead of sending garbage', () => {
    const fetchFn = okFetch()
    const analytics = createAnalytics({
      keys: KEYS,
      getClientId: () => 'c',
      isEnabled: () => true,
      fetchFn,
    })
    analytics.track('bad name!')
    expect(fetchFn).not.toHaveBeenCalled()
  })

  it('drops invalid param keys and clamps oversized string values', () => {
    const fetchFn = okFetch()
    const analytics = createAnalytics({
      keys: KEYS,
      getClientId: () => 'c',
      isEnabled: () => true,
      fetchFn,
    })
    analytics.track('app_launch', { 'bad key': 'x', long: 'y'.repeat(200), n: 3 })
    const [, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit]
    const payload = JSON.parse(init.body as string) as {
      events: Array<{ params: Record<string, unknown> }>
    }
    expect(payload.events[0].params['bad key']).toBeUndefined()
    expect(payload.events[0].params.long).toBe('y'.repeat(100))
    expect(payload.events[0].params.n).toBe(3)
  })

  it('never throws when the network call fails', () => {
    const rejecting = vi.fn(() => Promise.reject(new Error('offline')))
    const throwing = vi.fn(() => {
      throw new Error('no fetch')
    })
    for (const fetchFn of [rejecting, throwing]) {
      const analytics = createAnalytics({
        keys: KEYS,
        getClientId: () => 'c',
        isEnabled: () => true,
        fetchFn: fetchFn as unknown as typeof fetch,
      })
      expect(() => analytics.track('app_launch')).not.toThrow()
    }
  })
})
