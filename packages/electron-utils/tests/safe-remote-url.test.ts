import { afterEach, describe, expect, it, vi } from 'vitest'
import { lookup } from 'node:dns/promises'
import { Agent } from 'undici'
import {
  fetchWithSsrfGuard,
  isBlockedAddress,
  isFakeIpDnsArtifact,
  isSafeRemoteUrl,
  pinnedLookup,
  resolveSafeRemoteTarget,
  type SafeRemoteTarget,
} from '../src/safe-remote-url'

// hostname checks resolve through node:dns/promises; tests stub it so VPN
// fake-ip answers can be simulated without touching the real resolver
vi.mock('node:dns/promises', async (importOriginal) => ({
  ...(await importOriginal<typeof import('node:dns/promises')>()),
  lookup: vi.fn(),
}))

const dnsResult = (...addresses: string[]) =>
  vi.mocked(lookup).mockResolvedValue(addresses.map((address) => ({ address, family: 4 })) as never)

afterEach(() => {
  vi.mocked(lookup).mockReset()
})

describe('isBlockedAddress', () => {
  it.each([
    '127.0.0.1',
    '169.254.169.254', // cloud metadata
    '10.1.2.3',
    '172.20.0.5',
    '192.168.1.1',
    '0.0.0.0',
    '100.64.0.1',
    '198.18.1.66', // VPN fake-ip placeholder
    '255.255.255.255',
  ])('blocks non-public IPv4 %s', (ip) => {
    expect(isBlockedAddress(ip)).toBe(true)
  })

  it.each(['::1', '::', 'fe80::1', 'fd00::1', 'ff02::1'])('blocks non-public IPv6 %s', (ip) => {
    expect(isBlockedAddress(ip)).toBe(true)
  })

  // URL normalizes [::ffff:127.0.0.1] to ::ffff:7f00:1, so the hex form must be
  // classified through the IPv4 rules rather than read as a plain IPv6 address
  it.each([
    '::ffff:7f00:1', // ::ffff:127.0.0.1
    '::ffff:a9fe:a9fe', // ::ffff:169.254.169.254
    '::ffff:a00:1', // ::ffff:10.0.0.1
    '::7f00:1', // deprecated IPv4-compatible ::127.0.0.1
    '::a9fe:a9fe',
  ])('blocks IPv4-mapped/compatible form %s', (ip) => {
    expect(isBlockedAddress(ip)).toBe(true)
  })

  it.each(['8.8.8.8', '1.1.1.1', '93.184.216.34', '2001:4860:4860::8888'])(
    'allows public address %s',
    (ip) => {
      expect(isBlockedAddress(ip)).toBe(false)
    },
  )

  it('treats non-addresses as unsafe', () => {
    expect(isBlockedAddress('example.com')).toBe(true)
    expect(isBlockedAddress('')).toBe(true)
  })
})

describe('isSafeRemoteUrl', () => {
  it.each([
    'file:///etc/passwd',
    'javascript:alert(1)',
    'data:image/png;base64,AAAA',
    'ftp://example.com/x.png',
  ])('rejects non-http scheme %s', async (url) => {
    await expect(isSafeRemoteUrl(url)).resolves.toBe(false)
  })

  it.each([
    'http://127.0.0.1/x.png',
    'http://169.254.169.254/latest/meta-data/',
    'http://[::1]/x.png',
    'http://[::ffff:127.0.0.1]/x.png',
    'http://10.0.0.1/x.png',
    'http://198.18.9.9/x.png', // fake-ip range as a URL literal is still refused
  ])('rejects literal internal target %s', async (url) => {
    await expect(isSafeRemoteUrl(url)).resolves.toBe(false)
  })

  it.each([
    'http://localhost:8080/x.png',
    'http://foo.localhost/x.png',
    'http://printer.local/x.png',
    'http://metadata.internal/x.png',
  ])('rejects internal hostname %s', async (url) => {
    await expect(isSafeRemoteUrl(url)).resolves.toBe(false)
  })

  it.each([
    'http://localhost./x.png',
    'http://LOCALHOST./x.png',
    'http://printer.local./x.png',
    'http://metadata.internal./x.png',
  ])('rejects trailing-dot internal hostname %s without DNS', async (url) => {
    await expect(isSafeRemoteUrl(url)).resolves.toBe(false)
  })

  it('rejects malformed input and non-strings', async () => {
    await expect(isSafeRemoteUrl('not a url')).resolves.toBe(false)
    await expect(isSafeRemoteUrl(undefined)).resolves.toBe(false)
    await expect(isSafeRemoteUrl(42)).resolves.toBe(false)
  })

  it('allows a public literal address', async () => {
    await expect(isSafeRemoteUrl('https://8.8.8.8/x.png')).resolves.toBe(true)
  })

  it('classifies the fake-ip range as a DNS artifact, not a public address', () => {
    expect(isFakeIpDnsArtifact('198.18.1.66')).toBe(true)
    expect(isFakeIpDnsArtifact('198.19.255.255')).toBe(true)
    expect(isFakeIpDnsArtifact('198.20.0.1')).toBe(false)
    expect(isFakeIpDnsArtifact('8.8.8.8')).toBe(false)
  })

  it('allows a hostname whose VPN resolver answered with fake-ip placeholders', async () => {
    dnsResult('198.18.1.66')
    await expect(isSafeRemoteUrl('https://www.alamy.com/x.png')).resolves.toBe(true)
  })

  it.each([['169.254.169.254'], ['10.0.0.7'], ['192.168.1.50']])(
    'still rejects a hostname resolving into %s',
    async (address) => {
      dnsResult(address)
      await expect(isSafeRemoteUrl('https://internal.example/x.png')).resolves.toBe(false)
    },
  )

  it('rejects a name whose answers mix fake-ip with a real internal address', async () => {
    dnsResult('198.18.1.66', '192.168.1.50')
    await expect(isSafeRemoteUrl('https://mixed.example/x.png')).resolves.toBe(false)
  })
})

describe('resolveSafeRemoteTarget', () => {
  it('hands back the validated addresses so a connection can bind to them', async () => {
    dnsResult('93.184.216.34', '93.184.216.35')
    const target = await resolveSafeRemoteTarget('https://www.example.com/x.png')
    expect(target).toMatchObject({
      url: 'https://www.example.com/x.png',
      hostname: 'www.example.com',
      literal: false,
      addresses: [
        { address: '93.184.216.34', family: 4 },
        { address: '93.184.216.35', family: 4 },
      ],
    })
  })

  it('marks literal targets and carries the literal itself as the address', async () => {
    const target = await resolveSafeRemoteTarget('https://8.8.8.8/x.png')
    expect(target).toMatchObject({
      hostname: '8.8.8.8',
      literal: true,
      addresses: [{ address: '8.8.8.8', family: 4 }],
    })
    expect(vi.mocked(lookup)).not.toHaveBeenCalled()
  })

  it('carries fake-ip answers as valid targets (the connection rides the tunnel)', async () => {
    dnsResult('198.18.1.66')
    const target = await resolveSafeRemoteTarget('https://www.alamy.com/x.png')
    expect(target?.addresses).toEqual([{ address: '198.18.1.66', family: 4 }])
  })
})

describe('pinnedLookup', () => {
  const target: SafeRemoteTarget = {
    url: 'https://cdn.example/x.png',
    hostname: 'cdn.example',
    literal: false,
    addresses: [
      { address: '198.18.1.66', family: 4 },
      { address: '2001:db8::1', family: 6 },
    ],
  }

  it('answers the single-address shape net.connect expects without all:true', () => {
    const cb = vi.fn()
    pinnedLookup(target)('cdn.example', {}, cb)
    expect(cb).toHaveBeenCalledWith(null, '198.18.1.66', 4)
  })

  it('honors the connect-time family hint on the single-address path', () => {
    const cb = vi.fn()
    pinnedLookup(target)('cdn.example', { family: 6 }, cb)
    expect(cb).toHaveBeenCalledWith(null, '2001:db8::1', 6)
  })

  it('answers the happy-eyeballs array shape under all:true', () => {
    const cb = vi.fn()
    pinnedLookup(target)('cdn.example', { all: true }, cb)
    expect(cb).toHaveBeenCalledWith(null, target.addresses)
  })

  it('normalizes case, brackets and trailing dots before comparing the host', () => {
    const cb = vi.fn()
    pinnedLookup(target)('CDN.EXAMPLE.', {}, cb)
    expect(cb).toHaveBeenCalledWith(null, '198.18.1.66', 4)
  })

  it('fails closed on a hostname that was not validated', () => {
    const cb = vi.fn()
    pinnedLookup(target)('evil.example', {}, cb)
    expect(cb).toHaveBeenCalledTimes(1)
    expect(cb.mock.calls[0][0]).toBeInstanceOf(Error)
    expect(cb.mock.calls[0][1]).toBe('')
  })
})

describe('fetchWithSsrfGuard', () => {
  const res = (status: number, location?: string) =>
    new Response(null, {
      status,
      headers: location ? { location } : undefined,
    })

  it('rejects a redirect from a public host to an internal address', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(res(302, 'http://169.254.169.254/latest/meta-data/'))
    const out = await fetchWithSsrfGuard('https://8.8.8.8/start.png', { fetchImpl })
    expect(out).toBeNull()
    // the internal hop is never requested
    expect(fetchImpl).toHaveBeenCalledTimes(1)
    expect(fetchImpl.mock.calls[0][0]).toBe('https://8.8.8.8/start.png')
  })

  it('follows a redirect between public hosts and returns the final response', async () => {
    const final = new Response('ok', { status: 200 })
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(res(301, 'https://1.1.1.1/final.png'))
      .mockResolvedValueOnce(final)
    const out = await fetchWithSsrfGuard('https://8.8.8.8/start.png', { fetchImpl })
    expect(out).toBe(final)
    expect(fetchImpl.mock.calls[1][0]).toBe('https://1.1.1.1/final.png')
  })

  it('resolves relative Location headers against the current URL', async () => {
    const final = new Response('ok', { status: 200 })
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(res(302, '/moved.png'))
      .mockResolvedValueOnce(final)
    await fetchWithSsrfGuard('https://8.8.8.8/a/b.png', { fetchImpl })
    expect(fetchImpl.mock.calls[1][0]).toBe('https://8.8.8.8/moved.png')
  })

  it('gives up on a redirect loop instead of following forever', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(res(302, 'https://8.8.8.8/loop.png'))
    const out = await fetchWithSsrfGuard('https://8.8.8.8/loop.png', {
      fetchImpl,
      maxRedirects: 3,
    })
    expect(out).toBeNull()
    expect(fetchImpl).toHaveBeenCalledTimes(4) // initial + 3 hops
  })

  it('normalizes a non-finite redirect budget instead of looping forever', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(res(302, 'https://8.8.8.8/loop.png'))
    const out = await fetchWithSsrfGuard('https://8.8.8.8/loop.png', {
      fetchImpl,
      maxRedirects: Infinity,
    })
    expect(out).toBeNull()
    expect(fetchImpl.mock.calls.length).toBeLessThanOrEqual(11) // default 5, hard cap 10
  })

  it('returns null when a redirect has no Location header', async () => {
    const fetchImpl = vi.fn().mockResolvedValueOnce(res(302))
    await expect(fetchWithSsrfGuard('https://8.8.8.8/x.png', { fetchImpl })).resolves.toBeNull()
  })

  it('passes headers through and disables automatic redirects', async () => {
    const fetchImpl = vi.fn().mockResolvedValueOnce(new Response('ok', { status: 200 }))
    await fetchWithSsrfGuard('https://8.8.8.8/x.png', {
      fetchImpl,
      headers: { 'User-Agent': 'test-agent' },
    })
    expect(fetchImpl.mock.calls[0][1]).toMatchObject({
      headers: { 'User-Agent': 'test-agent' },
      redirect: 'manual',
    })
  })

  it('never requests a target that fails the initial check', async () => {
    const fetchImpl = vi.fn()
    await expect(
      fetchWithSsrfGuard('http://169.254.169.254/latest/meta-data/', { fetchImpl }),
    ).resolves.toBeNull()
    expect(fetchImpl).not.toHaveBeenCalled()
  })

  it('pins hostname hops with a dispatcher bound to the validated addresses', async () => {
    dnsResult('198.18.1.66')
    const final = new Response('ok', { status: 200 })
    const fetchImpl = vi.fn().mockResolvedValue(final)
    await fetchWithSsrfGuard('https://www.alamy.com/x.png', { fetchImpl })
    const dispatcher = (fetchImpl.mock.calls[0][1] as { dispatcher?: unknown }).dispatcher
    expect(dispatcher).toBeInstanceOf(Agent)
  })

  it('sends literal-IP hops without a dispatcher (nothing to pin)', async () => {
    const final = new Response('ok', { status: 200 })
    const fetchImpl = vi.fn().mockResolvedValue(final)
    await fetchWithSsrfGuard('https://8.8.8.8/x.png', { fetchImpl })
    expect((fetchImpl.mock.calls[0][1] as { dispatcher?: unknown }).dispatcher).toBeUndefined()
  })

  it('skips pinning when a forward proxy owns the global dispatcher', async () => {
    const undici = await import('undici')
    const prior = undici.getGlobalDispatcher()
    undici.setGlobalDispatcher(new undici.ProxyAgent('http://127.0.0.1:9'))
    try {
      dnsResult('198.18.1.66')
      const final = new Response('ok', { status: 200 })
      const fetchImpl = vi.fn().mockResolvedValue(final)
      await fetchWithSsrfGuard('https://www.alamy.com/x.png', { fetchImpl })
      // the hop rides the proxy; the gate's per-hop validation still ran above
      expect((fetchImpl.mock.calls[0][1] as { dispatcher?: unknown }).dispatcher).toBeUndefined()
      expect(fetchImpl).toHaveBeenCalledTimes(1)
    } finally {
      undici.setGlobalDispatcher(prior)
    }
  })

  it('still pins when the global slot holds a foreign plain dispatcher', async () => {
    const undici = await import('undici')
    const prior = undici.getGlobalDispatcher()
    // Node's bundled undici materializes its own wrapper into the shared
    // slot once anything touches global fetch/Response — it is not a proxy,
    // so binding must survive it (duck-typed symbol detection, no instanceof)
    undici.setGlobalDispatcher({ dispatch: () => {} } as never)
    try {
      dnsResult('198.18.1.66')
      const final = new Response('ok', { status: 200 })
      const fetchImpl = vi.fn().mockResolvedValue(final)
      await fetchWithSsrfGuard('https://www.alamy.com/x.png', { fetchImpl })
      expect((fetchImpl.mock.calls[0][1] as { dispatcher?: unknown }).dispatcher).toBeDefined()
    } finally {
      undici.setGlobalDispatcher(prior)
    }
  })

  it('refuses a redirect onto a literal fake-ip address', async () => {
    const fetchImpl = vi.fn().mockResolvedValueOnce(res(302, 'http://198.18.9.9/payload'))
    const out = await fetchWithSsrfGuard('https://8.8.8.8/r.png', { fetchImpl })
    expect(out).toBeNull()
    expect(fetchImpl).toHaveBeenCalledTimes(1)
  })
})
