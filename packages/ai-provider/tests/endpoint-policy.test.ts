import { describe, expect, it } from 'vitest'
import {
  assertValidAiEndpoint,
  endpointKindForUrl,
  isLocalEndpointHost,
  validateAiEndpoint,
} from '../src/endpoint-policy'

describe('endpoint policy', () => {
  it('normalizes a provider path and classifies public HTTPS endpoints', () => {
    const result = validateAiEndpoint('https://api.example.com/v1///')
    expect(result).toEqual({ ok: true, normalized: 'https://api.example.com/v1', kind: 'cloud' })
    expect(endpointKindForUrl('https://api.example.com/v1')).toBe('cloud')
  })

  it('allows local HTTP servers by default', () => {
    expect(validateAiEndpoint('http://localhost:11434/v1').kind).toBe('local')
    expect(validateAiEndpoint('http://192.168.1.20:1234/v1').kind).toBe('local')
    expect(isLocalEndpointHost('fd00::1')).toBe(true)
    expect(isLocalEndpointHost('::ffff:127.0.0.1')).toBe(true)
  })

  it('rejects public HTTP and URL-embedded credentials', () => {
    expect(validateAiEndpoint('http://api.example.com/v1').ok).toBe(false)
    expect(validateAiEndpoint('https://user:secret@example.com/v1').ok).toBe(false)
    expect(validateAiEndpoint('https://example.com/v1?key=secret').ok).toBe(false)
  })

  it('supports an explicit policy for locked-down deployments', () => {
    const result = validateAiEndpoint('http://localhost:8000/v1', {
      allowLocal: false,
      allowInsecureHttp: false,
      allowUrlCredentials: false,
    })
    expect(result).toEqual({ ok: false, reason: 'Local network endpoints are disabled by policy' })
  })

  it('throws a concise error for invalid endpoints', () => {
    expect(() => assertValidAiEndpoint('file:///tmp/model')).toThrow('http:// or https://')
  })
})
