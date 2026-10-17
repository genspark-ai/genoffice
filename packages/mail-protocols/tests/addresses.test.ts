import { describe, expect, it } from 'vitest'
import {
  decodeDisplayName,
  formatAddress,
  formatAddressList,
  parseAddress,
  parseAddressList,
  toEnvelopeStrings,
} from '../src/addresses'

describe('parseAddress', () => {
  it('parses name plus angle address', () => {
    expect(parseAddress('Alice Example <alice@example.com>')).toEqual({
      name: 'Alice Example',
      address: 'alice@example.com',
    })
  })

  it('parses a bare address with no name', () => {
    expect(parseAddress('bob@example.com')).toEqual({ name: '', address: 'bob@example.com' })
  })

  it('degrades hostile tokens to raw text without throwing', () => {
    expect(() => parseAddress('not-an-address<>><<')).not.toThrow()
    expect(parseAddress('')).toEqual({ name: '', address: '' })
  })
})

describe('parseAddressList', () => {
  it('splits comma lists respecting quoted names', () => {
    const list = parseAddressList('"Doe, John" <john@example.com>, anne@example.com')
    expect(list).toHaveLength(2)
    expect(list[0]).toEqual({ name: 'Doe, John', address: 'john@example.com' })
    expect(list[1]).toEqual({ name: '', address: 'anne@example.com' })
  })

  it('unwraps group syntax for envelope display', () => {
    const list = parseAddressList('Friends: anne@example.com, ben@example.com;')
    expect(list.map((a) => a.address)).toEqual(['anne@example.com', 'ben@example.com'])
  })

  it('returns empty for empty headers', () => {
    expect(parseAddressList('')).toEqual([])
  })
})

describe('decodeDisplayName', () => {
  it('decodes base64 and quoted-printable encoded words', () => {
    expect(decodeDisplayName('=?UTF-8?B?5byg5LiJ?=')).not.toContain('=?')
    expect(decodeDisplayName('=?UTF-8?Q?Hello_=E2=80=93_World?=?')).toContain('Hello')
  })

  it('passes unknown charsets through raw', () => {
    const raw = '=?X-UNKNOWN?B?abcd?='
    expect(decodeDisplayName(raw)).toBe(raw)
  })
})

describe('formatAddress', () => {
  it('round-trips through parse for common shapes', () => {
    const addrs = parseAddressList('"Doe, John" <john@example.com>, anne@example.com')
    const formatted = formatAddressList(addrs)
    expect(parseAddressList(formatted)).toEqual(addrs)
    expect(formatAddress({ name: '', address: 'x@example.com' })).toBe('x@example.com')
  })
})

describe('toEnvelopeStrings', () => {
  it('keeps addresses only', () => {
    expect(toEnvelopeStrings([{ name: 'A', address: 'a@example.com' }])).toEqual(['a@example.com'])
  })
})
