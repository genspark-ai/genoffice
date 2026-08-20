import { describe, expect, it } from 'vitest'
import { safeExternalUrl } from '@genoffice/electron-utils'
import { DOCS_LINK_PROTOCOLS } from '../src/shared/link-protocols'

const docsGate = (url: string): string | null =>
  safeExternalUrl(url, { allowedProtocols: DOCS_LINK_PROTOCOLS })

describe('aof-review:// links in Docs', () => {
  it('passes the docs link gate (routes to shell.openExternal)', () => {
    const url = 'aof-review://seek?file=Poker_Final_Table_CamA.mp4&tc=00:01:17:00'
    expect(docsGate(url)).toBe(url)
  })

  it('open?folder form passes too', () => {
    const url = 'aof-review://open?folder=%2FVolumes%2FFootage%2FProjX'
    expect(docsGate(url)).toBe(url)
  })

  it('http/https keep working', () => {
    expect(docsGate('https://artisansonfire.com')).toBe('https://artisansonfire.com')
    expect(docsGate('http://example.com')).toBe('http://example.com')
  })

  it('dangerous schemes stay blocked in the docs gate', () => {
    expect(docsGate('javascript:alert(1)')).toBeNull()
    expect(docsGate('file:///etc/passwd')).toBeNull()
    expect(docsGate('smb://attacker/share')).toBeNull()
    expect(docsGate('vscode://malicious')).toBeNull()
  })

  it('the suite-wide DEFAULT gate still rejects aof-review (docs-only opt-in)', () => {
    expect(safeExternalUrl('aof-review://seek?file=a.mp4&tc=00:00:01')).toBeNull()
  })
})
