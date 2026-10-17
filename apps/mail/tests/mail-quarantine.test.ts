import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  extractAddress,
  extractDomain,
  lookalikeDomain,
  replyToMismatch,
  sanitizeMailHtml,
} from '../src/renderer/reader/quarantine'

/**
 * Lane R quarantine tests: hostile markup is stripped, authored colors
 * survive, remote media is blocked by default, and the body layer never
 * emits a chrome token reference.
 */

describe('sanitizeMailHtml', () => {
  it('strips scripts, styles, frames, forms, and event handlers', () => {
    const { html } = sanitizeMailHtml(
      '<p onclick="steal()">Hi</p><script>evil()</script><style>p{color:red}</style>' +
        '<iframe src="https://evil.example/x"></iframe><form><input name="q"></form>',
    )
    expect(html).not.toContain('<script')
    expect(html).not.toContain('<style')
    expect(html).not.toContain('<iframe')
    expect(html).not.toContain('<form')
    expect(html).not.toContain('onclick')
    expect(html).not.toContain('evil()')
    expect(html).toContain('Hi')
  })

  it('neutralizes javascript: links but keeps http(s) links with safe rel', () => {
    const { html } = sanitizeMailHtml(
      '<a href="javascript:evil()">click</a><a href="https://example.com/x">ok</a>',
    )
    expect(html).not.toContain('javascript:')
    expect(html).toContain('click')
    expect(html).toContain('href="https://example.com/x"')
    expect(html).toContain('target="_blank"')
    expect(html).toContain('rel="noopener noreferrer"')
  })

  it('preserves authored inline colors and drops url()/var() payloads', () => {
    const { html } = sanitizeMailHtml(
      '<p style="color: #107c41; background-color: #fff4ce;">green</p>' +
        '<p style="color: var(--surface); background-image: url(https://evil.example/x.png);">bad</p>',
    )
    expect(html).toContain('#107c41')
    expect(html).toContain('#fff4ce')
    expect(html).not.toContain('var(--')
    expect(html).not.toContain('url(')
  })

  it('blocks remote images by default and reports the count', () => {
    const source =
      '<p>Banner</p><img src="https://tracker.example.com/pixel.png" alt="Weekly banner">' +
      '<img src="data:image/png;base64,iVBORw0KGgo=" alt="inline">'
    const blocked = sanitizeMailHtml(source)
    expect(blocked.blockedCount).toBe(1)
    expect(blocked.html).not.toContain('tracker.example.com')
    expect(blocked.html).toContain('mq-blocked')
    expect(blocked.html).toContain('data:image/png')
    expect(blocked.html).not.toContain('var(--')
  })

  it('keeps remote images when explicitly allowed', () => {
    const allowed = sanitizeMailHtml('<img src="https://tracker.example.com/pixel.png" alt="x">', {
      allowRemoteImages: true,
    })
    expect(allowed.blockedCount).toBe(0)
    expect(allowed.html).toContain('https://tracker.example.com/pixel.png')
  })

  it('drops app classes/ids but keeps quarantine-namespaced classes', () => {
    const { html } = sanitizeMailHtml(
      '<div class="app-shell secret" id="root"><pre class="mq-text">plain</pre></div>',
    )
    expect(html).not.toContain('app-shell')
    expect(html).not.toContain('id="root"')
    expect(html).toContain('mq-text')
    expect(html).toContain('plain')
  })
})

describe('mail-quarantine.css isolation', () => {
  it('references only app-scoped --mq-* custom properties', () => {
    // Vitest rewrites import.meta.url, so resolve from the package cwd.
    const css = readFileSync(join(process.cwd(), 'src/renderer/mail-quarantine.css'), 'utf8')
    const refs = [...css.matchAll(/var\(\s*(--[\w-]+)/g)].map((match) => match[1])
    expect(refs.length).toBeGreaterThan(0)
    for (const ref of refs) {
      expect(ref.startsWith('--mq-')).toBe(true)
    }
    expect(css).not.toContain('var(--surface')
    expect(css).not.toContain('var(--text')
    expect(css).not.toContain('var(--accent')
  })
})

describe('phishing helpers', () => {
  it('extracts bare addresses and domains', () => {
    expect(extractAddress('Acme <Billing@Example.COM>')).toBe('billing@example.com')
    expect(extractAddress('plain@example.net')).toBe('plain@example.net')
    expect(extractDomain('Acme <billing@example.com>')).toBe('example.com')
    expect(extractDomain('not-an-address')).toBe('')
  })

  it('flags Reply-To divergence only on real divergence', () => {
    expect(replyToMismatch(['a@example.com'], 'outside@example.org')).toBe(true)
    expect(replyToMismatch(['a@example.com'], 'other@example.com')).toBe(false)
    expect(replyToMismatch(['a@example.com'], undefined)).toBe(false)
    expect(replyToMismatch([], 'outside@example.org')).toBe(false)
  })

  it('flags internationalised and near-miss domains', () => {
    expect(lookalikeDomain('xn--acme-pay-example.com', [])).toBe(true)
    expect(lookalikeDomain('secure.xn--pple-43d.com', ['apple.com'])).toBe(true)
    expect(lookalikeDomain('paypa1.com', ['paypal.com'])).toBe(true)
    expect(lookalikeDomain('paypal.com', ['paypal.com'])).toBe(false)
    expect(lookalikeDomain('totally-different.example', ['paypal.com'])).toBe(false)
    expect(lookalikeDomain('', ['paypal.com'])).toBe(false)
  })
})
