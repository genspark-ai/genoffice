import { describe, expect, it } from 'vitest'
import { eagerPagePrefetch, hasEagerPrefetch, isPwaPath, stripPwaHtml } from '../scripts/studio-snapshot.mjs'

describe('studio snapshot helpers', () => {
  it('strips the stock PWA registration from the published index', () => {
    const html = [
      '<link rel="stylesheet" href="/rhwp/assets/index.css">',
      '<link rel="manifest" href="/rhwp/manifest.webmanifest">',
      '<script id="vite-plugin-pwa:register-sw" src="/rhwp/registerSW.js"></script>',
      '<script type="module" src="/rhwp/assets/index.js"></script>',
    ].join('')
    const next = stripPwaHtml(html)
    expect(next).toContain('/rhwp/assets/index.js')
    expect(next).not.toContain('registerSW')
    expect(next).not.toContain('manifest.webmanifest')
  })

  it('rejects service-worker and manifest paths', () => {
    expect(isPwaPath('/rhwp/sw.js')).toBe(true)
    expect(isPwaPath('/rhwp/registerSW.js')).toBe(true)
    expect(isPwaPath('/rhwp/manifest.webmanifest')).toBe(true)
    expect(isPwaPath('/rhwp/assets/index.js')).toBe(false)
  })

  it('turns idle neighbor-page prefetch into an immediate paint', () => {
    const stock =
      'schedulePrefetchPages(e){if(typeof r.requestIdleCallback==`function`){this.deferredPrefetchTask={kind:`idle`,id:r.requestIdleCallback(n,{timeout:1e3})};return}this.deferredPrefetchTask={kind:`timeout`,id:window.setTimeout(n,250)}}'
    const next = eagerPagePrefetch(stock)
    expect(hasEagerPrefetch(next)).toBe(true)
    expect(next).toContain('n()')
    expect(next).not.toContain('requestIdleCallback(n,{timeout:1e3})')
    expect(eagerPagePrefetch(next)).toBe(next)
  })

  it('fails loudly when the idle-prefetch snippet is no longer in the bundle', () => {
    expect(() =>
      eagerPagePrefetch('schedulePrefetchPages(e){requestIdleCallback(n)}'),
    ).toThrow('prefetch idle deferral changed')
  })
})
