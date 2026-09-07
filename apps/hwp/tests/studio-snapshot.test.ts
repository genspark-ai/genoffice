import { describe, expect, it } from 'vitest'
import {
  hasEmbedNewDoc,
  isPwaPath,
  keepEmbedNewDoc,
  stripAbandonedStudioPatches,
  stripPwaHtml,
} from '../scripts/studio-snapshot.mjs'

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

  it('keeps file:new-doc registered in embed so the host can create untitled docs', () => {
    const stock =
      'bA.registerAll(yA===`embed`?Ev.filter(e=>!sD.includes(e.id)):Ev),file:new-doc'
    const next = keepEmbedNewDoc(stock)
    expect(hasEmbedNewDoc(next)).toBe(true)
    expect(next).toContain('e.id===`file:new-doc`')
    expect(next).toContain('!sD.includes(e.id)')
    expect(keepEmbedNewDoc(next)).toBe(next)
  })

  it('fails loudly when the embed command filter is no longer in the bundle', () => {
    expect(() => keepEmbedNewDoc('file:new-doc registerAll embed')).toThrow(
      'embed command filter changed',
    )
  })

  it('removes abandoned page-turn patches from a local snapshot', () => {
    const patched = [
      '/*genoffice-eager-prefetch*/n()',
      '/*genoffice-prefetch-overscan*/for(let e of[s-2,s-1,c+1,c+2])',
      'flushDeferredPaginationIfNeeded(`before-navigation`,/*genoffice-nav-pagination*/!0)',
    ].join(';')
    const next = stripAbandonedStudioPatches(patched)
    expect(next).toContain('requestIdleCallback')
    expect(next).toContain('[s-1,c+1]')
    expect(next).toContain('before-navigation`,!1')
    expect(next).not.toContain('genoffice-eager-prefetch')
    expect(next).not.toContain('genoffice-prefetch-overscan')
    expect(next).not.toContain('genoffice-nav-pagination')
  })
})
