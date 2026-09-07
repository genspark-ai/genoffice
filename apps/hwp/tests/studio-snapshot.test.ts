import { describe, expect, it } from 'vitest'
import {
  hasEmbedNewDoc,
  isPwaPath,
  exposePrepareTextCommand,
  hasPrepareTextCommand,
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

  it('exposes prepareTextCommand next to getSelectionContext', () => {
    const stock = [
      'try{Gk(this.deps.wasm,i),this.currentFormat(),a=!0}catch{a=!1}',
      'return{selectedTextSha256:o}}async applyTextCommand(e){return e}',
      'async getSelectionContext(){if(await $,!sA)throw Error(`Document agent is not initialized`);return sA.getSelectionContext()},async applyTextCommand(e){return e}',
      'case`getSelectionContext`:return ak(i,`getSelectionContext params`),n.getSelectionContext();case`applyTextCommand`:return n.applyTextCommand(e)',
    ].join(';')
    const next = exposePrepareTextCommand(stock)
    expect(hasPrepareTextCommand(next)).toBe(true)
    expect(next).toContain('prepareTextCommand(){')
    expect(next).toContain('case`prepareTextCommand`')
    expect(next).toContain('case`listBodyParagraphs`')
    expect(next).toContain('case`listTables`')
    expect(next).toContain('case`setField`')
    expect(next).toContain('selectionStart')
    expect(next).toContain('Gk(this.deps.wasm,e.target)')
    expect(next).not.toMatch(/,listBodyParagraphs\(\)\{this\.syncGeneration\(\)/)
    expect(next).toContain('listBodyParagraphs(){this.syncGeneration()')
    expect(exposePrepareTextCommand(next)).toBe(next)
  })

  it('upgrades a v2 prepare surface to include tables', () => {
    const v2 = [
      'try{Gk(this.deps.wasm,i),this.currentFormat(),a=!0}catch{a=!1}',
      '/*genoffice-prepare-text-v2*/prepareTextCommand(){return 1},setField(e,t){this.syncGeneration();return this.deps.wasm.setFieldValueByName(String(e??``),String(t??``))}async applyTextCommand(e){return e}',
      'async setField(e,t){if(await $,!sA)throw Error(`Document agent is not initialized`);return sA.setField(e,t)},async applyTextCommand(e){return e}',
      'case`setField`:return n.setField(i.name,i.value);case`applyTextCommand`:return n.applyTextCommand(e)',
    ].join(';')
    const next = exposePrepareTextCommand(v2)
    expect(next).toContain('/*genoffice-prepare-text-v3*/')
    expect(next).toContain('listTables(){')
    expect(next).toContain('case`listTables`')
    expect(next).toContain('case`replaceCell`')
    expect(exposePrepareTextCommand(next)).toBe(next)
  })

  it('fails loudly when the document-agent surface is no longer in the bundle', () => {
    expect(() => exposePrepareTextCommand('getSelectionContext applyTextCommand')).toThrow(
      'paragraph snapshot helper changed',
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
