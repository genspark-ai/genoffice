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
    expect(next).toContain('case`insertBodyParagraphs`')
    expect(next).toContain('case`insertFilledParagraphs`')
    expect(next).toContain('case`insertTable`')
    expect(next).toContain('case`applyBodyCharFormat`')
    expect(next).toContain('/*genoffice-prepare-text-v7*/')
    expect(next).toContain('async insertFilledParagraphs(e,t,n){')
    expect(next).toContain('insertTable(e,t,n,r){')
    expect(next).toContain('findOrCreateFontId(String(a.fontName))')
    expect(next).toContain('selectionStart')
    expect(next).toContain('Gk(this.deps.wasm,e.target)')
    expect(next).not.toMatch(/,listBodyParagraphs\(\)\{this\.syncGeneration\(\)/)
    expect(next).not.toMatch(/,insertBodyParagraphs\(e,t,n\)\{this\.syncGeneration\(\)/)
    expect(next).not.toMatch(/,insertTable\(e,t,n,r\)\{/)
    expect(next).toContain('listBodyParagraphs(){this.syncGeneration()')
    expect(next).toContain('let s=r.getParagraphCount(e);')
    expect(next).toContain('r.splitParagraph(e,')
    expect(next).not.toContain('r.insertParagraph(e,t+a)')
    expect(next).toContain('async insertBodyParagraphs(e,t,n){if(await')
    expect(next).not.toContain('insertBodyParagraphs(e.section,e.index,e.count)')
    expect(next).toContain('selectionEnd:i}}}listBodyParagraphs(){this.syncGeneration()')
    expect(next).not.toMatch(/selectionEnd:i\}\}listBodyParagraphs\(\)\{this\.syncGeneration\(\)/)
    expect(exposePrepareTextCommand(next)).toBe(next)
  })

  it('closes a v4 prepareTextCommand that was missing its method brace', () => {
    const broken = [
      '/*genoffice-prepare-text-v4*/prepareTextCommand(){return{selectionEnd:i}}listBodyParagraphs(){this.syncGeneration();return []}',
      'insertBodyParagraphs(e,t,n){this.syncGeneration();let r=this.deps.wasm,i=Number(n);if(!Number.isInteger(e)||!Number.isInteger(t)||!Number.isInteger(i)||i<1)throw Error(`insert count must be a positive integer`);let s=r.getParagraphCount(e);if(!s)throw Error(`문서가 로드되지 않았습니다`);r.splitParagraph(e,0,0);return{section:e,index:t,count:i}}async applyTextCommand(e){return e}',
      'async insertBodyParagraphs(e,t,n){if(await $,!sA)throw Error(`Document agent is not initialized`);return sA.insertBodyParagraphs(e,t,n)},async applyTextCommand(e){return e}',
      'case`insertBodyParagraphs`:return n.insertBodyParagraphs(i.section,i.index,i.count);case`applyTextCommand`:return n.applyTextCommand(e)',
    ].join('')
    const next = exposePrepareTextCommand(broken)
    expect(next).toContain('selectionEnd:i}}}listBodyParagraphs(){this.syncGeneration()')
    expect(next).not.toMatch(/selectionEnd:i\}\}listBodyParagraphs\(\)\{this\.syncGeneration\(\)/)
    expect(exposePrepareTextCommand(next)).toBe(next)
  })

  it('upgrades a v2 prepare surface to include tables and inserts', () => {
    const v2 = [
      'try{Gk(this.deps.wasm,i),this.currentFormat(),a=!0}catch{a=!1}',
      '/*genoffice-prepare-text-v2*/prepareTextCommand(){return 1},setField(e,t){this.syncGeneration();return this.deps.wasm.setFieldValueByName(String(e??``),String(t??``))}async applyTextCommand(e){return e}',
      'async setField(e,t){if(await $,!sA)throw Error(`Document agent is not initialized`);return sA.setField(e,t)},async applyTextCommand(e){return e}',
      'case`setField`:return n.setField(i.name,i.value);case`applyTextCommand`:return n.applyTextCommand(e)',
    ].join(';')
    const next = exposePrepareTextCommand(v2)
    expect(next).toContain('/*genoffice-prepare-text-v7*/')
    expect(next).toContain('listTables(){')
    expect(next).toContain('case`listTables`')
    expect(next).toContain('case`replaceCell`')
    expect(next).toContain('case`insertBodyParagraphs`')
    expect(next).toContain('case`insertFilledParagraphs`')
    expect(next).toContain('case`insertTable`')
    expect(exposePrepareTextCommand(next)).toBe(next)
  })

  it('upgrades a v3 table surface to insertBodyParagraphs', () => {
    const v3 = [
      'try{Gk(this.deps.wasm,i),this.currentFormat(),a=!0}catch{a=!1}',
      '/*genoffice-prepare-text-v3*/prepareTextCommand(){return 1}setField(e,t){this.syncGeneration();return this.deps.wasm.setFieldValueByName(String(e??``),String(t??``))}listTables(){this.syncGeneration();return []}replaceCell(e,t,n,r,i){this.syncGeneration();let a=this.deps.wasm,o=a.getCellParagraphLength(e,t,n,r,0),s=a.replaceTextInCellDeferredPagination(e,t,n,r,0,0,o,String(i??``));if(typeof s==`string`)try{s=JSON.parse(s)}catch{}return s}async applyTextCommand(e){return e}',
      'async replaceCell(e,t,n,r,i){if(await $,!sA)throw Error(`Document agent is not initialized`);return sA.replaceCell(e,t,n,r,i)},async applyTextCommand(e){return e}',
      'case`replaceCell`:return n.replaceCell(i.section,i.paragraph,i.control,i.cellIndex,i.text);case`applyTextCommand`:return n.applyTextCommand(e)',
    ].join(';')
    const next = exposePrepareTextCommand(v3)
    expect(next).toContain('/*genoffice-prepare-text-v7*/')
    expect(next).toContain('async insertFilledParagraphs(e,t,n){')
    expect(next).toContain('let s=r.getParagraphCount(e);')
    expect(next).toContain('r.splitParagraph(e,')
    expect(next).toContain('case`insertBodyParagraphs`')
    expect(next).toContain('case`insertTable`')
    expect(next).not.toMatch(/,insertBodyParagraphs\(e,t,n\)\{this\.syncGeneration\(\)/)
    expect(next).toContain('async insertBodyParagraphs(e,t,n){if(await')
    expect(next).not.toContain('insertBodyParagraphs(e.section,e.index,e.count)')
    expect(exposePrepareTextCommand(next)).toBe(next)
  })

  it('upgrades a v5 fill surface to insertTable and format', () => {
    const v5 = [
      '/*genoffice-prepare-text-v5*/prepareTextCommand(){return{selectionEnd:i}}}listBodyParagraphs(){this.syncGeneration();return []}',
      'async insertFilledParagraphs(e,t,n){if(!Array.isArray(n)||n.length<1)throw Error(`insert texts must be a non-empty array`);return{section:e,index:t,count:n.length}}async applyTextCommand(e){return e}',
      'async insertFilledParagraphs(e,t,n){if(await $,!sA)throw Error(`Document agent is not initialized`);return sA.insertFilledParagraphs(e,t,n)},async applyTextCommand(e){return e}',
      'case`insertFilledParagraphs`:return n.insertFilledParagraphs(i.section,i.index,i.texts);case`applyTextCommand`:return n.applyTextCommand(e)',
    ].join('')
    const next = exposePrepareTextCommand(v5)
    expect(next).toContain('/*genoffice-prepare-text-v7*/')
    expect(next).toContain('insertTable(e,t,n,r){')
    expect(next).toContain('applyBodyCharFormat(e,t,n,r,i){')
    expect(next).toContain('findOrCreateFontId(String(a.fontName))')
    expect(next).toContain('case`insertTable`')
    expect(next).toContain('case`applyBodyParaFormat`')
    expect(next).not.toMatch(/,insertTable\(e,t,n,r\)\{/)
    expect(exposePrepareTextCommand(next)).toBe(next)
  })

  it('upgrades a v6 format surface to resolve font names', () => {
    const v6 = [
      '/*genoffice-prepare-text-v6*/prepareTextCommand(){return{selectionEnd:i}}}listBodyParagraphs(){this.syncGeneration();return []}',
      'applyBodyCharFormat(e,t,n,r,i){this.syncGeneration();let a=this.deps.wasm.applyCharFormat(e,t,n,r,i);if(typeof a==`string`)try{a=JSON.parse(a)}catch{}if(a&&a.ok===!1)throw Error(String(a.error||a.message||`applyCharFormat failed`));return a}async applyTextCommand(e){return e}',
    ].join('')
    const next = exposePrepareTextCommand(v6)
    expect(next).toContain('/*genoffice-prepare-text-v7*/')
    expect(next).toContain('findOrCreateFontId(String(a.fontName))')
    expect(exposePrepareTextCommand(next)).toBe(next)
  })

  it('repairs an insert handler that expected a params object', () => {
    const broken = [
      '/*genoffice-prepare-text-v4*/prepareTextCommand(){return{selectionEnd:i}}}listBodyParagraphs(){this.syncGeneration();return []}',
      'insertBodyParagraphs(e,t,n){this.syncGeneration();let r=this.deps.wasm,i=Number(n);if(!Number.isInteger(e)||!Number.isInteger(t)||!Number.isInteger(i)||i<1)throw Error(`insert count must be a positive integer`);let s=r.getParagraphCount(e);r.splitParagraph(e,0,0);return{section:e,index:t,count:i}}async applyTextCommand(e){return e}',
      'async insertBodyParagraphs(e){if(await $,!sA)throw Error(`Document agent is not initialized`);return sA.insertBodyParagraphs(e.section,e.index,e.count)},async applyTextCommand(e){return e}',
      'case`insertBodyParagraphs`:return n.insertBodyParagraphs(i.section,i.index,i.count);case`applyTextCommand`:return n.applyTextCommand(e)',
    ].join('')
    const next = exposePrepareTextCommand(broken)
    expect(next).toContain('async insertBodyParagraphs(e,t,n){if(await $,!sA)')
    expect(next).toContain('sA.insertBodyParagraphs(e,t,n)')
    expect(next).not.toContain('insertBodyParagraphs(e.section,e.index,e.count)')
    expect(exposePrepareTextCommand(next)).toBe(next)
  })

  it('repairs insertBodyParagraphs that called insertParagraph on the wasm facade', () => {
    const broken = [
      '/*genoffice-prepare-text-v4*/prepareTextCommand(){return{selectionEnd:i}}}listBodyParagraphs(){this.syncGeneration();return []}',
      'insertBodyParagraphs(e,t,n){this.syncGeneration();let r=this.deps.wasm,i=Number(n);if(!Number.isInteger(e)||!Number.isInteger(t)||!Number.isInteger(i)||i<1)throw Error(`insert count must be a positive integer`);for(let a=0;a<i;a+=1)r.insertParagraph(e,t+a);return{section:e,index:t,count:i}}async applyTextCommand(e){return e}',
      'async insertBodyParagraphs(e,t,n){if(await $,!sA)throw Error(`Document agent is not initialized`);return sA.insertBodyParagraphs(e,t,n)},async applyTextCommand(e){return e}',
      'case`insertBodyParagraphs`:return n.insertBodyParagraphs(i.section,i.index,i.count);case`applyTextCommand`:return n.applyTextCommand(e)',
    ].join('')
    const next = exposePrepareTextCommand(broken)
    expect(next).toContain('let s=r.getParagraphCount(e);')
    expect(next).toContain('r.splitParagraph(e,')
    expect(next).toContain('async insertFilledParagraphs(e,t,n){')
    expect(next).not.toContain('r.insertParagraph(e,t+a)')
    expect(exposePrepareTextCommand(next)).toBe(next)
  })

  it('repairs insertBodyParagraphs that used borrowDocumentHandle insertParagraph', () => {
    const broken = [
      '/*genoffice-prepare-text-v4*/prepareTextCommand(){return{selectionEnd:i}}}listBodyParagraphs(){this.syncGeneration();return []}',
      'insertBodyParagraphs(e,t,n){this.syncGeneration();let r=this.deps.wasm.borrowDocumentHandle();if(!r)throw Error(`문서가 로드되지 않았습니다`);let i=Number(n);if(!Number.isInteger(e)||!Number.isInteger(t)||!Number.isInteger(i)||i<1)throw Error(`insert count must be a positive integer`);for(let a=0;a<i;a+=1)r.insertParagraph(e,t+a);return{section:e,index:t,count:i}}async applyTextCommand(e){return e}',
      'async insertBodyParagraphs(e,t,n){if(await $,!sA)throw Error(`Document agent is not initialized`);return sA.insertBodyParagraphs(e,t,n)},async applyTextCommand(e){return e}',
      'case`insertBodyParagraphs`:return n.insertBodyParagraphs(i.section,i.index,i.count);case`applyTextCommand`:return n.applyTextCommand(e)',
    ].join('')
    const next = exposePrepareTextCommand(broken)
    expect(next).toContain('r.splitParagraph(e,')
    expect(next).not.toContain('borrowDocumentHandle()')
    expect(next).not.toContain('r.insertParagraph(e,t+a)')
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
