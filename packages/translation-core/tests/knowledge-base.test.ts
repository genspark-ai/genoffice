import { describe, expect, it } from 'vitest'

import {
  KnowledgeBase,
  SCHEMA_IDS,
  SCOPES,
  type KnowledgeBaseFileSystem,
  type KBStore,
} from '../src/knowledge-base'

class MemoryFS implements KnowledgeBaseFileSystem {
  files = new Map<string, string>()
  async mkdir(p: string) {
    void p
  }
  async readFile(p: string) {
    const v = this.files.get(p)
    if (v === undefined) {
      const err = new Error('ENOENT') as NodeJS.ErrnoException
      err.code = 'ENOENT'
      throw err
    }
    return v
  }
  async writeFile(p: string, contents: string) {
    this.files.set(p, contents)
  }
  async rename(from: string, to: string) {
    const v = this.files.get(from)
    if (v !== undefined) {
      this.files.set(to, v)
      this.files.delete(from)
    }
  }
}

function makeFS(): MemoryFS {
  return new MemoryFS()
}

describe('KnowledgeBase', () => {
  it('upserts entries into the right schema by discriminator', () => {
    const kb = new KnowledgeBase({ filePath: '/tmp/kb.json', fileSystem: makeFS() })
    const termId = kb.upsert({
      id: 't1',
      scope: 'company',
      priority: 3,
      sourceTerm: '克重',
      targetTerm: 'GSM',
    })
    expect(termId).toMatchObject({ sourceTerm: '克重' })
    const brand = kb.upsert({
      id: 'b1',
      scope: 'global',
      priority: 5,
      word: 'YKK',
      policy: 'neverTranslate',
    })
    expect(kb.list({ schema: 'term' })).toHaveLength(1)
    expect(kb.list({ schema: 'brand' })).toHaveLength(1)
    expect(kb.list()).toHaveLength(2)
    void brand
  })

  it('replaces an entry with the same id on upsert', () => {
    const kb = new KnowledgeBase({ fileSystem: makeFS() })
    kb.upsert({
      id: 't1',
      scope: 'company',
      priority: 3,
      sourceTerm: '克重',
      targetTerm: 'GSM',
    })
    kb.upsert({
      id: 't1',
      scope: 'company',
      priority: 3,
      sourceTerm: '克重',
      targetTerm: 'Weight (g/m²)',
    })
    expect(kb.list({ schema: 'term' })).toHaveLength(1)
    expect((kb.list({ schema: 'term' })[0] as { targetTerm: string }).targetTerm).toBe(
      'Weight (g/m²)',
    )
  })

  it('resolve() orders rules scope > priority > id', () => {
    const kb = new KnowledgeBase({ fileSystem: makeFS() })
    kb.upsert({
      id: 't-company-3',
      scope: 'company',
      priority: 3,
      sourceTerm: '克重',
      targetTerm: 'GSM',
    })
    kb.upsert({
      id: 't-customer-3',
      scope: 'customer',
      priority: 3,
      sourceTerm: '克重',
      targetTerm: 'Weight',
    })
    kb.upsert({
      id: 't-customer-9',
      scope: 'customer',
      priority: 9,
      sourceTerm: '面料',
      targetTerm: 'Fabric',
    })
    kb.upsert({
      id: 'f-company',
      scope: 'company',
      priority: 3,
      forbiddenText: 'all cotton',
      replacement: '100% cotton',
    })
    kb.upsert({
      id: 's-company',
      scope: 'company',
      priority: 3,
      name: 'formal',
      description: 'use formal business tone',
    })

    const r = kb.resolve({ sourceLang: 'en-US', targetLang: 'zh-CN' })
    // Highest priority (customer / 9) wins over the customer/3 entry.
    expect((r.terms[0] as { sourceTerm: string }).sourceTerm).toBe('面料')
    expect((r.terms[0] as { targetTerm: string }).targetTerm).toBe('Fabric')
    // The customer/3 entry still surfaces — the company-scope duplicate loses.
    const customerThree = r.terms.find(
      (t) => (t as { sourceTerm: string }).sourceTerm === '克重',
    )
    expect(customerThree).toBeDefined()
    expect((customerThree as { targetTerm: string }).targetTerm).toBe('Weight')
    expect(r.forbidden).toHaveLength(1)
    expect(r.styleRules).toHaveLength(1)
    expect(r.brands).toHaveLength(0)
    // customer-preference stays empty because customerName is unset
    expect(r.customerPreferences).toHaveLength(0)
    expect(r.promptBlock).toContain('Translation rules (en-US -> zh-CN)')
    expect(r.promptBlock).toContain('Weight')
    expect(r.promptBlock).toContain('100% cotton')
  })

  it('resolve() filters by language pair and customerName', () => {
    const kb = new KnowledgeBase({ fileSystem: makeFS() })
    kb.upsert({
      id: 't-1',
      scope: 'global',
      priority: 1,
      sourceTerm: 'pin',
      targetTerm: '别针',
      sourceLang: 'en',
      targetLang: 'zh-CN',
    })
    kb.upsert({
      id: 't-2',
      scope: 'global',
      priority: 1,
      sourceTerm: 'pin',
      targetLang: 'zh-CN',
      sourceLang: 'en',
      targetTerm: '徽章',
    })
    // Same source but different target lang — only the matching pair should resolve.
    kb.upsert({
      id: 't-3',
      scope: 'global',
      priority: 1,
      sourceTerm: 'pin',
      targetTerm: '針',
      sourceLang: 'en',
      targetLang: 'zh-TW',
    })
    kb.upsert({
      id: 'cp-1',
      scope: 'customer',
      priority: 5,
      customerName: 'Nike',
      preferenceType: 'fabricUnit',
      value: 'GSM',
    })
    kb.upsert({
      id: 'cp-2',
      scope: 'customer',
      priority: 5,
      customerName: 'Adidas',
      preferenceType: 'fabricUnit',
      value: 'g/m²',
    })

    const en = kb.resolve({ sourceLang: 'en', targetLang: 'zh-CN' })
    expect(en.terms.map((t) => (t as { targetTerm: string }).targetTerm)).toEqual(
      expect.arrayContaining(['别针', '徽章']),
    )
    expect(en.terms.find((t) => (t as { targetTerm: string }).targetTerm === '針')).toBeUndefined()

    const nike = kb.resolve({
      sourceLang: 'en',
      targetLang: 'zh-CN',
      customerName: 'Nike',
    })
    expect(nike.customerPreferences).toHaveLength(1)
    expect((nike.customerPreferences[0] as { customerName: string }).customerName).toBe('Nike')
  })

  it('save() then load() round-trips', async () => {
    const fs = makeFS()
    const path = '/tmp/kb.json'
    const kb1 = new KnowledgeBase({ filePath: path, fileSystem: fs })
    kb1.upsert({
      id: 't-1',
      scope: 'company',
      priority: 3,
      sourceTerm: '克重',
      targetTerm: 'GSM',
    })
    kb1.upsert({
      id: 'b-1',
      scope: 'global',
      priority: 5,
      word: 'YKK',
      policy: 'neverTranslate',
    })
    await kb1.save()
    expect(kb1.isDirty()).toBe(false)

    const kb2 = new KnowledgeBase({ filePath: path, fileSystem: fs })
    await kb2.load()
    expect(kb2.list()).toHaveLength(2)
    expect(kb2.list({ schema: 'term' })).toHaveLength(1)
    expect(kb2.list({ schema: 'brand' })).toHaveLength(1)
  })

  it('load() returns an empty store when the file is missing', async () => {
    const fs = makeFS()
    const kb = new KnowledgeBase({ filePath: '/tmp/missing.json', fileSystem: fs })
    await kb.load()
    expect(kb.list()).toEqual([])
  })

  it('remove() deletes the entry from every schema bucket', () => {
    const kb = new KnowledgeBase({ fileSystem: makeFS() })
    kb.upsert({
      id: 'shared-id',
      scope: 'company',
      priority: 3,
      sourceTerm: '克重',
      targetTerm: 'GSM',
    })
    kb.upsert({
      id: 'shared-id',
      scope: 'global',
      priority: 5,
      word: 'YKK',
      policy: 'neverTranslate',
    })
    expect(kb.remove('shared-id')).toBe(true)
    expect(kb.remove('shared-id')).toBe(false)
    expect(kb.list()).toEqual([])
  })

  it('every schema + scope is enumerable', () => {
    expect(SCHEMA_IDS).toHaveLength(5)
    expect(SCOPES).toEqual(['session', 'customer', 'project', 'company', 'global'])
  })

  it('seed constructor populates the store', () => {
    const seed: KBStore = {
      'trade.translation.term': [
        {
          id: 't-1',
          scope: 'company',
          priority: 3,
          sourceTerm: '克重',
          targetTerm: 'GSM',
        },
      ],
    }
    const kb = new KnowledgeBase({ seed, fileSystem: makeFS() })
    expect(kb.list({ schema: 'term' })).toHaveLength(1)
  })
})
