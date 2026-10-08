import { describe, expect, it } from 'vitest'

import {
  DEFAULT_SHEET_PROTECTION_ALLOW,
  applySheetProtection,
  parseSheetProtectionElement,
  readSheetProtection,
  sheetProtectionElement,
} from '../src/gateway/xlsx-protection'
import {
  hashSheetPassword,
  legacySheetPasswordHash,
  verifySheetPassword,
} from '../src/gateway/xlsx-protection-hash'

const SHEET =
  '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
  '<pageMargins left="0.7"/></worksheet>'

describe('legacy sheet password hash', () => {
  it('matches the values Excel writes', () => {
    // openpyxl's reference vector; the empty password is the bare XOR constant.
    expect(legacySheetPasswordHash('test')).toBe('CBEB')
    expect(legacySheetPasswordHash('')).toBe('CE4B')
  })

  it('verifies case-insensitively against the stored attribute', async () => {
    expect(await verifySheetPassword('test', { legacy: 'cbeb' })).toBe(true)
    expect(await verifySheetPassword('Test', { legacy: 'CBEB' })).toBe(false)
  })
})

describe('SHA-512 sheet password hash', () => {
  it('is deterministic for a fixed salt and spin count', async () => {
    const salt = new Uint8Array(16).fill(7)
    const first = await hashSheetPassword('secret', 10, salt)
    const second = await hashSheetPassword('secret', 10, salt)
    expect(first).toEqual(second)
    if ('legacy' in first) throw new Error('expected a hashed password')
    expect(first.algorithmName).toBe('SHA-512')
    expect(first.spinCount).toBe(10)
    expect(first.saltValue).toBe(Buffer.from(salt).toString('base64'))
    expect(Buffer.from(first.hashValue, 'base64')).toHaveLength(64)
  })

  it('round-trips through verify and rejects the wrong password', async () => {
    const hash = await hashSheetPassword('p@ss', 50)
    expect(await verifySheetPassword('p@ss', hash)).toBe(true)
    expect(await verifySheetPassword('p@sS', hash)).toBe(false)
    expect(await verifySheetPassword('', hash)).toBe(false)
  })

  it('fails closed on unknown algorithms and malformed base64', async () => {
    expect(
      await verifySheetPassword('x', {
        algorithmName: 'MD5',
        hashValue: 'AA==',
        saltValue: 'AA==',
        spinCount: 1,
      }),
    ).toBe(false)
    expect(
      await verifySheetPassword('x', {
        algorithmName: 'SHA-512',
        hashValue: '%%%',
        saltValue: 'AA==',
        spinCount: 1,
      }),
    ).toBe(false)
  })

  it('matches the ECMA-376 iteration (hash of previous digest + LE32 counter)', async () => {
    const { createHash } = await import('node:crypto')
    const salt = Buffer.from('c2FsdHNhbHRzYWx0c2Fs', 'base64')
    let digest = createHash('sha512')
      .update(Buffer.concat([salt, Buffer.from('pw', 'utf16le')]))
      .digest()
    for (let index = 0; index < 3; index += 1) {
      const counter = Buffer.alloc(4)
      counter.writeUInt32LE(index)
      digest = createHash('sha512')
        .update(Buffer.concat([digest, counter]))
        .digest()
    }
    const ours = await hashSheetPassword('pw', 3, new Uint8Array(salt))
    if ('legacy' in ours) throw new Error('expected a hashed password')
    expect(ours.hashValue).toBe(digest.toString('base64'))
  })
})

describe('sheetProtection permission mapping', () => {
  it('reads absent lock-style flags as not allowed and default-zero flags as allowed', () => {
    const parsed = parseSheetProtectionElement('<sheetProtection sheet="1"/>')
    expect(parsed.protected).toBe(true)
    expect(parsed.password).toBeNull()
    expect(parsed.allow).toEqual({
      ...DEFAULT_SHEET_PROTECTION_ALLOW,
      objects: true,
      scenarios: true,
    })
  })

  it('reads Excel attribute values into allow flags and the hash', () => {
    const parsed = parseSheetProtectionElement(
      '<sheetProtection algorithmName="SHA-512" hashValue="aGFzaA==" saltValue="c2FsdA==" ' +
        'spinCount="100000" sheet="1" objects="1" scenarios="1" formatCells="0" sort="0" ' +
        'selectLockedCells="1"/>',
    )
    expect(parsed.allow.formatCells).toBe(true)
    expect(parsed.allow.sort).toBe(true)
    expect(parsed.allow.formatRows).toBe(false)
    expect(parsed.allow.objects).toBe(false)
    expect(parsed.allow.selectLockedCells).toBe(false)
    expect(parsed.allow.selectUnlockedCells).toBe(true)
    expect(parsed.password).toEqual({
      algorithmName: 'SHA-512',
      hashValue: 'aGFzaA==',
      saltValue: 'c2FsdA==',
      spinCount: 100000,
    })
    expect(
      parseSheetProtectionElement('<sheetProtection password="CBEB" sheet="1"/>').password,
    ).toEqual({ legacy: 'CBEB' })
  })

  it('writes the dialog defaults the way Excel does and round-trips every flag', () => {
    expect(sheetProtectionElement(DEFAULT_SHEET_PROTECTION_ALLOW, null)).toBe(
      '<sheetProtection sheet="1" objects="1" scenarios="1"/>',
    )
    const allow = {
      ...DEFAULT_SHEET_PROTECTION_ALLOW,
      selectLockedCells: false,
      formatCells: true,
      insertRows: true,
      deleteColumns: true,
      autoFilter: true,
      objects: true,
    }
    const element = sheetProtectionElement(allow, { legacy: 'CBEB' })
    expect(element).toContain('password="CBEB"')
    expect(parseSheetProtectionElement(element).allow).toEqual(allow)
  })
})

describe('applySheetProtection', () => {
  it('writes flags and hash on protect and removes the element when verified', () => {
    const protectedXml = applySheetProtection(SHEET, {
      protected: true,
      allow: { ...DEFAULT_SHEET_PROTECTION_ALLOW, sort: true },
      password: {
        algorithmName: 'SHA-512',
        hashValue: 'aGFzaA==',
        saltValue: 'c2FsdA==',
        spinCount: 5,
      },
    })
    expect(protectedXml).toContain(
      '</sheetData><sheetProtection algorithmName="SHA-512" hashValue="aGFzaA==" ' +
        'saltValue="c2FsdA==" spinCount="5" sheet="1" objects="1" scenarios="1" sort="0"/>',
    )
    expect(readSheetProtection(protectedXml)?.allow.sort).toBe(true)
    expect(() => applySheetProtection(protectedXml, false)).toThrow(/password/)
    expect(applySheetProtection(protectedXml, { protected: false, verified: true })).toBe(SHEET)
  })

  it('replaces an existing element when flags are given and keeps it on a bare toggle', () => {
    const existing = SHEET.replace(
      '</sheetData>',
      '</sheetData><sheetProtection sheet="0" formatCells="0"/>',
    )
    expect(applySheetProtection(existing, true)).toContain(
      '<sheetProtection sheet="1" formatCells="0"/>',
    )
    const rewritten = applySheetProtection(existing, {
      protected: true,
      allow: DEFAULT_SHEET_PROTECTION_ALLOW,
      password: null,
    })
    expect(rewritten).toContain('<sheetProtection sheet="1" objects="1" scenarios="1"/>')
    expect(rewritten).not.toContain('formatCells')
  })
})
