import { describe, it, expect } from 'vitest'
import { cfbKind, isCfbHeader } from '../src/main/cfb-sniff'

const CFB = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])

describe('cfbKind', () => {
  it('plain zip (pptx) does not match', () => {
    expect(isCfbHeader(Buffer.from('PK\x03\x04rest'))).toBe(false)
    expect(cfbKind(Buffer.from('PK\x03\x04rest'))).toBeNull()
  })

  it('CFB without EncryptedPackage stream -> legacy .ppt', () => {
    const buf = Buffer.concat([CFB, Buffer.alloc(512), Buffer.from('PowerPoint Document', 'utf16le')])
    expect(cfbKind(buf)).toBe('legacy')
  })

  it('CFB containing EncryptedPackage stream name (UTF-16LE) -> encrypted OOXML', () => {
    const buf = Buffer.concat([CFB, Buffer.alloc(64), Buffer.from('EncryptedPackage', 'utf16le'), Buffer.alloc(64)])
    expect(cfbKind(buf)).toBe('encrypted')
  })

  it('ASCII-form EncryptedPackage is not misdetected (stream name must be UTF-16LE)', () => {
    const buf = Buffer.concat([CFB, Buffer.from('EncryptedPackage', 'ascii')])
    expect(cfbKind(buf)).toBe('legacy')
  })
})
