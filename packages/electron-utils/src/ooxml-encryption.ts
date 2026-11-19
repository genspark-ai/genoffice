/** Shared ECMA-376 encrypted OOXML helpers for Electron main processes. */
import officeCrypto from 'officecrypto-tool'

const CFB_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
const ENCRYPTED_STREAM_UTF16 = Buffer.from('EncryptedPackage', 'utf16le')

export function isCfbFile(bytes: Buffer): boolean {
  return bytes.length >= CFB_MAGIC.length && bytes.subarray(0, 8).equals(CFB_MAGIC)
}

/** Encrypted Office documents wrap the OOXML ZIP package in a CFB container. */
export function isEncryptedOoxml(bytes: Buffer): boolean {
  return isCfbFile(bytes) && bytes.includes(ENCRYPTED_STREAM_UTF16)
}

export type OoxmlDecryptFailReason = 'wrong-password' | 'unsupported'

export class OoxmlDecryptError extends Error {
  readonly reason: OoxmlDecryptFailReason

  constructor(reason: OoxmlDecryptFailReason, message: string) {
    super(message)
    this.name = 'OoxmlDecryptError'
    this.reason = reason
  }
}

export async function decryptOoxml(bytes: Buffer, password: string): Promise<Buffer> {
  try {
    return await officeCrypto.decrypt(bytes, { password })
  } catch (error) {
    const message = String((error as Error)?.message ?? error)
    if (message.includes('password is incorrect')) {
      throw new OoxmlDecryptError('wrong-password', message)
    }
    throw new OoxmlDecryptError('unsupported', message)
  }
}

/** Encrypt plain OOXML bytes using Office Agile encryption. */
export function encryptOoxml(bytes: Buffer, password: string): Buffer {
  return officeCrypto.encrypt(bytes, { password })
}
