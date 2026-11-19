import { unpackMtx } from './lzcomp'
import { ctfToSfnt } from './ctf'

export { MtxError } from './lzcomp'

/**
 * Inflate a MicroType-Express-compressed EOT payload (TTEMBED_TTCOMPRESSED, already
 * XOR-decrypted by the caller) into a plain sfnt, or null when the data is not decodable.
 */
export function mtxToSfnt(payload: Uint8Array): Uint8Array | null {
  try {
    return ctfToSfnt(unpackMtx(payload))
  } catch {
    return null
  }
}
