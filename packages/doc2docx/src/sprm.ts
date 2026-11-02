import { LegacyDocError } from './errors'

/**
 * sprm (single property modifier) decoding, MS-DOC 2.6.1. A grpprl is a
 * sequence of sprms: a 2-byte opcode followed by an operand whose size comes
 * from the 3-bit spra class. Unknown sprms are skipped by that size rule, so
 * new formatting can pass through without breaking the reader.
 */

/** sprmTDefTable / sprmPChgTabs: the length byte counts itself (see sprmOperandSize) */
export const SPRM_LONG_TABLE = 0xd608
export const SPRM_LONG_PARAGRAPH = 0xc615

export interface Sprm {
  opcode: number
  operand: Uint8Array
}

function badSprm(detail: string): never {
  throw new LegacyDocError('bad_sprm', detail)
}

function isLongLengthOpcode(opcode: number): boolean {
  return opcode === SPRM_LONG_TABLE || opcode === SPRM_LONG_PARAGRAPH
}

/**
 * Payload size (bytes after the length byte(s)) of the sprm whose 2-byte
 * opcode starts at `offset`: spra 0/1 -> 1, 2/4/5 -> 2, 3 -> 4, 7 -> 3;
 * spra 6 -> the length byte (u8), except sprmTDefTable/sprmPChgTabs, which use
 * a u16 length that counts one of its own bytes, so the payload is one byte
 * shorter than the declared value (MS-DOC 2.6.1; matches POI's SprmOperation,
 * whose total size is declared + 3).
 */
export function sprmOperandSize(grpprl: Uint8Array, offset: number): number {
  const opcode = grpprl[offset]! | (grpprl[offset + 1]! << 8)
  const spra = (opcode >> 13) & 7
  switch (spra) {
    case 0:
    case 1:
      return 1
    case 2:
    case 4:
    case 5:
      return 2
    case 3:
      return 4
    case 7:
      return 3
    case 6: {
      const lengthBytes = isLongLengthOpcode(opcode) ? 2 : 1
      if (offset + 2 + lengthBytes > grpprl.length)
        badSprm(`sprm 0x${opcode.toString(16)} has no operand length byte`)
      if (lengthBytes === 2) {
        const declared = grpprl[offset + 2]! | (grpprl[offset + 3]! << 8)
        if (declared < 1)
          badSprm(`sprm 0x${opcode.toString(16)} declares an empty u16 operand length`)
        return declared - 1
      }
      return grpprl[offset + 2]!
    }
    default:
      return 0
  }
}

/** Decode a grpprl into sprms; a truncated opcode or operand is a 'bad_sprm' error. */
export function iterateSprms(grpprl: Uint8Array): Sprm[] {
  const sprms: Sprm[] = []
  let offset = 0
  while (offset < grpprl.length) {
    if (offset + 2 > grpprl.length)
      badSprm(
        `the last ${grpprl.length - offset} byte(s) of a grpprl do not hold a 2-byte sprm opcode`,
      )
    const opcode = grpprl[offset]! | (grpprl[offset + 1]! << 8)
    const size = sprmOperandSize(grpprl, offset)
    const spra = (opcode >> 13) & 7
    const headerBytes = spra === 6 ? (isLongLengthOpcode(opcode) ? 4 : 3) : 2
    if (offset + headerBytes + size > grpprl.length)
      badSprm(
        `sprm 0x${opcode.toString(16)} needs ${size} operand byte(s) but the grpprl ends after ${grpprl.length - offset - headerBytes}`,
      )
    sprms.push({
      opcode,
      operand: grpprl.slice(offset + headerBytes, offset + headerBytes + size),
    })
    offset += headerBytes + size
  }
  return sprms
}

/** operand readers; undefined when the operand is too short for the field (unknown sprm) */
export function readSprmU8(operand: Uint8Array, offset = 0): number | undefined {
  return offset + 1 <= operand.length ? operand[offset] : undefined
}

export function readSprmU16(operand: Uint8Array, offset = 0): number | undefined {
  return offset + 2 <= operand.length ? operand[offset]! | (operand[offset + 1]! << 8) : undefined
}

export function readSprmI16(operand: Uint8Array, offset = 0): number | undefined {
  const value = readSprmU16(operand, offset)
  return value === undefined ? undefined : (value << 16) >> 16
}

export function readSprmU32(operand: Uint8Array, offset = 0): number | undefined {
  return offset + 4 <= operand.length
    ? (operand[offset]! |
        (operand[offset + 1]! << 8) |
        (operand[offset + 2]! << 16) |
        (operand[offset + 3]! << 24)) >>>
        0
    : undefined
}
