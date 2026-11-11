import { LegacyDocError } from './errors'

/** A PLCF BTE: the FC boundaries and FKP page numbers of one formatting table. */
export interface PlcBte {
  /** n+1 FC boundaries of the FKPs (FC of the first character each page covers) */
  fcs: number[]
  /** n FKP page numbers, already masked to the 22-bit page field */
  pages: number[]
}

/** page number of an FKP: 22 bits of the PnFkp entry (MS-DOC 2.8.1) */
const PAGE_MASK = 0x3fffff
const BTE_BYTES = 8
const FC_BYTES = 4

/**
 * Parse a PlcBteChpx / PlcBtePapx (MS-DOC 2.8.1): n+1 FCs (u32) followed by n
 * PnFkp entries (u32; the low 22 bits are the FKP page number). A zero-length
 * table is valid (documents without tracked formatting tables).
 */
export function parsePlcBte(table: Uint8Array, fc: number, lcb: number): PlcBte {
  if (lcb === 0) return { fcs: [], pages: [] }
  if (lcb < BTE_BYTES || (lcb - 4) % BTE_BYTES !== 0)
    throw new LegacyDocError(
      'bad_fkp',
      `a formatting table has an invalid length (lcb ${lcb}, expected 8n+4)`,
    )
  if (fc > table.length || lcb > table.length - fc)
    throw new LegacyDocError(
      'bad_fkp',
      `a formatting table range (${fc}+${lcb}) lies outside the ${table.length}-byte table stream`,
    )
  const view = new DataView(table.buffer, table.byteOffset, table.byteLength)
  const count = (lcb - 4) / BTE_BYTES
  const fcs: number[] = []
  for (let index = 0; index <= count; index++) fcs.push(view.getUint32(fc + index * FC_BYTES, true))
  const pages: number[] = []
  const pagesStart = fc + (count + 1) * FC_BYTES
  for (let index = 0; index < count; index++)
    pages.push(view.getUint32(pagesStart + index * FC_BYTES, true) & PAGE_MASK)
  return { fcs, pages }
}
