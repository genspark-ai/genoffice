/**
 * @genoffice/hwp2docx — local Hancom word-processor import filter:
 * .hwpx (OWPML, KS X 6101) → .docx, with binary .hwp read through rhwp into
 * HWPX first. Pure-function package: bytes in, bytes out; nothing
 * here touches Electron or the filesystem. Like the ODF filters, a construct
 * that cannot be converted provably-correct is counted in `dropped` instead of
 * approximated into wrong markup.
 */

export { convertHwpxToDocx, HwpConvertError, type HwpxToDocxResult } from './hwpx-docx'
export { isHwpBinary, MAX_HWPX_PART_BYTES, type HwpConvertErrorCode } from './package'
export { convertHwpToDocx, hwpToHwpx } from './hwp'
export { bmpToPng } from './bmp'
export { classifyFont, type FontClass, type FontHint } from './fonts'
export {
  convertHwpToTiptap,
  docxToTiptap,
  type HwpToTiptapResult,
  type TiptapMark,
  type TiptapNode,
} from './tiptap'
