/**
 * @genoffice/odf2ooxml — local ODF import filters: .odt → .docx and
 * .odp → .pptx. Pure-function package: bytes in, bytes out; nothing here
 * touches Electron or the filesystem. Both filters follow the .ods import
 * rule — a construct that cannot be converted provably-correct is counted in
 * `dropped` instead of approximated into wrong markup.
 */

export { convertOdtToDocx, OdfConvertError, type OdtToDocxResult } from './odt-docx'
export { convertOdpToPptx, type OdpToPptxResult } from './odp-pptx'
