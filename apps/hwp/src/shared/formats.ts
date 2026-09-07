/** Extensions rhwp opens. Keep in sync with shell routing and drop-open. */
export const HWP_EXTENSIONS = ['hwp', 'hwpx', 'hml'] as const

export type HwpExtension = (typeof HWP_EXTENSIONS)[number]

export const HWP_RE = /\.(hwp|hwpx|hml)$/i

export function isHwpPath(path: string): boolean {
  return HWP_RE.test(path)
}

export type HwpSaveFormat = 'hwp' | 'hwpx' | 'hml'

/** New documents start as HWP; Save As may pick HWPX/HML. */
export function saveFormatForPath(path: string): HwpSaveFormat {
  if (/\.hwpx$/i.test(path)) return 'hwpx'
  if (/\.hml$/i.test(path)) return 'hml'
  return 'hwp'
}

/** Dialog paths without an extension become `.hwp`. */
export function ensureHwpSavePath(path: string): string {
  return HWP_RE.test(path) ? path : `${path}.hwp`
}

export const HML_UNAVAILABLE = 'hwp: HML export unavailable'

export function bytesForSaveFormat(
  format: HwpSaveFormat,
  payload: { hwp: Uint8Array; hwpx: Uint8Array; hml?: Uint8Array },
): Uint8Array {
  if (format === 'hwpx') return payload.hwpx
  if (format === 'hml') {
    if (!payload.hml?.byteLength) throw new Error(HML_UNAVAILABLE)
    return payload.hml
  }
  return payload.hwp
}
