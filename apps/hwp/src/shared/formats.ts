/** Extensions rhwp opens. Keep in sync with shell routing and drop-open. */
export const HWP_EXTENSIONS = ['hwp', 'hwpx', 'hml'] as const

export type HwpExtension = (typeof HWP_EXTENSIONS)[number]

export const HWP_RE = /\.(hwp|hwpx|hml)$/i

export function isHwpPath(path: string): boolean {
  return HWP_RE.test(path)
}
