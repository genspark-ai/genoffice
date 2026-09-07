/** Cap on document / selection plain text returned to the model / UI. */
export const PLAIN_TEXT_MAX_CHARS = 80_000
/** Short selection preview attached to every user turn. */
export const SELECTION_PREVIEW_CHARS = 400
export const PLAIN_TEXT_UNAVAILABLE = 'plain text unavailable'

export interface HangulStudioFacade {
  pageCount(): Promise<number>
  getPlainText(): Promise<string>
  getSelectionText(): Promise<string | null>
  hasSelection(): Promise<boolean>
}

export interface StudioTextSource {
  pageCount(): Promise<number>
  exportHml(): Promise<Uint8Array>
  getHmlSaveState?(): Promise<{ hmlSavable: boolean }>
  getSelectionContext(): Promise<{ collapsed: boolean; selectedTextSha256: string | null }>
  hwpctrl: { call(method: string, args?: unknown[]): Promise<unknown> }
}

export function clipPlainText(text: string, max = PLAIN_TEXT_MAX_CHARS): string {
  if (text.length <= max) return text
  return `${text.slice(0, max)}\n\n[truncated]`
}

function decodeCodePoint(value: number): string {
  if (!Number.isInteger(value) || value < 0 || value > 0x10ffff) return ''
  if (value >= 0xd800 && value <= 0xdfff) return ''
  return String.fromCodePoint(value)
}

export function stripHmlToPlainText(xml: string): string {
  return xml
    .replace(/<\?xml[\s\S]*?\?>/g, '')
    .replace(/<!(?:DOCTYPE|-- )[\s\S]*?>/g, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&apos;/gi, "'")
    .replace(/&#(\d+);/g, (_, n: string) => decodeCodePoint(Number(n)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n: string) => decodeCodePoint(parseInt(n, 16)))
    .replace(/&amp;/gi, '&')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim()
}

function asText(value: unknown): string | null {
  if (typeof value === 'string' && value.length > 0) return value
  if (value && typeof value === 'object' && 'result' in value) {
    const result = (value as { result: unknown }).result
    if (typeof result === 'string' && result.length > 0) return result
  }
  return null
}

async function tryHwpctrl(
  studio: StudioTextSource,
  method: string,
  args?: unknown[],
): Promise<string | null> {
  try {
    return asText(await studio.hwpctrl.call(method, args))
  } catch {
    return null
  }
}

async function plainTextFromHml(studio: StudioTextSource): Promise<string | null> {
  if (studio.getHmlSaveState) {
    try {
      const state = await studio.getHmlSaveState()
      if (!state.hmlSavable) return null
    } catch {
      /* still try exportHml — some hosts omit the state API */
    }
  }
  try {
    return stripHmlToPlainText(new TextDecoder('utf-8').decode(await studio.exportHml()))
  } catch {
    return null
  }
}

export async function getPlainText(studio: StudioTextSource): Promise<string> {
  const fromCtrl = await tryHwpctrl(studio, 'GetTextFile', ['TEXT', ''])
  if (fromCtrl) return clipPlainText(fromCtrl)
  const fromHml = await plainTextFromHml(studio)
  if (fromHml !== null) return clipPlainText(fromHml)
  throw new Error(PLAIN_TEXT_UNAVAILABLE)
}

export async function getSelectionText(studio: StudioTextSource): Promise<string | null> {
  const fromCtrl = await tryHwpctrl(studio, 'GetTextFile', ['TEXT', 'saveblock'])
  return fromCtrl ? clipPlainText(fromCtrl) : null
}

export async function hasSelection(studio: StudioTextSource): Promise<boolean> {
  try {
    const sel = await studio.getSelectionContext()
    return !sel.collapsed && Boolean(sel.selectedTextSha256)
  } catch {
    return false
  }
}

export function fileNameOf(path: string | null): string {
  if (!path) return 'untitled.hwp'
  const parts = path.split(/[\\/]/)
  return parts[parts.length - 1] ?? path
}

export function createStudioFacade(studio: StudioTextSource): HangulStudioFacade {
  return {
    pageCount: () => studio.pageCount(),
    getPlainText: () => getPlainText(studio),
    getSelectionText: () => getSelectionText(studio),
    hasSelection: () => hasSelection(studio),
  }
}
