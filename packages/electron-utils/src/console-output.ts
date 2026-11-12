import { execFileSync } from 'node:child_process'
import { TextDecoder } from 'node:util'

/**
 * Decode output captured from a child process.
 *
 * Windows console programs (reg.exe, powershell.exe) write the active OEM code
 * page, not UTF-8. Reading those bytes as UTF-8 turns localised text (registry
 * values, certificate subjects, paths) into U+FFFD replacement characters.
 * The code page is read once from `HKLM\...\Nls\CodePage` (OEMCP) and mapped to
 * a TextDecoder label, so a Shift-JIS, EUC-KR, Big5 or Western-European system
 * decodes correctly too. An unknown page falls back to UTF-8 with replacement
 * characters, which keeps the callers' "unreadable name" fallbacks working
 * instead of producing confident garbage.
 *
 * PowerShell can also emit UTF-16 when redirected, so honour BOMs.
 * Non-Windows output stays UTF-8.
 */

/** WHATWG TextDecoder labels for the OEM code pages Windows reports. 437/850
 *  have no label of their own; windows-1252 is the closest Western mapping. */
const OEM_DECODER_LABELS: Readonly<Record<number, string>> = {
  437: 'windows-1252',
  850: 'windows-1252',
  852: 'iso-8859-2',
  855: 'iso-8859-5',
  857: 'iso-8859-9',
  858: 'windows-1252',
  860: 'iso-8859-1',
  861: 'iso-8859-1',
  862: 'iso-8859-8',
  863: 'iso-8859-1',
  865: 'iso-8859-1',
  866: 'ibm866',
  869: 'iso-8859-7',
  874: 'windows-874',
  932: 'shift_jis',
  936: 'gbk',
  949: 'euc-kr',
  950: 'big5',
  1250: 'windows-1250',
  1251: 'windows-1251',
  1252: 'windows-1252',
  1253: 'windows-1253',
  1254: 'windows-1254',
  1255: 'windows-1255',
  1256: 'windows-1256',
  1257: 'windows-1257',
  1258: 'windows-1258',
  10000: 'macintosh',
  65001: 'utf-8',
}

let oemCodePage: number | undefined
let oemCodePageRead = false

/** The active OEM code page on Windows, read once; undefined when unknown. */
export function windowsOemCodePage(): number | undefined {
  if (oemCodePageRead) return oemCodePage
  oemCodePageRead = true
  if (process.platform !== 'win32') return undefined
  try {
    // The OEMCP value is ASCII, so the default decoding is safe here.
    const out = execFileSync(
      'reg',
      ['query', 'HKLM\\SYSTEM\\CurrentControlSet\\Control\\Nls\\CodePage', '/v', 'OEMCP'],
      { encoding: 'ascii', windowsHide: true, timeout: 5000 },
    )
    const match = /OEMCP\s+REG_SZ\s+(\d+)/i.exec(out)
    if (match) oemCodePage = Number(match[1])
  } catch {
    oemCodePage = undefined
  }
  return oemCodePage
}

export interface DecodeCommandOptions {
  /** Override the OEM code page (tests, or a caller that already knows it). */
  codePage?: number
}

export function decodeCommandOutput(
  output: Buffer | string,
  platform: NodeJS.Platform = process.platform,
  options: DecodeCommandOptions = {},
): string {
  if (typeof output === 'string') return output
  if (output.length >= 2 && output[0] === 0xff && output[1] === 0xfe) {
    return new TextDecoder('utf-16le').decode(output).replace(/^\uFEFF/, '')
  }
  if (output.length >= 2 && output[0] === 0xfe && output[1] === 0xff) {
    return new TextDecoder('utf-16be').decode(output).replace(/^\uFEFF/, '')
  }
  if (platform !== 'win32') return output.toString('utf8')
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(output)
  } catch {
    const label = OEM_DECODER_LABELS[options.codePage ?? windowsOemCodePage() ?? -1]
    // No label for this page: UTF-8 with replacement characters, never a guess.
    return new TextDecoder(label ?? 'utf-8').decode(output)
  }
}

/** UTF-8 for children known to emit it (the genoffice CLI; PowerShell after
 *  `[Console]::OutputEncoding = UTF8`). Non-fatal, so a truncated tail at the
 *  byte cap or on timeout becomes U+FFFD instead of failing the whole buffer. */
export function decodeUtf8(output: Buffer | string): string {
  return typeof output === 'string' ? output : output.toString('utf8')
}

/** Join stream chunks BEFORE decoding. A multi-byte character can straddle two
 *  'data' events, so decoding each chunk on its own silently corrupts it. */
export function decodeChunks(
  chunks: readonly (Buffer | string)[],
  platform: NodeJS.Platform = process.platform,
  options: DecodeCommandOptions = {},
): string {
  if (chunks.length === 0) return ''
  if (chunks.every((c) => typeof c === 'string')) return (chunks as string[]).join('')
  const buf = Buffer.concat(chunks.map((c) => (typeof c === 'string' ? Buffer.from(c, 'utf8') : c)))
  return decodeCommandOutput(buf, platform, options)
}

/** Same join-then-decode rule, for children that always write UTF-8. */
export function decodeChunksUtf8(chunks: readonly (Buffer | string)[]): string {
  if (chunks.length === 0) return ''
  if (chunks.every((c) => typeof c === 'string')) return (chunks as string[]).join('')
  const buf = Buffer.concat(chunks.map((c) => (typeof c === 'string' ? Buffer.from(c, 'utf8') : c)))
  return buf.toString('utf8')
}
