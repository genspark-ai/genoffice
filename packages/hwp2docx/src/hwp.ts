/**
 * .hwp (HWP 5.x / compressed HWP 3) → .docx. The binary format is read by
 * rhwp (MIT, https://github.com/edwardkim/rhwp), which serializes its document
 * model as HWPX; the HWPX filter then builds the .docx, so both formats share
 * one conversion path and one set of fidelity rules. rhwp's per-export loss
 * report is a debugging aid only — the conversion always carries everything
 * rhwp read.
 */

import { readFile } from 'node:fs/promises'
import { createRequire } from 'node:module'
import { join } from 'node:path'
import { convertHwpxToDocx, type HwpxToDocxResult } from './hwpx-docx'
import { HwpConvertError, isHwpBinary } from './package'

type Rhwp = typeof import('@rhwp/core')

let loading: Promise<Rhwp> | null = null

/**
 * node_modules in dev, tests and the CLI; the packaged app ships no
 * node_modules (everything is bundled), so electron-builder copies the module
 * to Resources/wasm (apps/shell/electron-builder.cjs, like pdfium.wasm).
 */
function rhwpWasmPath(): string {
  try {
    return createRequire(import.meta.url).resolve('@rhwp/core/rhwp_bg.wasm')
  } catch {
    const resources = (process as { resourcesPath?: string }).resourcesPath
    if (resources) return join(resources, 'wasm', 'rhwp_bg.wasm')
    throw new HwpConvertError('unreadable')
  }
}

/** rhwp's WebAssembly module, compiled once per process */
function loadRhwp(): Promise<Rhwp> {
  loading ??= (async () => {
    const mod = await import('@rhwp/core')
    const wasm = await readFile(rhwpWasmPath())
    mod.initSync({ module: wasm })
    return mod
  })()
  loading.catch(() => {
    loading = null
  })
  return loading
}

function isZip(bytes: Uint8Array): boolean {
  return bytes[0] === 0x50 && bytes[1] === 0x4b
}

/** HWP 3.x files open with this ASCII signature */
function isHwp3(bytes: Uint8Array): boolean {
  return new TextDecoder('latin1').decode(bytes.subarray(0, 17)) === 'HWP Document File'
}

/** HWP binary → HWPX bytes through rhwp */
export async function hwpToHwpx(hwp: Uint8Array): Promise<Uint8Array> {
  const rhwp = await loadRhwp()
  let document: InstanceType<Rhwp['HwpDocument']> | null = null
  try {
    document = new rhwp.HwpDocument(hwp)
    return document.exportHwpx()
  } catch {
    throw new HwpConvertError('unreadable')
  } finally {
    document?.free()
  }
}

/**
 * Any Hancom word-processor file → .docx, routed by content rather than by
 * extension: a zip is HWPX (some sites serve it as .hwp), an OLE compound file
 * or an HWP 3 signature goes through rhwp first (others serve HWP 5 as .hwpx).
 */
export async function convertHwpToDocx(bytes: Uint8Array): Promise<HwpxToDocxResult> {
  if (isZip(bytes)) return convertHwpxToDocx(bytes)
  if (!isHwpBinary(bytes) && !isHwp3(bytes)) throw new HwpConvertError('unreadable')
  const result = await convertHwpxToDocx(await hwpToHwpx(bytes))
  return { ...result, via: 'rhwp' }
}
