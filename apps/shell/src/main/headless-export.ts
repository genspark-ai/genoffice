/**
 * Headless export host: the single entry every dialog-free export path can be
 * driven through without a visible editor window.
 *
 *   <app binary> --headless-export <input> --to <format> --out <path> [--json]
 *
 * Routing is by input extension; each editor module owns the hidden-window
 * export of its own format and reuses the very same renderer pipeline the
 * File menu uses, so GUI and CLI output cannot drift.
 */
import { existsSync, statSync } from 'node:fs'
import { dirname, resolve } from 'node:path'

import {
  HEADLESS_EXIT,
  HEADLESS_SUPPORTED_EXTENSIONS,
  HEADLESS_TARGETS,
  headlessModuleFor,
  type HeadlessExportFormat,
  type HeadlessExportModule,
  type HeadlessExportOutcome,
  type HeadlessExportRequest,
} from '@genoffice/electron-utils'

type HeadlessFailure = Extract<HeadlessExportOutcome, { ok: false }>

/** Per-module hidden-window exporters, injected so this file stays testable. */
export type HeadlessExporters = Record<
  HeadlessExportModule,
  (input: string, outPath: string, format: HeadlessExportFormat) => Promise<void>
>

/** Disk probes, injected so the host is testable without touching the filesystem. */
export interface HeadlessFs {
  exists(path: string): boolean
  isFile(path: string): boolean
}

const realFs: HeadlessFs = {
  exists: existsSync,
  isFile: (path) => statSync(path).isFile(),
}

/** Checks a caller-supplied path pair before any window is created. */
export function validateHeadlessPaths(
  request: HeadlessExportRequest,
  fs: HeadlessFs = realFs,
): { ok: true; input: string; outPath: string; module: HeadlessExportModule } | HeadlessFailure {
  const input = resolve(request.input)
  if (!fs.exists(input)) {
    return { ok: false, code: HEADLESS_EXIT.inputError, message: `input file not found: ${input}` }
  }
  if (!fs.isFile(input)) {
    return { ok: false, code: HEADLESS_EXIT.inputError, message: `input is not a file: ${input}` }
  }
  const module = headlessModuleFor(input)
  if (!module) {
    return {
      ok: false,
      code: HEADLESS_EXIT.inputError,
      message: `cannot export ${input} (supported inputs: ${HEADLESS_SUPPORTED_EXTENSIONS})`,
    }
  }
  if (!HEADLESS_TARGETS[module].includes(request.targetFormat)) {
    return {
      ok: false,
      code: HEADLESS_EXIT.badArgs,
      message: `cannot export ${input} to ${request.targetFormat} (this input supports: ${HEADLESS_TARGETS[module].join(', ')})`,
    }
  }
  const outPath = resolve(request.outPath)
  if (!fs.exists(dirname(outPath))) {
    return {
      ok: false,
      code: HEADLESS_EXIT.badArgs,
      message: `output directory does not exist: ${dirname(outPath)}`,
    }
  }
  return { ok: true, input, outPath, module }
}

/**
 * Runs one export end to end. Returns the outcome instead of throwing so the
 * caller can render the envelope and pick the exit code in one place.
 */
export async function runHeadlessExport(
  request: HeadlessExportRequest,
  exporters: HeadlessExporters,
  fs: HeadlessFs = realFs,
): Promise<HeadlessExportOutcome> {
  const checked = validateHeadlessPaths(request, fs)
  if (!checked.ok) return checked
  try {
    await exporters[checked.module](checked.input, checked.outPath, request.targetFormat)
  } catch (err) {
    return {
      ok: false,
      code: HEADLESS_EXIT.conversionFailure,
      message: err instanceof Error ? err.message : String(err),
    }
  }
  if (!fs.exists(checked.outPath)) {
    return {
      ok: false,
      code: HEADLESS_EXIT.conversionFailure,
      message: `export reported success but wrote no file at ${checked.outPath}`,
    }
  }
  return { ok: true, input: checked.input, outPath: checked.outPath }
}
