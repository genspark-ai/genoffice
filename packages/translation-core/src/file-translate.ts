/**
 * File-level translation bridge — routes .pdf / .xls / .xlsx / .pptx / .docx
 * through the LumosAI translate suite that ships under
 * `~/.lumos/bundled-skills/<hash>/translate-*`.
 *
 * Why a subprocess instead of a port:
 *   The Python handlers are mature (700+ lines) and depend on format-specific
 *   libraries we deliberately do not want in the Node bundle — `python-docx`,
 *   `python-pptx`, `openpyxl`/`xlrd`, PyMuPDF, and LibreOffice's `soffice` for
 *   the legacy .xls path. Spawning `translate.py` keeps that stack exactly as
 *   upstream maintains it and keeps the web-server bundle at ~20 MB.
 *
 *   The text-level path (`ai:translate` / `ai:translate-batch`) stays entirely
 *   in TypeScript via `@genoffice/translation-core` — only whole-file
 *   translation crosses the language boundary.
 *
 * Discovery order for the scripts directory:
 *   1. `GENOFFICE_TRANSLATE_SKILLS_DIR` (explicit override)
 *   2. `$LUMOS_HOME/bundled-skills/<hash>/translate` (newest hash wins)
 *   3. `<LUMOS_HOME|~/.lumos>/skills/translate`
 *
 * The `translate` entry script dispatches to `translate-pdf` / `-xls` / `-ppt`
 * / `-docx` by extension, so a single spawn covers every supported format.
 */
import { spawn } from 'node:child_process'
import { existsSync, mkdirSync, readdirSync, statSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, extname, join, resolve } from 'node:path'

/** Formats the upstream script dispatches on. */
export const SUPPORTED_EXTENSIONS = ['.pdf', '.xls', '.xlsx', '.pptx', '.docx'] as const
export type SupportedExtension = (typeof SUPPORTED_EXTENSIONS)[number]

export interface TranslateFileRequest {
  /** Absolute path to the source file. */
  inputPath: string
  /**
   * Where to write the translated file. Defaults to
   * `<input>_translated<ext>` next to the source.
   */
  outputPath?: string
  /**
   * Optional translation dictionary JSON (`{ "source": "target" }`). When
   * omitted the upstream handlers only reformat / re-encode and report zero
   * translations — the caller is expected to build the dictionary from the KB
   * + LLM first.
   */
  dictionaryPath?: string
  /** PDF render scale (higher = sharper background). Defaults to 2. */
  scale?: number
  /** Kill the child after this many ms. Defaults to 5 minutes. */
  timeoutMs?: number
  /**
   * Extra environment for the child. `PYTHONPATH` etc. flow through the
   * inherited env; callers rarely need to set anything here.
   */
  env?: Record<string, string>
}

export interface TranslateFileResult {
  ok: boolean
  /** Absolute path of the produced file when ok=true. */
  outputPath?: string
  /** Bytes written, when the file exists. */
  bytes?: number
  /** How long the child ran, ms. */
  elapsedMs?: number
  /** The script that handled the file, for provenance/debugging. */
  scriptPath?: string
  /** stdout from the child (trimmed). */
  stdout?: string
  /** stderr from the child (trimmed). */
  stderr?: string
  error?: string
}

export interface TranslateSkillsLocation {
  /** Directory containing `translate/scripts/translate.py`. */
  skillDir: string
  /** The unified entry script. */
  scriptPath: string
  /** Python interpreter we will spawn. */
  pythonPath: string
  /** Where the location was resolved from, for diagnostics. */
  source: 'override' | 'bundled' | 'legacy' | 'missing'
}

/** Resolve the Python interpreter the same way the upstream scripts do. */
function resolvePython(): string {
  const override = process.env.GENOFFICE_PYTHON
  if (override && existsSync(override)) return override
  for (const candidate of ['/opt/homebrew/bin/python3', '/usr/local/bin/python3', '/usr/bin/python3']) {
    if (existsSync(candidate)) return candidate
  }
  return 'python3'
}

/**
 * Locate the translate skill. Bundled skills live under a content hash
 * directory, so we pick the most recently modified one when several exist.
 */
export function resolveTranslateSkills(): TranslateSkillsLocation {
  const pythonPath = resolvePython()
  const override = process.env.GENOFFICE_TRANSLATE_SKILLS_DIR
  if (override) {
    const scriptPath = join(override, 'scripts', 'translate.py')
    if (existsSync(scriptPath)) {
      return { skillDir: override, scriptPath, pythonPath, source: 'override' }
    }
  }

  const lumosHome = process.env.LUMOS_HOME ?? join(homedir(), '.lumos')
  const bundledRoot = join(lumosHome, 'bundled-skills')
  if (existsSync(bundledRoot)) {
    let best: { dir: string; mtime: number } | null = null
    for (const entry of readdirSync(bundledRoot)) {
      const candidate = join(bundledRoot, entry, 'translate', 'scripts', 'translate.py')
      if (!existsSync(candidate)) continue
      let mtime = 0
      try {
        mtime = statSync(candidate).mtimeMs
      } catch {
        /* unreadable — skip */
      }
      if (!best || mtime > best.mtime) best = { dir: join(bundledRoot, entry, 'translate'), mtime }
    }
    if (best) {
      return {
        skillDir: best.dir,
        scriptPath: join(best.dir, 'scripts', 'translate.py'),
        pythonPath,
        source: 'bundled',
      }
    }
  }

  const legacy = join(lumosHome, 'skills', 'translate')
  const legacyScript = join(legacy, 'scripts', 'translate.py')
  if (existsSync(legacyScript)) {
    return { skillDir: legacy, scriptPath: legacyScript, pythonPath, source: 'legacy' }
  }

  return { skillDir: '', scriptPath: '', pythonPath, source: 'missing' }
}

/** True when the extension is one the upstream script knows how to handle. */
export function isSupportedExtension(pathOrExt: string): boolean {
  const ext = pathOrExt.startsWith('.') ? pathOrExt.toLowerCase() : extname(pathOrExt).toLowerCase()
  return (SUPPORTED_EXTENSIONS as readonly string[]).includes(ext)
}

/** Default output path: `<input>_translated<ext>` next to the source. */
export function defaultOutputPath(inputPath: string): string {
  const ext = extname(inputPath)
  const base = inputPath.slice(0, inputPath.length - ext.length)
  return `${base}_translated${ext}`
}

/**
 * Spawn `translate.py` for a single file. Never throws for a failed
 * translation — the result carries `ok: false` plus the child's stderr so the
 * caller can surface a real error instead of a generic one.
 */
export async function translateFile(request: TranslateFileRequest): Promise<TranslateFileResult> {
  const inputPath = resolve(request.inputPath ?? '')
  if (!request.inputPath || !existsSync(inputPath)) {
    return { ok: false, error: `translate:file input not found: ${request.inputPath ?? '(empty)'}` }
  }
  if (!isSupportedExtension(inputPath)) {
    return {
      ok: false,
      error: `translate:file unsupported extension ${extname(inputPath) || '(none)'}; expected one of ${SUPPORTED_EXTENSIONS.join(', ')}`,
    }
  }
  const location = resolveTranslateSkills()
  if (location.source === 'missing') {
    return {
      ok: false,
      error:
        'translate:file could not locate the translate skill. Set GENOFFICE_TRANSLATE_SKILLS_DIR or install the LumosAI translate suite under ~/.lumos/bundled-skills/.',
    }
  }
  const outputPath = resolve(request.outputPath ?? defaultOutputPath(inputPath))
  const outDir = dirname(outputPath)
  if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true })

  const args = [location.scriptPath, inputPath, outputPath]
  if (request.dictionaryPath) args.push('--dictionary', request.dictionaryPath)
  if (typeof request.scale === 'number') args.push('--scale', String(request.scale))

  const started = Date.now()
  const timeoutMs = request.timeoutMs ?? 5 * 60_000

  return await new Promise<TranslateFileResult>((resolvePromise) => {
    const child = spawn(location.pythonPath, args, {
      cwd: location.skillDir,
      env: { ...process.env, ...(request.env ?? {}) },
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stderr = ''
    let settled = false
    const settle = (result: TranslateFileResult): void => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolvePromise(result)
    }
    const timer = setTimeout(() => {
      try {
        child.kill('SIGKILL')
      } catch {
        /* already gone */
      }
      settle({
        ok: false,
        error: `translate:file timed out after ${timeoutMs} ms`,
        elapsedMs: Date.now() - started,
        scriptPath: location.scriptPath,
        stdout: stdout.trim().slice(-2000),
        stderr: stderr.trim().slice(-2000),
      })
    }, timeoutMs)

    child.stdout?.on('data', (chunk: Buffer) => {
      stdout += chunk.toString('utf8')
      if (stdout.length > 100_000) stdout = stdout.slice(-50_000)
    })
    child.stderr?.on('data', (chunk: Buffer) => {
      stderr += chunk.toString('utf8')
      if (stderr.length > 100_000) stderr = stderr.slice(-50_000)
    })
    child.on('error', (err) => {
      settle({
        ok: false,
        error: `translate:file failed to spawn ${location.pythonPath}: ${err.message}`,
        elapsedMs: Date.now() - started,
        scriptPath: location.scriptPath,
      })
    })
    child.on('close', (code) => {
      const elapsedMs = Date.now() - started
      if (code !== 0) {
        settle({
          ok: false,
          error: `translate:file exited with code ${code ?? 'null'}`,
          elapsedMs,
          scriptPath: location.scriptPath,
          stdout: stdout.trim().slice(-2000),
          stderr: stderr.trim().slice(-2000),
        })
        return
      }
      if (!existsSync(outputPath)) {
        settle({
          ok: false,
          error: 'translate:file finished but produced no output file',
          elapsedMs,
          scriptPath: location.scriptPath,
          stdout: stdout.trim().slice(-2000),
          stderr: stderr.trim().slice(-2000),
        })
        return
      }
      let bytes = 0
      try {
        bytes = statSync(outputPath).size
      } catch {
        /* ignore */
      }
      settle({
        ok: true,
        outputPath,
        bytes,
        elapsedMs,
        scriptPath: location.scriptPath,
        stdout: stdout.trim().slice(-2000),
        stderr: stderr.trim().slice(-2000),
      })
    })
  })
}
