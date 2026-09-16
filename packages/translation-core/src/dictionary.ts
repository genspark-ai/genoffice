/**
 * Dictionary builder — KB + LLM in one pass.
 *
 * The upstream `translate.py` handlers consume a flat
 * `{ "source": "target" }` JSON dictionary and apply it longest-match-first.
 * Nothing generates that dictionary for the user, so today they have to hand
 * write it before a file translation is worth anything. This module closes
 * that gap:
 *
 *   1. extract every text segment from the file (`@genoffice/file-parse`)
 *   2. seed the dictionary with the knowledge base's mandatory terms
 *   3. send the remaining segments to the active provider in batches
 *   4. post-process the model output against the KB
 *        - forbidden text is rewritten to its replacement
 *        - brand `neverTranslate` words are restored verbatim
 *        - mandatory terms override whatever the model produced
 *   5. write `{ source: target }` to a JSON file the caller passes as
 *      `--dictionary`
 *
 * Segmentation is line-oriented because that is what the format handlers do:
 * a DOCX run, a PPTX run and an XLSX cell are all short, single-line strings,
 * so a line-split dictionary matches them exactly and still works as the
 * substring fallback.
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, extname, join } from 'node:path'

import { parseFileToText } from '@genoffice/file-parse'
import {
  KnowledgeBase,
  extractTranslationText,
  type TranslateBatchUnitResult,
} from '@genoffice/translation-core'

/** A segment worth putting in a dictionary. */
export interface DictionarySegment {
  /** The original text, exactly as it appears in the file. */
  source: string
  /** Where it came from, for provenance in the UI. */
  origin: 'kb' | 'llm'
  target?: string
}

export interface BuildDictionaryRequest {
  inputPath: string
  sourceLang: string
  targetLang: string
  /** Where to write the JSON. Defaults to `<DATA_DIR>/translation-dictionaries/<name>.json`. */
  outputPath?: string
  /** Max segments sent to the model. Defaults to 400. */
  maxSegments?: number
  /** Skip segments shorter than this many characters. Defaults to 2. */
  minChars?: number
  /** Optional customer name, forwarded to the KB resolver. */
  customerName?: string
  /** Optional glossary bucket, forwarded to the KB resolver. */
  glossaryCategory?: string
  /** When false, only the KB contributes and no model call is made. */
  useLlm?: boolean
  /** Data directory for the default output path. */
  dataDir: string
}

export interface BuildDictionaryResult {
  ok: boolean
  /** Absolute path of the written dictionary. */
  dictionaryPath?: string
  /** How many segments came from the KB (mandatory terms). */
  kbEntries?: number
  /** How many segments the model translated. */
  llmEntries?: number
  /** Segments we could not translate (model failure or empty response). */
  missed?: string[]
  /** Total segments considered. */
  totalSegments?: number
  /** How long the whole build took, ms. */
  elapsedMs?: number
  /** All dictionary entries in insertion order, with KB/ LLM provenance for the UI. */
  segments?: DictionarySegment[]
  error?: string
}

/**
 * Split extracted file text into the line-oriented segments the format
 * handlers will actually match against.
 *
 * Rules:
 *  - split on newlines (the extractors already emit one segment per run/cell)
 *  - collapse internal whitespace so a dictionary key matches the run text
 *  - drop blanks and anything below `minChars` (punctuation-only runs)
 *  - dedupe, preserving first-seen order so the earlier context wins
 */
export function mineSegments(text: string, minChars = 2): string[] {
  const seen = new Set<string>()
  const out: string[] = []
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.replace(/\s+/g, ' ').trim()
    if (line.length < minChars) continue
    // Pure numbers / punctuation do not need a translation and would only
    // add noise to the model call.
    if (!/[\p{L}]/u.test(line)) continue
    if (seen.has(line)) continue
    seen.add(line)
    out.push(line)
  }
  return out
}

/** Split `segments` into batches that stay under `maxChars` each. */
export function batchSegments(segments: string[], maxChars = 3000): string[][] {
  const batches: string[][] = []
  let current: string[] = []
  let size = 0
  for (const segment of segments) {
    if (current.length > 0 && size + segment.length > maxChars) {
      batches.push(current)
      current = []
      size = 0
    }
    current.push(segment)
    size += segment.length + 1
  }
  if (current.length > 0) batches.push(current)
  return batches
}

/** Default on-disk location for generated dictionaries. */
export function defaultDictionaryPath(dataDir: string, inputPath: string): string {
  const ext = extname(inputPath)
  const base = inputPath.slice(0, inputPath.length - ext.length).split('/').pop() ?? 'dictionary'
  const safe = base.replace(/[^a-zA-Z0-9._-]/g, '_')
  const stamp = new Date().toISOString().replace(/[:.]/g, '-')
  return join(dataDir, 'translation-dictionaries', `${safe}-${stamp}.json`)
}

/**
 * Apply the KB to a model translation.
 *
 *  - forbidden entries: rewrite the banned phrasing to its replacement
 *  - brand `neverTranslate`: if the model translated the word away, restore it
 *  - mandatory terms: force the target term in place of whatever came back
 *
 * Exported for tests — the ordering here is the part most likely to regress.
 */
export function applyKbRules(
  translated: string,
  kb: KnowledgeBase,
  opts: { sourceLang: string; targetLang: string; customerName?: string },
): { text: string; matchedTerms: string[] } {
  const resolved = kb.resolve({
    sourceLang: opts.sourceLang,
    targetLang: opts.targetLang,
    ...(opts.customerName !== undefined ? { customerName: opts.customerName } : {}),
  })
  let text = translated
  const matchedTerms: string[] = []

  for (const forbidden of resolved.forbidden) {
    if (!forbidden.forbiddenText) continue
    if (!text.includes(forbidden.forbiddenText)) continue
    text = forbidden.replacement
      ? text.split(forbidden.forbiddenText).join(forbidden.replacement)
      : text.split(forbidden.forbiddenText).join('')
  }

  for (const brand of resolved.brands) {
    if (brand.policy === 'neverTranslate' && brand.word && !text.includes(brand.word)) {
      // The model translated the brand away; append the original so the
      // never-translate rule is respected downstream.
      text = text.includes(brand.word) ? text : `${text} (${brand.word})`
    } else if (brand.policy === 'translateAs' && brand.word && brand.translateAs) {
      if (text.includes(brand.word) && !text.includes(brand.translateAs)) {
        text = text.split(brand.word).join(brand.translateAs)
      }
    }
  }

  for (const term of resolved.terms) {
    if (!term.sourceTerm || !term.targetTerm) continue
    if (!text.includes(term.sourceTerm)) continue
    text = text.split(term.sourceTerm).join(term.targetTerm)
    matchedTerms.push(term.sourceTerm)
  }

  return { text, matchedTerms }
}

export interface BuildDictionaryDeps {
  /** Injected so tests do not need a live provider. */
  translateBatch: (input: {
    units: Array<{ unitId: string; kind: 'paragraph'; sourceText: string; order: number }>
    sourceLang: string
    targetLang: string
    scene: string
    glossaryCategory?: string
    customerName?: string
  }) => Promise<{ ok: boolean; units?: TranslateBatchUnitResult[]; error?: string }>
}

/**
 * Build the dictionary. Never throws — failures land in `result.error` so the
 * IPC handler can forward them verbatim.
 */
export async function buildDictionary(
  request: BuildDictionaryRequest,
  deps: BuildDictionaryDeps,
): Promise<BuildDictionaryResult> {
  const started = Date.now()
  if (!request.inputPath) return { ok: false, error: 'expected a non-empty `inputPath`' }
  if (!request.targetLang) return { ok: false, error: 'expected a non-empty `targetLang`' }

  let text: string | undefined
  try {
    const parsed = await parseFileToText(request.inputPath)
    if (!parsed.ok) {
      return { ok: false, error: parsed.error ?? 'could not read the input file' }
    }
    if (parsed.kind === 'image') {
      return { ok: false, error: 'image files have no extractable text; run OCR first' }
    }
    text = parsed.text
  } catch (error) {
    return {
      ok: false,
      error: `failed to extract text: ${error instanceof Error ? error.message : String(error)}`,
    }
  }
  if (!text) return { ok: false, error: 'the file contained no extractable text' }

  const maxSegments = request.maxSegments ?? 400
  const minChars = request.minChars ?? 2
  const allSegments = mineSegments(text, minChars)
  const segments = allSegments.slice(0, maxSegments)
  if (segments.length === 0) {
    return { ok: false, error: 'no translatable segments found in the file' }
  }

  const dictionary: Record<string, string> = {}
  const missed: string[] = []

  // 1) KB seed — mandatory terms are authoritative and do not need a model call.
  const kb = new KnowledgeBase()
  await kb.load().catch(() => undefined)
  const resolved = kb.resolve({
    sourceLang: request.sourceLang,
    targetLang: request.targetLang,
    ...(request.glossaryCategory !== undefined ? { category: request.glossaryCategory } : {}),
    ...(request.customerName !== undefined ? { customerName: request.customerName } : {}),
  })
  let kbEntries = 0
  const dictSegments: DictionarySegment[] = []
  for (const term of resolved.terms) {
    if (!term.sourceTerm || !term.targetTerm) continue
    if (dictionary[term.sourceTerm] === term.targetTerm) continue
    dictionary[term.sourceTerm] = term.targetTerm
    dictSegments.push({ source: term.sourceTerm, target: term.targetTerm, origin: 'kb' })
    kbEntries++
  }

  // 2) LLM pass over the remaining segments.
  let llmEntries = 0
  const useLlm = request.useLlm !== false
  if (useLlm) {
    const remaining = segments.filter((s) => dictionary[s] === undefined)
    const batches = batchSegments(remaining)
    let order = 0
    for (const batch of batches) {
      const units = batch.map((sourceText) => ({
        unitId: `seg-${order}`,
        kind: 'paragraph' as const,
        sourceText,
        order: order++,
      }))
      let response: { ok: boolean; units?: TranslateBatchUnitResult[]; error?: string }
      try {
        response = await deps.translateBatch({
          units,
          sourceLang: request.sourceLang,
          targetLang: request.targetLang,
          scene: 'dictionary',
          ...(request.glossaryCategory !== undefined
            ? { glossaryCategory: request.glossaryCategory }
            : {}),
          ...(request.customerName !== undefined ? { customerName: request.customerName } : {}),
        })
      } catch (error) {
        return {
          ok: false,
          error: `provider call failed: ${error instanceof Error ? error.message : String(error)}`,
          elapsedMs: Date.now() - started,
        }
      }
      if (!response.ok && !response.units?.length) {
        return {
          ok: false,
          error: response.error ?? 'the provider returned no translation',
          elapsedMs: Date.now() - started,
        }
      }
      for (const unit of response.units ?? []) {
        const cleaned = unit.translatedText ? extractTranslationText(unit.translatedText) : null
        if (unit.status === 'failed' || !cleaned) {
          missed.push(unit.sourceText)
          continue
        }
        const { text: ruled } = applyKbRules(cleaned, kb, {
          sourceLang: request.sourceLang,
          targetLang: request.targetLang,
          ...(request.customerName !== undefined ? { customerName: request.customerName } : {}),
        })
        dictionary[unit.sourceText] = ruled
        dictSegments.push({ source: unit.sourceText, target: ruled, origin: 'llm' })
        llmEntries++
      }
    }
  }

  // 3) Write it out.
  const outputPath = request.outputPath ?? defaultDictionaryPath(request.dataDir, request.inputPath)
  try {
    const dir = dirname(outputPath)
    mkdirSync(dir, { recursive: true })
    // Sorted keys make the file diffable when a user keeps it in git.
    const sorted: Record<string, string> = {}
    for (const key of Object.keys(dictionary).sort()) sorted[key] = dictionary[key]!
    writeFileSync(outputPath, JSON.stringify(sorted, null, 2), 'utf8')
  } catch (error) {
    return {
      ok: false,
      error: `could not write the dictionary: ${error instanceof Error ? error.message : String(error)}`,
      elapsedMs: Date.now() - started,
    }
  }

  return {
    ok: true,
    dictionaryPath: outputPath,
    kbEntries,
    llmEntries,
    missed,
    totalSegments: segments.length,
    elapsedMs: Date.now() - started,
    segments: dictSegments,
  }
}
