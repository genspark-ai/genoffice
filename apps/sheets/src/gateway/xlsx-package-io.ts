import { randomUUID } from 'node:crypto'
import { mkdir, mkdtemp, open, readFile, rename, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'

import { z } from 'zod'

import type { WorkbookChartEdit, WorkbookVisualEdit } from '../shared/desktop-api'
import type { SheetFilterState } from './xlsx-filter'
import type { DefinedNamesState } from './xlsx-defined-names'
import type { SheetPageSetupState } from './xlsx-page-setup'
import type {
  CellEdit,
  EntrySource,
  MutationPlan,
  PivotRefreshUpdate,
  SheetCfState,
  SheetDvState,
  SheetHyperlinkEdits,
  SheetNoteState,
  SheetProtectionState,
  SheetPivotAddition,
  SheetSparklineAddition,
  SheetStructuralOps,
  SheetTableAddition,
  SheetVisualAddition,
  SheetFormulaValues,
} from './xlsx-gateway'
import { planCellEditsToXlsx } from './xlsx-gateway'
import type { SheetEditPlan } from './xlsx-sheets'

/// Mirrors the sidecar's per-entry extraction cap: only entries the gateway
/// patches must fit in memory — the archive as a whole has no size limit.
const MAX_PATCH_ENTRY_BYTES = 256 * 1024 * 1024

const archiveEntrySchema = z.object({
  name: z.string(),
  crc32: z.number(),
  compressedSize: z.number(),
  uncompressedSize: z.number(),
})

export type ArchiveEntry = z.infer<typeof archiveEntrySchema>

const manifestResultSchema = z.object({ entries: z.array(archiveEntrySchema) })
const readEntriesResultSchema = z.object({
  entries: z.array(z.object({ name: z.string(), path: z.string() })),
})
const scanEntriesResultSchema = z.object({ matches: z.array(z.string()) })
const saveArchiveResultSchema = z.object({
  beforeEntries: z.array(archiveEntrySchema),
  afterEntries: z.array(archiveEntrySchema),
})

/// The subset of XlsxSidecarClient the streaming save path needs; tests can
/// substitute a stub.
export interface ArchiveClient {
  archiveManifest(path: string): Promise<unknown>
  readEntries(input: {
    readonly path: string
    readonly entries: readonly string[]
    readonly outputDir: string
  }): Promise<unknown>
  scanEntries(input: {
    readonly path: string
    readonly entries: readonly string[]
    readonly needle: string
  }): Promise<unknown>
  saveArchive(input: {
    readonly sourcePath: string
    readonly targetPath: string
    readonly replacements: readonly { name: string; contentPath: string }[]
    readonly removals: readonly string[]
    readonly additions: readonly { name: string; contentPath: string }[]
  }): Promise<unknown>
}

export interface StreamingSaveRequest {
  readonly client: ArchiveClient
  readonly sourcePath: string
  readonly targetPath: string
  readonly edits: readonly CellEdit[]
  readonly structuralOps?: readonly SheetStructuralOps[] | undefined
  readonly chartEdits?: readonly WorkbookChartEdit[] | undefined
  readonly sheetPlan?: SheetEditPlan | undefined
  readonly filterStates?: readonly SheetFilterState[] | undefined
  readonly hyperlinkEdits?: readonly SheetHyperlinkEdits[] | undefined
  readonly cfStates?: readonly SheetCfState[] | undefined
  readonly dvStates?: readonly SheetDvState[] | undefined
  readonly sheetProtections?: readonly SheetProtectionState[] | undefined
  readonly definedNamesState?: DefinedNamesState | null | undefined
  readonly visualAdditions?: readonly SheetVisualAddition[] | undefined
  readonly pageSetupStates?: readonly SheetPageSetupState[] | undefined
  readonly noteStates?: readonly SheetNoteState[] | undefined
  readonly tableAdditions?: readonly SheetTableAddition[] | undefined
  readonly pivotAdditions?: readonly SheetPivotAddition[] | undefined
  readonly pivotCacheRefreshPaths?: readonly string[] | undefined
  readonly pivotRefreshUpdates?: readonly PivotRefreshUpdate[] | undefined
  readonly visualEdits?: readonly WorkbookVisualEdit[] | undefined
  readonly sparklineAdditions?: readonly SheetSparklineAddition[] | undefined
  /// Recalculated formula-cell values written into <v>
  readonly formulaValues?: readonly SheetFormulaValues[] | undefined
}

export interface StreamingSaveResult {
  readonly touchedEntries: readonly string[]
  readonly removedEntries: readonly string[]
  readonly addedEntries: readonly string[]
}

/// One-shot text read of an archive entry via the sidecar (extract to a
/// temp dir, read, clean up).
export async function readArchiveEntryText(
  client: ArchiveClient,
  sourcePath: string,
  entryName: string,
): Promise<string> {
  const workDir = await mkdtemp(join(tmpdir(), 'ai-excel-read-'))
  try {
    const extracted = readEntriesResultSchema.parse(
      await client.readEntries({ path: sourcePath, entries: [entryName], outputDir: workDir }),
    )
    const filePath = extracted.entries[0]?.path
    if (!filePath) throw new Error(`Workbook is missing ${entryName}.`)
    return await readFile(filePath, 'utf8')
  } finally {
    await rm(workDir, { recursive: true, force: true })
  }
}

/// Streaming save channel: the gateway plans patched entry contents in
/// memory, the sidecar reassembles the archive on disk — untouched entries
/// are raw-copied compressed bytes, verified afterwards via CRC manifests.
export async function saveWorkbookViaSidecar(
  request: StreamingSaveRequest,
): Promise<StreamingSaveResult> {
  const workDir = await mkdtemp(join(tmpdir(), 'ai-excel-save-'))
  const temporaryTarget = join(dirname(request.targetPath), `.${randomUUID()}.tmp.xlsx`)
  try {
    const manifest = manifestResultSchema.parse(
      await request.client.archiveManifest(request.sourcePath),
    ).entries
    const source = createSidecarEntrySource(request.client, request.sourcePath, manifest, workDir)
    const plan = await planCellEditsToXlsx(
      source,
      request.edits,
      request.structuralOps ?? [],
      request.chartEdits ?? [],
      request.sheetPlan,
      request.filterStates ?? [],
      request.hyperlinkEdits ?? [],
      request.cfStates ?? [],
      request.dvStates ?? [],
      request.sheetProtections ?? [],
      request.definedNamesState ?? null,
      request.visualAdditions ?? [],
      request.pageSetupStates ?? [],
      request.noteStates ?? [],
      request.tableAdditions ?? [],
      request.pivotAdditions ?? [],
      request.pivotCacheRefreshPaths ?? [],
      request.pivotRefreshUpdates ?? [],
      request.visualEdits ?? [],
      request.sparklineAdditions ?? [],
      request.formulaValues ?? [],
    )

    const replacements = await writePlanContents(workDir, 'replace', plan.replaced)
    const additions = [
      ...(await writePlanContents(workDir, 'add', plan.added)),
      ...(await writePlanContents(workDir, 'add-bin', plan.addedBinary)),
    ]
    const result = saveArchiveResultSchema.parse(
      await request.client.saveArchive({
        sourcePath: request.sourcePath,
        targetPath: temporaryTarget,
        replacements,
        removals: plan.removedEntries,
        additions,
      }),
    )

    // The source manifest was read before planning; if the file changed on
    // disk in between, the save's own before-manifest exposes the drift.
    if (!manifestsEqual(manifest, result.beforeEntries)) {
      throw new Error('The workbook changed on disk while saving — aborted.')
    }
    assertManifestPreserved(plan, result.beforeEntries, result.afterEntries)

    await promoteFileAtomically(temporaryTarget, request.targetPath)
    return {
      touchedEntries: plan.touchedEntries,
      removedEntries: plan.removedEntries,
      addedEntries: plan.addedEntries,
    }
  } catch (error: unknown) {
    await rm(temporaryTarget, { force: true })
    throw error
  } finally {
    await rm(workDir, { recursive: true, force: true })
  }
}

function createSidecarEntrySource(
  client: ArchiveClient,
  sourcePath: string,
  manifest: readonly ArchiveEntry[],
  workDir: string,
): EntrySource {
  const entryByName = new Map(manifest.map((entry) => [entry.name, entry]))
  const cache = new Map<string, string>()
  let extractionCount = 0
  return {
    paths: async () => manifest.map((entry) => entry.name),
    has: async (path) => entryByName.has(path),
    canPatch: async (path) =>
      (entryByName.get(path)?.uncompressedSize ?? 0) <= MAX_PATCH_ENTRY_BYTES,
    containsText: async (path, needle) => {
      const scanned = scanEntriesResultSchema.parse(
        await client.scanEntries({ path: sourcePath, entries: [path], needle }),
      )
      return scanned.matches.includes(path)
    },
    readText: async (path) => {
      const cached = cache.get(path)
      if (cached !== undefined) return cached
      const entry = entryByName.get(path)
      if (!entry) throw new Error(`Workbook is missing ${path}.`)
      if (entry.uncompressedSize > MAX_PATCH_ENTRY_BYTES) {
        throw new Error(
          `${path} is ${entry.uncompressedSize} bytes uncompressed — too large to edit. ` +
            'Entries above 256MB can be preserved but not patched.',
        )
      }
      const extractDir = join(workDir, `extract-${extractionCount}`)
      extractionCount += 1
      await mkdir(extractDir, { recursive: true })
      const extracted = readEntriesResultSchema.parse(
        await client.readEntries({ path: sourcePath, entries: [path], outputDir: extractDir }),
      )
      const filePath = extracted.entries[0]?.path
      if (!filePath) throw new Error(`Sidecar did not extract ${path}.`)
      const content = await readFile(filePath, 'utf8')
      cache.set(path, content)
      return content
    },
  }
}

async function writePlanContents(
  workDir: string,
  prefix: string,
  contents: ReadonlyMap<string, string | Uint8Array>,
): Promise<{ name: string; contentPath: string }[]> {
  const written: { name: string; contentPath: string }[] = []
  let index = 0
  for (const [name, content] of contents) {
    const contentPath = join(workDir, `${prefix}-${index}.bin`)
    index += 1
    if (typeof content === 'string') await writeFile(contentPath, content, 'utf8')
    else await writeFile(contentPath, content)
    written.push({ name, contentPath })
  }
  return written
}

function manifestsEqual(left: readonly ArchiveEntry[], right: readonly ArchiveEntry[]): boolean {
  if (left.length !== right.length) return false
  const key = (entry: ArchiveEntry): string =>
    `${entry.name}\u0000${entry.crc32}\u0000${entry.compressedSize}\u0000${entry.uncompressedSize}`
  const leftKeys = new Set(left.map(key))
  return right.every((entry) => leftKeys.has(key(entry)))
}

/// Fails closed when the saved archive differs from the source anywhere the
/// plan did not declare: untouched entries must survive byte-preserving raw
/// copy (identical CRC and sizes), removals must vanish, additions must
/// appear, and nothing else may exist.
export function assertManifestPreserved(
  plan: Pick<MutationPlan, 'replaced' | 'added' | 'removedEntries'> &
    Partial<Pick<MutationPlan, 'addedBinary'>>,
  before: readonly ArchiveEntry[],
  after: readonly ArchiveEntry[],
): void {
  const removed = new Set(plan.removedEntries)
  const replaced = new Set(plan.replaced.keys())
  const added = new Set([...plan.added.keys(), ...(plan.addedBinary?.keys() ?? [])])
  const beforeByName = new Map(before.map((entry) => [entry.name, entry]))
  const afterByName = new Map(after.map((entry) => [entry.name, entry]))

  for (const entry of before) {
    const saved = afterByName.get(entry.name)
    if (removed.has(entry.name)) {
      if (saved) throw new Error(`Saving should have removed ${entry.name} but did not — aborted.`)
      continue
    }
    if (!saved) throw new Error(`Saving would drop ${entry.name} — aborted.`)
    if (replaced.has(entry.name)) continue
    if (
      saved.crc32 !== entry.crc32 ||
      saved.compressedSize !== entry.compressedSize ||
      saved.uncompressedSize !== entry.uncompressedSize
    ) {
      throw new Error(`Saving would unexpectedly modify ${entry.name} — aborted.`)
    }
  }
  for (const name of added) {
    if (beforeByName.has(name)) {
      throw new Error(`Saving should have created ${name} but it already existed — aborted.`)
    }
    if (!afterByName.has(name)) {
      throw new Error(`Saving should have created ${name} but did not — aborted.`)
    }
  }
  for (const entry of after) {
    if (!beforeByName.has(entry.name) && !added.has(entry.name)) {
      throw new Error(`Saving would unexpectedly create ${entry.name} — aborted.`)
    }
  }
}

async function promoteFileAtomically(temporaryPath: string, path: string): Promise<void> {
  const handle = await open(temporaryPath, 'r')
  try {
    await handle.sync()
  } finally {
    await handle.close()
  }
  await rename(temporaryPath, path)
}
