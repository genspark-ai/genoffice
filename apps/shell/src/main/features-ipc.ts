import { type BrowserWindow, dialog, ipcMain } from 'electron'
import {
  validateAppearance,
  validateCommand,
  validateHook,
  validateKnowledge,
  validateScript,
  validateSlot,
  type FeatureAppearance,
  type FeatureDraft,
  type FeaturePatch,
  type FeatureRecord,
  type FeatureCommand,
  type FeatureHook,
  type FeatureKnowledge,
  type SlotDeclaration,
  type ScriptRunOutcome,
} from '@genoffice/agent-core'
import {
  FEATURE_CHANNELS,
  type HookActivity,
  type ReadDocumentResult,
  type ScriptApplyResult,
} from '../shared/feature-channels'
import { applyScriptToDocument, runScriptReadOnly, scriptFormatFor } from './script-host'
import type { ScriptFormat } from '@genoffice/agent-core'
import { createHookRunner, sandboxHookRunner, setLiveHookRunner } from './feature-hooks'
import {
  buildPack,
  describePack,
  installPack,
  readPack,
  type PackDescription,
} from './feature-packs'
import {
  createFeature,
  getFeature,
  listFeatures,
  removeFeature,
  updateFeature,
} from './feature-store'
import { readFileSync, writeFileSync } from 'node:fs'
import { readDocumentText } from './feature-documents'

/**
 * The feature registry's IPC boundary: the only way the mod assistant's panel
 * reaches the records. Every payload is schema-checked here (SECURITY.md): the
 * renderer's copy of a name or a path is data, never trusted, and the store
 * re-validates the id on top because ids name files.
 */

export interface FeaturesIpcDeps {
  /** directory holding one JSON file per feature */
  dir: string
  /** read-only document reader; injected so the IPC module stays testable */
  readDocument?: (path: string) => Promise<ReadDocumentResult>
  /**
   * Whether a document is open in an editor right now. An apply refuses those:
   * the file belongs to that editor, and a write from here would be lost on its
   * next save. Injected so the IPC module stays testable.
   */
  isDocumentOpen?: (path: string) => boolean
  /** the window a file dialog belongs to; absent means the pack channels refuse */
  window?: () => BrowserWindow | null
}

export interface FeaturesIpc {
  /** tell the hooks that an event happened (the shell calls this on tab changes) */
  emitHook(event: 'document.opened' | 'document.closed', path: string): void
  dispose(): void
}

function asObject(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {}
}

function asText(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined
}

/**
 * A script that declared its format runs on that format's files only: the run
 * would otherwise fail deep inside the sandbox with "$.sheets is not defined",
 * so the mismatch is named here instead.
 */
function requireScriptFormat(format: ScriptFormat | undefined, path: string): void {
  if (!format) return
  const actual = scriptFormatFor(path)
  if (actual && actual !== format) {
    throw new Error(
      'that script targets ' + format + ' documents, and ' + path + ' is a ' + actual + ' file',
    )
  }
}

/**
 * A slot payload from the renderer is data, never trusted: undefined means "not
 * mentioned", null means "stop drawing it", anything else must validate. The
 * error names every offending field, and the renderer's panel shows it.
 */
function slotFrom(value: unknown): SlotDeclaration | null | undefined {
  if (value === undefined) return undefined
  if (value === null) return null
  const validated = validateSlot(value)
  if (!validated.ok) {
    throw new Error(
      'invalid slot: ' +
        validated.errors.map((error) => error.field + ' ' + error.message).join('; '),
    )
  }
  return validated.slot
}

/**
 * An appearance payload is a choice from host-defined names; a raw value is
 * refused with the field named, exactly like a malformed slot.
 */
function appearanceFrom(value: unknown): FeatureAppearance | null | undefined {
  if (value === undefined) return undefined
  if (value === null) return null
  const validated = validateAppearance(value)
  if (!validated.ok) {
    throw new Error(
      'invalid appearance: ' +
        validated.errors.map((error) => error.field + ' ' + error.message).join('; '),
    )
  }
  return validated.appearance
}

/** A script payload is code the host will run: shape and size are checked here too. */
function scriptFrom(value: unknown): { source: string } | null | undefined {
  if (value === undefined) return undefined
  if (value === null) return null
  const validated = validateScript(value)
  if (!validated.ok) {
    throw new Error(
      'invalid script: ' +
        validated.errors.map((error) => error.field + ' ' + error.message).join('; '),
    )
  }
  return validated.script
}

/** A command payload from the renderer: the same action rules as a drawn control. */
function commandFrom(value: unknown): FeatureCommand | null | undefined {
  if (value === undefined) return undefined
  if (value === null) return null
  const validated = validateCommand(value)
  if (!validated.ok) {
    throw new Error(
      'invalid command: ' +
        validated.errors.map((error) => error.field + ' ' + error.message).join('; '),
    )
  }
  return validated.command
}

/** A knowledge payload from the renderer: local, absolute, read-only. */
function knowledgeFrom(value: unknown): FeatureKnowledge | null | undefined {
  if (value === undefined) return undefined
  if (value === null) return null
  const validated = validateKnowledge(value)
  if (!validated.ok) {
    throw new Error(
      'invalid knowledge: ' +
        validated.errors.map((error) => error.field + ' ' + error.message).join('; '),
    )
  }
  return validated.knowledge
}

/** A hook payload from the renderer: the same shape rules as a script. */
function hookFrom(value: unknown): FeatureHook | null | undefined {
  if (value === undefined) return undefined
  if (value === null) return null
  const validated = validateHook(value)
  if (!validated.ok) {
    throw new Error(
      'invalid hook: ' +
        validated.errors.map((error) => error.field + ' ' + error.message).join('; '),
    )
  }
  return validated.hook
}

export function registerFeaturesIpc(deps: FeaturesIpcDeps): FeaturesIpc {
  const readDocument = deps.readDocument ?? readDocumentText
  // the hooks a feature registered, and what they did recently
  const hooks = createHookRunner({
    list: () => listFeatures(deps.dir),
    run: sandboxHookRunner(),
  })

  ipcMain.handle(FEATURE_CHANNELS.hookActivity, (): HookActivity[] => hooks.activity())
  // the save guard asks this runner whether a hook stops a save
  setLiveHookRunner(hooks)

  ipcMain.handle(FEATURE_CHANNELS.list, (): FeatureRecord[] => listFeatures(deps.dir))

  ipcMain.handle(FEATURE_CHANNELS.get, (_event, raw: unknown): FeatureRecord | null => {
    const id = asText(asObject(raw).id)
    return id ? getFeature(deps.dir, id) : null
  })

  ipcMain.handle(FEATURE_CHANNELS.create, (_event, raw: unknown): FeatureRecord => {
    const draft = asObject(raw)
    const name = asText(draft.name)
    const description = asText(draft.description)
    if (name === undefined || description === undefined) {
      throw new Error('create needs a name and a description')
    }
    // create takes a slot or an appearance or nothing; an explicit null is the
    // same as omitting it
    const slot = slotFrom(draft.slot)
    const appearance = appearanceFrom(draft.appearance)
    const script = scriptFrom(draft.script)
    const hook = hookFrom(draft.hook)
    const knowledge = knowledgeFrom(draft.knowledge)
    const command = commandFrom(draft.command)
    const payload: FeatureDraft = {
      name,
      description,
      ...(asText(draft.kind) === undefined ? {} : { kind: asText(draft.kind)! }),
      ...(slot ? { slot } : {}),
      ...(appearance ? { appearance } : {}),
      ...(script ? { script } : {}),
      ...(hook ? { hook } : {}),
      ...(knowledge ? { knowledge } : {}),
      ...(command ? { command } : {}),
    }
    return createFeature(deps.dir, payload)
  })

  ipcMain.handle(FEATURE_CHANNELS.update, (_event, raw: unknown): FeatureRecord | null => {
    const request = asObject(raw)
    const id = asText(request.id)
    if (!id) throw new Error('update needs an id')
    const slot = slotFrom(request.slot)
    const appearance = appearanceFrom(request.appearance)
    const script = scriptFrom(request.script)
    const hook = hookFrom(request.hook)
    const knowledge = knowledgeFrom(request.knowledge)
    const command = commandFrom(request.command)
    const patch: FeaturePatch = {
      ...(asText(request.name) === undefined ? {} : { name: asText(request.name)! }),
      ...(asText(request.description) === undefined
        ? {}
        : { description: asText(request.description)! }),
      ...(asText(request.kind) === undefined ? {} : { kind: asText(request.kind)! }),
      ...(slot === undefined ? {} : { slot }),
      ...(appearance === undefined ? {} : { appearance }),
      ...(script === undefined ? {} : { script }),
      ...(hook === undefined ? {} : { hook }),
      ...(knowledge === undefined ? {} : { knowledge }),
      ...(command === undefined ? {} : { command }),
    }
    return updateFeature(deps.dir, id, patch)
  })

  ipcMain.handle(FEATURE_CHANNELS.delete, (_event, raw: unknown): boolean => {
    const id = asText(asObject(raw).id)
    return id ? removeFeature(deps.dir, id) : false
  })

  ipcMain.handle(
    FEATURE_CHANNELS.readDocument,
    async (_event, raw: unknown): Promise<ReadDocumentResult> => {
      const path = asText(asObject(raw).path)
      if (!path) throw new Error('read needs a path')
      return readDocument(path)
    },
  )

  // A dry run: the feature's own script, read-only, against one document. The
  // registry is the source of the source, so a caller cannot run code that was
  // never recorded (and validated) here.
  ipcMain.handle(FEATURE_CHANNELS.run, async (_event, raw: unknown): Promise<ScriptRunOutcome> => {
    const request = asObject(raw)
    const id = asText(request.id)
    const path = asText(request.path)
    if (!id || !path) throw new Error('a dry run needs an id and a path')
    const feature = getFeature(deps.dir, id)
    if (!feature) throw new Error('no feature with id ' + id)
    if (!feature.script) throw new Error('feature ' + id + ' is not a script')
    requireScriptFormat(feature.script.format, path)
    return runScriptReadOnly({
      scriptId: id,
      source: feature.script.source,
      documentPath: path,
      ...(feature.script.format ? { format: feature.script.format } : {}),
      ...(request.input && typeof request.input === 'object' && !Array.isArray(request.input)
        ? { input: request.input as Record<string, unknown> }
        : {}),
      ...(typeof request.timeoutMs === 'number' ? { timeoutMs: request.timeoutMs } : {}),
    })
  })

  // Packs: the only way a feature leaves or enters this machine as a file. Export
  // writes what the registry holds; import shows what a pack declares *before*
  // anything is written, and an existing feature is never replaced by default.
  ipcMain.handle(FEATURE_CHANNELS.packExport, async (): Promise<string | null> => {
    const win = deps.window?.()
    if (!win) throw new Error('export needs a window')
    const pack = buildPack(listFeatures(deps.dir))
    const picked = await dialog.showSaveDialog(win, {
      title: 'Export features',
      defaultPath: 'genoffice-features.json',
      filters: [{ name: 'GenOffice features', extensions: ['json'] }],
    })
    if (picked.canceled || !picked.filePath) return null
    writeFileSync(picked.filePath, JSON.stringify(pack, null, 2), 'utf8')
    return picked.filePath
  })

  ipcMain.handle(
    FEATURE_CHANNELS.packPreview,
    async (): Promise<{
      path: string | null
      description?: PackDescription
      errors?: string[]
    }> => {
      const win = deps.window?.()
      if (!win) throw new Error('import needs a window')
      const picked = await dialog.showOpenDialog(win, {
        title: 'Import features',
        properties: ['openFile'],
        filters: [{ name: 'GenOffice features', extensions: ['json'] }],
      })
      const path = picked.canceled ? undefined : picked.filePaths[0]
      if (!path) return { path: null }
      const read = readPack(readFileSync(path, 'utf8'))
      if (!read.ok) return { path, errors: read.errors }
      return {
        path,
        description: describePack(
          read.pack,
          listFeatures(deps.dir).map((feature) => feature.id),
        ),
      }
    },
  )

  ipcMain.handle(
    FEATURE_CHANNELS.packInstall,
    (_event, raw: unknown): { installed: string[]; replaced: string[]; skipped: string[] } => {
      const request = asObject(raw)
      const path = asText(request.path)
      if (!path) throw new Error('installing a pack needs a path')
      const read = readPack(readFileSync(path, 'utf8'))
      if (!read.ok) throw new Error(read.errors.join('; '))
      return installPack(deps.dir, read.pack, { replace: request.replace === true })
    },
  )

  // Apply a script feature to a document. Two gates beyond validation: the
  // caller must say confirm: true, and the file must not be open in an editor.
  ipcMain.handle(
    FEATURE_CHANNELS.apply,
    async (_event, raw: unknown): Promise<ScriptApplyResult> => {
      const request = asObject(raw)
      const id = asText(request.id)
      const path = asText(request.path)
      if (!id || !path) throw new Error('an apply needs an id and a path')
      const feature = getFeature(deps.dir, id)
      if (!feature) throw new Error('no feature with id ' + id)
      if (!feature.script) throw new Error('feature ' + id + ' is not a script')
      requireScriptFormat(feature.script.format, path)
      if (deps.isDocumentOpen?.(path)) {
        throw new Error('that document is open in an editor: save or close it first')
      }
      return applyScriptToDocument({
        scriptId: id,
        source: feature.script.source,
        documentPath: path,
        confirm: request.confirm === true,
        ...(feature.script.format ? { format: feature.script.format } : {}),
        ...(request.input && typeof request.input === 'object' && !Array.isArray(request.input)
          ? { input: request.input as Record<string, unknown> }
          : {}),
        ...(typeof request.timeoutMs === 'number' ? { timeoutMs: request.timeoutMs } : {}),
      })
    },
  )

  return {
    emitHook(event, path): void {
      hooks.emit(event, path)
    },
    dispose(): void {
      for (const channel of Object.values(FEATURE_CHANNELS)) ipcMain.removeHandler(channel)
    },
  }
}
