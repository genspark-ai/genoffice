import type {
  FeatureDraft,
  FeaturePatch,
  FeatureRecord,
  KnowledgeEntry,
  ScriptRunOutcome,
} from '@genoffice/agent-core'

/**
 * The feature registry's channel names and payload types, shared by the
 * main-process handlers and the preload bridge. Kept free of electron so both
 * sides, and the tests, can import it.
 */

export const FEATURE_CHANNELS = {
  list: 'features:list',
  get: 'features:get',
  create: 'features:create',
  update: 'features:update',
  delete: 'features:delete',
  readDocument: 'features:read-document',
  /** run a script feature read-only against a document and report the plan */
  run: 'features:run',
  /** apply a script feature to a document: needs confirm, and the file must be closed */
  apply: 'features:apply',
  /** what the hooks have done recently, newest last */
  hookActivity: 'features:hook-activity',
  /** write the registry out as a shareable file */
  packExport: 'features:pack-export',
  /** read a pack and describe what it declares, before anything is written */
  packPreview: 'features:pack-preview',
  /** install a pack that has been previewed */
  packInstall: 'features:pack-install',
} as const

/**
 * The native view slot's channels. The renderer reports where the view belongs
 * and the shell mounts a real WebContentsView there; the renderer never owns a
 * webview tag, so a mod cannot reach one.
 */
export const FEATURE_VIEW_CHANNELS = {
  mount: 'features:view-mount',
  unmount: 'features:view-unmount',
} as const

/** The rectangle a slot reserves, in the window's content coordinates. */
export interface SlotViewRect {
  x: number
  y: number
  width: number
  height: number
}

export interface SlotViewMountRequest {
  /** one slot instance */
  id: string
  /** an empty address unmounts */
  url: string
  rect: SlotViewRect
}

export type { KnowledgeEntry }

/** One feature inside a pack, as the import review shows it. */
export interface PackLine {
  id: string
  name: string
  kind: string
  /** one line per thing this feature declares */
  declares: string[]
}

/** What a pack declares, before anything is installed. */
export interface PackDescription {
  version: number
  exportedAt: string
  count: number
  lines: PackLine[]
  /** the parts a reader should look at twice */
  warnings: string[]
}

/** What one hook did when an event happened (mirrors the main process's record). */
export interface HookActivity {
  at: string
  hookId: string
  event: string
  path: string
  log: string[]
  /** reasons this hook gave for stopping the event, when it gave any */
  vetoes?: string[]
  error?: string
}

/** What an apply did, on top of the plan a read-only run reports. */
export interface ScriptApplyResult extends ScriptRunOutcome {
  /** the file that was replaced */
  written: string
  /** bytes written */
  bytes: number
}

/** What read_document returns: parsed text, never a file handle or a writer. */
export interface ReadDocumentResult {
  name: string
  text: string
  chars: number
}

export type { FeatureDraft, FeaturePatch, FeatureRecord, ScriptRunOutcome }
