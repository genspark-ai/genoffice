/// Upper bound on journaled cell edits (and recalculated formula cached
/// values) carried inline in one save request. Bulk actions — pasting,
/// sorting, or moving a large range — journal every affected cell, so this
/// must accommodate whole-sheet-scale edit sets. Shared by the zod save
/// schema and the preload's hand-rolled validator (which stays zod-free).
export const PERSIST_TOOL_FIELD_MAX = 2_000
export const SET_RANGE_VALUES_MUTATION = 'sheet.mutation.set-range-values'
export const SET_NUMFMT_MUTATION = 'sheet.mutation.set-numfmt'
export const REMOVE_NUMFMT_MUTATION = 'sheet.mutation.remove-numfmt'
export const COPY_SHEET_COMMAND = 'sheet.command.copy-sheet'
export const MOVE_RANGE_MUTATION = 'sheet.mutation.move-range'
export const MOVE_ROWS_MUTATION = 'sheet.mutation.move-rows'
export const AUTO_FILL_COMMAND = 'sheet.command.auto-fill'
export const SORT_COMMAND_PATTERN = /^sheet\.command\.sort-/
export const FILTER_COMMAND_PATTERN = /^sheet\.command\.filter-/
export const STRUCTURAL_EDIT_COMMAND_PATTERN = /^sheet\.command\.(insert-row|insert-column|delete-row|delete-column|insert-range|delete-range|insert-sheet|delete-sheet)/
export const DV_EDIT_COMMAND_PATTERN = /^sheet\.command\.data-validation-/
export const SET_FROZEN_MUTATION = 'sheet.mutation.set-frozen'
export const REORDER_RANGE_MUTATION = 'sheet.mutation.reorder-range'
export const TOGGLE_GRIDLINES_MUTATION = 'sheet.mutation.toggle-gridlines'
export const MOVE_RANGE_MUTATION_OLD = 'sheet.mutation.move-range'

export const ROW_COLUMN_MUTATIONS: Record<string, { axis: 'row' | 'column'; kind: 'insert' | 'delete' | 'size' | 'hidden' | 'auto-size' }> = {
  'sheet.mutation.insert-row': { axis: 'row', kind: 'insert' },
  'sheet.mutation.insert-column': { axis: 'column', kind: 'insert' },
  'sheet.mutation.delete-row': { axis: 'row', kind: 'delete' },
  'sheet.mutation.delete-column': { axis: 'column', kind: 'delete' },
  'sheet.mutation.set-row-size': { axis: 'row', kind: 'size' },
  'sheet.mutation.set-column-size': { axis: 'column', kind: 'size' },
  'sheet.mutation.set-rows-hidden': { axis: 'row', kind: 'hidden' },
  'sheet.mutation.set-columns-hidden': { axis: 'column', kind: 'hidden' },
  'sheet.mutation.set-row-auto-height': { axis: 'row', kind: 'auto-size' },
}
export const MERGE_MUTATIONS: Record<string, string> = {
  'sheet.mutation.merge': 'merge',
  'sheet.mutation.unmerge': 'unmerge',
}
export const AXIS_ATTR_MUTATIONS: Record<string, { axis: 'row' | 'column'; kind: 'size' | 'hidden' | 'auto-size' }> = {
  'sheet.mutation.set-row-size': { axis: 'row', kind: 'size' },
  'sheet.mutation.set-column-size': { axis: 'column', kind: 'size' },
  'sheet.mutation.set-rows-hidden': { axis: 'row', kind: 'hidden' },
  'sheet.mutation.set-columns-hidden': { axis: 'column', kind: 'hidden' },
  'sheet.mutation.set-row-auto-height': { axis: 'row', kind: 'auto-size' },
}
export const SHEET_LIFECYCLE_MUTATIONS = new Set([
  'sheet.mutation.insert-sheet',
  'sheet.mutation.remove-sheet',
  'sheet.mutation.rename-sheet',
  'sheet.mutation.set-worksheet-order',
  'sheet.mutation.set-worksheet-hidden',
])
export const FILTER_MUTATIONS = new Set([
  'sheet.mutation.set-filter',
  'sheet.mutation.clear-filter',
  'sheet.mutation.clear-all-filter',
])
export const CF_MUTATIONS = new Set([
  'sheet.mutation.set-conditional-formatting',
  'sheet.mutation.delete-conditional-formatting',
  'sheet.mutation.move-conditional-formatting',
  'sheet.mutation.reorder-conditional-formatting',
])
export const DV_MUTATIONS = new Set([
  'sheet.mutation.add-data-validation',
  'sheet.mutation.delete-data-validation',
  'sheet.mutation.set-data-validation',
])
export const DEFINED_NAME_MUTATIONS = new Set([
  'sheet.mutation.add-defined-name',
  'sheet.mutation.delete-defined-name',
  'sheet.mutation.set-defined-name',
])
export const NOTE_MUTATIONS = new Set([
  'sheet.mutation.add-note',
  'sheet.mutation.delete-note',
  'sheet.mutation.update-note',
])

export const FORMULA_MODE_MAX_CELLS = 500_000

/** Upper bound on tool output text persisted per step in the run transcript (bytes). */
export const PERSIST_TOOL_FIELD_MAX_BYTES = 4_000

/** localStorage key for the docs AI chat history (legacy, migrated to project-store). */
export const CHAT_STORAGE_KEY = 'docs-ai-chat'

/** Byte budget for AI-editable content sent to the model per read-range call. */
export const FORMULA_MODE_MAX_CELLS_PER_RANGE = 200_000

/** Maximum number of characters the lazy snapshot serializer will send per range read. */
export const MAX_CHARS_PER_RANGE_READ = 1_000_000

/** Maximum number of cells the lazy snapshot serializer will send per range read. */
export const MAX_CELLS_PER_RANGE_READ = 500_000

/** Maximum number of rows the lazy snapshot serializer will read at once. */
export const MAX_ROWS_PER_RANGE_READ = 100_000

export const IPC_CHANNELS = {
  // Sheets lifecycle
  openWorkbook: 'sheets:open-workbook',
  saveWorkbook: 'sheets:save-workbook',
  getWorkbookFile: 'sheets:get-workbook-file',
  closeWorkbook: 'sheets:close-workbook',

  // Theme
  getTheme: 'app:get-theme',
  onThemeChanged: 'app:on-theme-changed',

  // Language
  getLanguage: 'app:get-language',
  onLanguageChanged: 'app:on-language-changed',

  // AI
  aiStream: 'ai:stream',
  aiStreamChunk: 'ai:stream-chunk',
  aiStreamCancel: 'ai:stream-cancel',
  aiTestConnection: 'ai:test-connection',
  aiGetSettings: 'ai:get-settings',
  aiSaveSettings: 'ai:save-settings',
  aiGskStatus: 'ai:gsk-status',
  aiGskLogin: 'ai:gsk-login',
  aiImageSearch: 'ai:image-search',
  aiFetchImage: 'ai:fetch-image',
  // sheets: prefix — slides' ai:generate-image only registers once a slides view exists
  aiGenerateImage: 'sheets:ai-generate-image',
  aiOllamaModels: 'ai:ollama-models',
  aiTestConnection: 'ai:test-connection',
  // Chat attachments (sheets: prefix — docs already registers global files:* in
  // the shell; avoids collisions)
  captureScreenSources: 'sheets:capture-screen-sources',
  captureScreenSource: 'sheets:capture-screen-source',
  filesPick: 'sheets:files-pick',
  filesAddPaths: 'sheets:files-add-paths',
  filesAddPastedImage: 'sheets:files-add-pasted-image',
  filesReadImage: 'sheets:files-read-image',
  filesList: 'sheets:files-list',

  // Workbook operations
  readFormulaCells: 'sheets:read-formula-cells',
  recalcWorkbook: 'sheets:recalc-workbook',
  readRange: 'sheets:read-range',
  getMedia: 'sheets:get-media',
  getPivot: 'sheets:get-pivot',
  savePivot: 'sheets:save-pivot',

  // PDF export
  exportPdf: 'sheets:export-pdf',

  // Workbook file info
  workbookRenamed: 'sheets:workbook-renamed',

  // Chat persistence
  chatLoad: 'ai:chat-load',
  chatSave: 'ai:chat-save',

  // Workspace Q&A
  workspaceIndex: 'workspace:index',
  workspaceSearch: 'workspace:search',

  // Local image loading
  loadLocalImage: 'sheets:load-local-image',

  // Screen capture
  screenCapture: 'sheets:screen-capture',
} as const
