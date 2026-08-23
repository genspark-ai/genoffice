/// Upper bound on journaled cell edits (and recalculated formula cached
/// values) carried inline in one save request. Bulk actions — pasting,
/// sorting, or moving a large range — journal every affected cell, so this
/// must accommodate whole-sheet-scale edit sets. Shared by the zod save
/// schema and the preload's hand-rolled validator (which stays zod-free).
/// Larger edit sets go through the chunked transfer below instead.
export const MAX_SAVE_EDITS = 1_000_000

/// Chunked save-edit transfer: above MAX_SAVE_EDITS the renderer uploads the
/// edits in slices of SAVE_EDITS_CHUNK_MAX (bounding each IPC message's
/// structured-clone spike), then sends the save request with an
/// editsTransferId instead of an inline edits array. The absolute ceiling is
/// MAX_SAVE_EDITS_TOTAL — roughly "one bulk action over an entire very large
/// sheet"; past that the bottleneck is main-process memory, not transport.
export const SAVE_EDITS_CHUNK_MAX = 100_000
export const MAX_SAVE_EDITS_TOTAL = 10_000_000

export const IPC_CHANNELS = {
  selectWorkbook: 'workbook:select',
  readWorkbookRange: 'workbook:read-range',
  readWorkbookFormulas: 'workbook:read-formulas',
  recalcWorkbook: 'workbook:recalc',
  readWorkbookMedia: 'workbook:read-media',
  readPivotDefinition: 'workbook:read-pivot-definition',
  readLocalImage: 'shell:read-local-image',
  closeWorkbook: 'workbook:close',
  saveWorkbook: 'workbook:save',
  /** Chunked upload of a large save's cell edits, consumed by the next save */
  saveEditsBegin: 'workbook:save-edits-begin',
  saveEditsChunk: 'workbook:save-edits-chunk',
  saveEditsAbort: 'workbook:save-edits-abort',
  /** Crash-recovery copy of a dirty workbook, written under userData */
  writeWorkbookRecovery: 'workbook:write-recovery',
  /** Main found a newer recovery copy while opening; renderer shows the styled prompt */
  recoveryPrompt: 'workbook:recovery-prompt',
  recoveryPromptReply: 'workbook:recovery-prompt-reply',
  autoRenameWorkbook: 'workbook:auto-rename',
  workbookRenamed: 'workbook:renamed',
  pendingEditsChanged: 'workbook:pending-edits',
  closeSaveRequest: 'workbook:close-save-request',
  closeSaveResult: 'workbook:close-save-result',
  exportPdf: 'workbook:export-pdf',
  openExternal: 'shell:open-external',
  menuAction: 'menu:action',
  aiGetSettings: 'ai:get-settings',
  aiSetSettings: 'ai:set-settings',
  aiChat: 'ai:chat',
  aiStream: 'ai:stream',
  aiStreamCancel: 'ai:stream-cancel',
  aiStreamChunk: 'ai:stream-chunk',
  aiGskStatus: 'ai:gsk-status',
  aiGskLogin: 'ai:gsk-login',
  aiImageSearch: 'ai:image-search',
  aiFetchImage: 'ai:fetch-image',
  // sheets: prefix — slides' ai:generate-image only registers once a slides view exists
  aiGenerateImage: 'sheets:ai-generate-image',
  // Chat attachments (sheets: prefix — docs already registers global files:* in
  // the shell; avoids collisions)
  captureScreenSources: 'sheets:capture-screen-sources',
  captureScreenSource: 'sheets:capture-screen-source',
  filesPick: 'sheets:files-pick',
  filesAdd: 'sheets:files-add',
  filesAddPastedImage: 'sheets:files-add-pasted-image',
  filesRead: 'sheets:files-read',
  filesReadImage: 'sheets:files-read-image',
} as const
