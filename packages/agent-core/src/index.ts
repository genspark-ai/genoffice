export type {
  AgentImage,
  AgentMessage,
  AgentStreamCallbacks,
  AgentStreamHandle,
  AgentStreamRequest,
  AgentTokenUsage,
  AgentToolCall,
  AgentToolDef,
  AgentToolResult,
  AgentTransport,
  MindMapTreeNode,
  ToolDisplay,
  ToolExecution,
} from './types'
export { composeSkills } from './skill'
export { createKnowledgeBaseSkill, formatKbSize } from './knowledge-base-skill'
export type {
  KnowledgeBaseDeps,
  KbFileInfo,
  KbReadPage,
  KbSearchHit,
  KbSourceInfo,
} from './knowledge-base-skill'
export {
  createDocumentSkill,
  CREATE_DOCUMENT_MAX_CHARS,
  CREATE_DOCUMENT_TOOL,
} from './create-document-skill'
export type {
  CreateDocumentDeps,
  CreateDocumentFormat,
  CreateDocumentRequest,
  CreateDocumentSkill,
} from './create-document-skill'
export type { AgentSkill, ExecutedToolCall } from './skill'
export { createFeaturesSkill } from './features-skill'
export { createModsSkill } from './mods-skill'
export type { ModsSkillDeps } from './mods-skill'
export {
  SCRIPT_API_SUMMARY,
  SCRIPT_FORMATS,
  SCRIPT_LIMITS,
  scriptsIn,
  validateScript,
} from './feature-script'
export type { ScriptFormat } from './feature-script'
export { HOOK_EVENTS, describeHooks, hooksFor, validateHook } from './feature-hooks'
export { COMMAND_LABEL_MAX, commandsIn, matchesCommand, validateCommand } from './feature-commands'
export type { CommandValidation, CommandValidationError, FeatureCommand } from './feature-commands'
export {
  KNOWLEDGE_SOURCES,
  describeKnowledge,
  filterKnowledgeEntries,
  knowledgeGlobMatches,
  knowledgeGlobSummary,
  knowledgeIn,
  validateKnowledge,
} from './feature-knowledge'
export type {
  FeatureKnowledge,
  KnowledgeEntry,
  KnowledgeSource,
  KnowledgeValidation,
  KnowledgeValidationError,
} from './feature-knowledge'
export type { FeatureHook, HookEvent, HookValidation, HookValidationError } from './feature-hooks'
export type {
  FeatureScript,
  ScriptRunOutcome,
  ScriptValidation,
  ScriptValidationError,
} from './feature-script'
export type {
  AppearanceAccent,
  AppearanceBackground,
  AppearanceValidation,
  AppearanceValidationError,
  FeatureAppearance,
  MarkdownCustomTheme,
  MarkdownReadTheme,
  MarkdownThemeFont,
} from './feature-appearance'
export {
  APPEARANCE_ACCENTS,
  APPEARANCE_BACKGROUNDS,
  MD_READ_THEMES,
  MD_THEME_FONTS,
  activeAppearance,
  contrastRatio,
  markdownCustomThemeVars,
  validateAppearance,
} from './feature-appearance'
export { OVERLAY_LIMITS, OVERLAY_TONES, validateOverlayShapes } from './slot-overlay'
export type { OverlayShape, OverlayTone } from './slot-overlay'
export {
  isAllowedViewUrl,
  viewHost,
  SLOT_FILL_FIELDS,
  SLOT_LIMITS,
  SLOT_PLACES,
  SLOT_SOURCES,
  slotsAt,
  validateSlot,
} from './feature-slots'
export type {
  SlotAction,
  SlotDeclaration,
  SlotListItem,
  SlotNode,
  SlotPlace,
  SlotSource,
  SlotValidation,
  SlotValidationError,
} from './feature-slots'
export type {
  FeatureDraft,
  FeaturePatch,
  FeatureRecord,
  FeatureRegistryDeps,
} from './features-skill'
export {
  AgentLoop,
  COMPLETED_VIA_TOOLS_TEXT,
  DEFAULT_MAX_TURNS,
  TOOL_ABORTED_OUTPUT,
  coerceArgumentFields,
  invalidArgumentFields,
  missingRequiredFields,
  runtimePreamble,
  sanitizeAgentPayload,
} from './loop'
export type {
  AgentLoopEvents,
  AgentLoopOptions,
  AgentRunResult,
  CompactionOptions,
  ToolExecutedEvent,
} from './loop'
export { createIpcTransport, IPC_STREAM_SILENCE_TIMEOUT_MS } from './electron-transport'
export { streamText } from './stream-text'
export type { StreamTextOptions, StreamTextOutcome } from './stream-text'
export type { IpcStreamChunk, IpcStreamStart, IpcTransportOptions } from './electron-transport'
