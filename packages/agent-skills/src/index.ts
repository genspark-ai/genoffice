/**
 * @genoffice/agent-skills — public surface
 */

export {
  createDocsSkillExtension,
  createReadBlocksTool,
  createGetDocumentContextTool,
  createInsertContentTool,
  createReplaceBlocksTool,
  createReplaceSelectionTool,
  createApplyOpsTool,
  createCreateDocumentTool,
  createReplaceDocumentTool,
  createReadCommentsTool,
  createReplyCommentTool,
  createResolveCommentTool,
  ALL_DOCS_TOOL_NAMES,
  type DocsSkillOptions,
  type DocsToolName,
  type DocsEditor,
  type DocsBlock,
  type ReadBlocksToolOptions,
  type CommentThread,
} from "./extensions/docs-skill";

export {
  createSheetsSkillExtension,
  createGetWorkbookContextTool,
  createReadRangeTool,
  createAggregateRangeTool,
  createFindCellsTool,
  createCreateSheetTool,
  ALL_SHEETS_TOOL_NAMES,
  type SheetsSkillOptions,
  type SheetsToolName,
  type SheetsEditor,
  type SheetsRange,
  type CellValue,
  type WorkbookSummary,
} from "./extensions/sheets-skill";

export {
  createSlidesSkillExtension,
  createReadSlideTool,
  createPlanDeckTool,
  createExecuteSlideScriptTool,
  createRegenerateSlideTool,
  ALL_SLIDES_TOOL_NAMES,
  type SlidesSkillOptions,
  type SlidesToolName,
  type SlidesEditor,
  type SlideSummary,
  type SlideContent,
  type DeckPlan,
  type SlideScript,
} from "./extensions/slides-skill";
export {
  createFrozenSelectionExtension,
  type FrozenSelection,
  type FrozenSelectionEditor,
  type FrozenSelectionOptions,
} from "./extensions/frozen-selection";

export {
  createVerifyResponseExtension,
  DEFAULT_VERIFY_RULES,
  VERIFY_BLOCK_MARKER,
  type VerifyResponseOptions,
} from "./extensions/verify-response";

export {
  installOfficeSafety,
  type OfficeSafetyOptions,
} from "./extensions/office-safety";

export {
  createOfficeWorkflowTool,
  installOfficeWorkflow,
  CrossOfficeWorkflowParams,
  WORKFLOW_OUTPUT_FORMATS,
  type CrossOfficeWorkflowArgs,
  type CrossOfficeWorkflowDetails,
  type InstallOfficeWorkflowOptions,
  type OfficeWorkflowCallbacks,
  type OfficeWorkflowOptions,
  type ComposeResult,
  type SpreadsheetReadResult,
  type WorkflowOutputFormat,
  type WorkflowRow,
} from "./extensions/office-workflow";

export {
  createRequestReviewTool,
  installAgentTeam,
  RequestReviewParams,
  BUILTIN_AGENT_ROLES,
  type AgentRole,
  type AgentTeamOptions,
  type InstallAgentTeamOptions,
  type RequestReviewArgs,
  type RequestReviewDetails,
} from "./extensions/agent-team";

export {
  installAuditLog,
  InMemoryAuditSink,
  JsonlAuditSink,
  CompositeAuditSink,
  redact,
  DEFAULT_REDACT_KEYS,
  DEFAULT_AUDIT_DIRECTORY,
  DEFAULT_AUDIT_FILE,
  type AuditLogEntry,
  type AuditLogOptions,
  type AuditSink,
  type JsonlAuditSinkOptions,
} from "./extensions/audit-log";

export {
  createOllamaProvider,
  installLocalModels,
  OLLAMA_API,
  OLLAMA_DEFAULT_BASE_URL,
  OLLAMA_DEFAULT_MODEL_ID,
  type InstallLocalModelsOptions,
  type OllamaProviderOptions,
} from "./extensions/local-models";

export {
  createSkillMarket,
  DEFAULT_SKILLS_DIRECTORY,
  type InstallRecord,
  type SkillMarket,
  type SkillMarketEntry,
  type SkillMarketFileSystem,
  type SkillMarketOptions,
} from "./extensions/skill-market";

export {
  createWebSearchExtension,
  webSearchExtensionDefaults,
} from "./extensions/web-search-skill";

export {
  createImageSearchExtension,
  imageSearchExtensionDefaults,
} from "./extensions/image-search-skill";

export {
  parseDuckDuckGo,
  type SearchHit,
} from "./extensions/web-search-skill";

export {
  parseDuckDuckGoImages,
  type ImageHit,
} from "./extensions/image-search-skill";
