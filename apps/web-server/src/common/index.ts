/**
 * Public surface of the shared `common` module.
 *
 * Each capability directory imports the bits it needs from here so the
 * split stays a pure refactor (no behaviour change, no handler lost).
 */
export { BYTES_TAG, encodeTransportValue, decodeTransportValue } from './codec.js'
export type { TypedArrayTag } from './codec.js'

export { registerHandle, getHandler, listChannels, handlerCount } from './registry.js'
export type { IpcHandler } from './registry.js'

export {
  DATA_DIR,
  FILES_DIR,
  PROJECTS_FILE,
  loadProjects,
  saveProjects,
  FILES_INDEX,
  DOCS_RECENT_FILE,
  loadRecentDocs,
  saveRecentDocs,
  DOCS_RECENT,
  DOCS_STARRED,
  SHEETS_RECENT_FILE,
  loadRecentSheets,
  saveRecentSheets,
  SLIDES_RECENT_FILE,
  loadRecentSlides,
  saveRecentSlides,
  AI_STREAMS,
  ACTIVE_STREAMS,
  COLLAB_SESSIONS,
  PRESENCE,
  DOC_PERMISSIONS,
  DOC_VERSIONS,
  DOC_COMMENTS,
  TEMPLATES,
  initDefaultTemplates,
  CLOUD_FILES,
  OFFLINE_QUEUE,
  SEARCH_INDEX,
  USERS,
  PERMISSIONS,
  TENANTS,
  MAILS,
  CALENDARS,
  WORKFLOWS,
  AUDIT_LOGS,
  NOTIFICATIONS,
  WEB_WINDOWS,
  TABS,
} from './state.js'
export type {
  Project,
  FileInfo,
  DocInfo,
  SheetInfo,
  SlideInfo,
  DocVersion,
  DocVersionHistory,
  DocComment,
  CommentReply,
  DocTemplate,
  UserRecord,
  TenantRecord,
  MailRecord,
  CalendarEventRecord,
  WorkflowRecord,
  AuditRecord,
  NotificationRecord,
  TabRecord,
} from './state.js'

export { MIME_TYPES } from './mime.js'
export { ROOT, PORT, HOST, APPS, STATIC_ROOT, WEB_TEMP_ROOT } from './paths.js'
