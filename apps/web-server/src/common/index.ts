/**
 * Public surface of the shared `common` module.
 *
 * Each capability directory imports the bits it needs from here so the
 * split stays a pure refactor (no behaviour change, no handler lost).
 */
export { BYTES_TAG, encodeTransportValue, decodeTransportValue } from './codec'
export type { TypedArrayTag } from './codec'

export { registerHandle, getHandler, listChannels, handlerCount } from './registry'
export type { IpcHandler } from './registry'

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
  initRecentState,
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
} from './state'
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
} from './state'

export { MIME_TYPES } from './mime'
export { ROOT, PORT, HOST, APPS, STATIC_ROOT, WEB_TEMP_ROOT } from './paths'
