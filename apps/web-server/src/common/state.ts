/**
 * Module-level singletons shared across capability domains.
 *
 * The web-server keeps in-memory state per process (no real DB). Each
 * capability module imports what it needs from here rather than re-declaring
 * Maps, so behaviour matches the legacy single-file implementation exactly.
 */

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

export const DATA_DIR = process.env.DATA_DIR || '/tmp/genoffice-data'
mkdirSync(DATA_DIR, { recursive: true })

export const FILES_DIR = join(DATA_DIR, 'files')
mkdirSync(FILES_DIR, { recursive: true })

// ----- Project persistence --------------------------------------------------
export const PROJECTS_FILE = join(DATA_DIR, 'projects.json')

export interface Project {
  id: string
  name: string
  createdAt: number
  updatedAt: number
  files: string[]
}

export function loadProjects(): Project[] {
  try {
    if (existsSync(PROJECTS_FILE)) {
      return JSON.parse(readFileSync(PROJECTS_FILE, 'utf-8'))
    }
  } catch {}
  return []
}

export function saveProjects(projects: Project[]): void {
  writeFileSync(PROJECTS_FILE, JSON.stringify(projects, null, 2))
}

// ----- File index -----------------------------------------------------------
export interface FileInfo {
  id: string
  name: string
  path: string
  size: number
  mimeType: string
  createdAt: number
  updatedAt: number
}

export const FILES_INDEX: Map<string, FileInfo> = new Map()

// ----- Recent file lists ---------------------------------------------------
export interface DocInfo {
  id: string
  path: string
  name: string
  openedAt: number
  modified: boolean
}

export const DOCS_RECENT_FILE = join(DATA_DIR, 'docs-recent.json')

export function loadRecentDocs(): DocInfo[] {
  try {
    if (existsSync(DOCS_RECENT_FILE)) {
      return JSON.parse(readFileSync(DOCS_RECENT_FILE, 'utf-8'))
    }
  } catch {}
  return []
}

export function saveRecentDocs(docs: DocInfo[]): void {
  writeFileSync(DOCS_RECENT_FILE, JSON.stringify(docs.slice(0, 10), null, 2))
}

export const DOCS_RECENT: Map<string, DocInfo> = new Map()
export const DOCS_STARRED: Set<string> = new Set()

export interface SheetInfo {
  id: string
  path: string
  name: string
  openedAt: number
}

export const SHEETS_RECENT_FILE = join(DATA_DIR, 'sheets-recent.json')

export function loadRecentSheets(): SheetInfo[] {
  try {
    if (existsSync(SHEETS_RECENT_FILE)) {
      return JSON.parse(readFileSync(SHEETS_RECENT_FILE, 'utf-8'))
    }
  } catch {}
  return []
}

export function saveRecentSheets(sheets: SheetInfo[]): void {
  writeFileSync(SHEETS_RECENT_FILE, JSON.stringify(sheets.slice(0, 10), null, 2))
}

export interface SlideInfo {
  id: string
  path: string
  name: string
  openedAt: number
}

export const SLIDES_RECENT_FILE = join(DATA_DIR, 'slides-recent.json')

export function loadRecentSlides(): SlideInfo[] {
  try {
    if (existsSync(SLIDES_RECENT_FILE)) {
      return JSON.parse(readFileSync(SLIDES_RECENT_FILE, 'utf-8'))
    }
  } catch {}
  return []
}

export function saveRecentSlides(slides: SlideInfo[]): void {
  writeFileSync(SLIDES_RECENT_FILE, JSON.stringify(slides.slice(0, 10), null, 2))
}

// ----- AI streaming state ---------------------------------------------------
export const AI_STREAMS: Map<string, {
  chunks: string[]
  abort: AbortController
}> = new Map()

export const ACTIVE_STREAMS: Map<string, {
  controller: ReadableStreamDefaultController
  aborted: boolean
}> = new Map()

// ----- Collab state ---------------------------------------------------------
export const COLLAB_SESSIONS: Map<string, {
  docId: string
  users: Set<string>
  lastActivity: number
  locks: Map<string, { userId: string; timestamp: number }>
  cursors: Map<string, { position: { x: number; y: number; offset: number }; selection?: { start: number; end: number }; timestamp: number }>
  changes: Array<{ id: string; docId: string; userId: string; change: unknown; timestamp: number; version: number }>
}> = new Map()

export const PRESENCE: Map<string, Map<string, {
  userId: string
  userName: string
  status: 'active' | 'idle' | 'away'
  lastSeen: number
  cursor?: { x: number; y: number; selection?: { start: number; end: number } }
  color: string
}>> = new Map()

export const DOC_PERMISSIONS: Map<string, Map<string, string>> = new Map()

export interface DocVersion {
  id: string
  content: string
  timestamp: number
  userId: string
  message?: string
}

export interface DocVersionHistory {
  docId: string
  versions: DocVersion[]
}

export const DOC_VERSIONS: Map<string, DocVersionHistory> = new Map()

export interface CommentReply {
  id: string
  userId: string
  userName: string
  content: string
  timestamp: number
}

export interface DocComment {
  id: string
  userId: string
  userName: string
  content: string
  timestamp: number
  resolved: boolean
  replies: CommentReply[]
  selection?: { start: number; end: number; text: string }
}

export const DOC_COMMENTS: Map<string, DocComment[]> = new Map()

export interface DocTemplate {
  id: string
  name: string
  type: 'docs' | 'sheets' | 'slides'
  content: string
  thumbnail?: string
  category: string
  tags: string[]
  createdAt: number
  updatedAt: number
}

export const TEMPLATES: Map<string, DocTemplate> = new Map()

export function initDefaultTemplates(): void {
  if (TEMPLATES.size > 0) return
  const defaultTemplates = [
    {
      id: 'tpl-resume',
      name: '简历',
      type: 'docs' as const,
      content: '<h1>个人简历</h1>',
      category: '办公',
      tags: ['简历', '个人'],
    },
    {
      id: 'tpl-report',
      name: '工作报告',
      type: 'docs' as const,
      content: '<h1>工作报告</h1>',
      category: '办公',
      tags: ['报告', '工作'],
    },
    {
      id: 'tpl-presentation',
      name: '商务演示',
      type: 'slides' as const,
      content: '[]',
      category: '演示',
      tags: ['演示', '商务'],
    },
  ]

  for (const tpl of defaultTemplates) {
    TEMPLATES.set(tpl.id, {
      ...tpl,
      createdAt: Date.now(),
      updatedAt: Date.now(),
    })
  }
}

// ----- Enterprise / shell state --------------------------------------------
export const CLOUD_FILES: Map<string, {
  id: string
  name: string
  size: number
  type: string
  url: string
  createdAt: number
  updatedAt: number
  public: boolean
}> = new Map()

export const OFFLINE_QUEUE: Map<string, {
  id: string
  action: string
  payload: unknown
  timestamp: number
  synced: boolean
}> = new Map()

export const SEARCH_INDEX: Map<string, {
  id: string
  type: string
  title: string
  content: string
  tags: string[]
  createdAt: number
}> = new Map()

export interface UserRecord {
  id: string
  name: string
  email: string
  role: 'admin' | 'editor' | 'viewer'
  createdAt: number
}

export const USERS: Map<string, UserRecord> = new Map()

export const PERMISSIONS: Map<string, Map<string, string[]>> = new Map()

export interface TenantRecord {
  id: string
  name: string
  domain: string
  plan: 'free' | 'pro' | 'enterprise'
  settings: Record<string, unknown>
  createdAt: number
  status: 'active' | 'suspended' | 'trial'
}

export const TENANTS: Map<string, TenantRecord> = new Map()

export interface MailRecord {
  id: string
  tenantId: string
  from: { name: string; email: string }
  to: Array<{ name: string; email: string }>
  subject: string
  body: string
  attachments: Array<{ name: string; size: number }>
  sentAt: number
  status: 'sent' | 'failed' | 'pending'
}

export const MAILS: Map<string, MailRecord> = new Map()

export interface CalendarEventRecord {
  id: string
  tenantId: string
  title: string
  description: string
  startTime: number
  endTime: number
  attendees: Array<{ name: string; email: string; status: 'pending' | 'accepted' | 'declined' }>
  location?: string
  reminders: number[]
  recurrence?: string
  status: 'confirmed' | 'cancelled' | 'tentative'
}

export const CALENDARS: Map<string, CalendarEventRecord> = new Map()

export interface WorkflowRecord {
  id: string
  tenantId: string
  name: string
  description: string
  steps: Array<{
    id: string
    type: 'approval' | 'notification' | 'condition' | 'integration'
    config: Record<string, unknown>
    next?: string
  }>
  triggers: string[]
  status: 'active' | 'paused' | 'archived'
  createdAt: number
}

export const WORKFLOWS: Map<string, WorkflowRecord> = new Map()

export interface AuditRecord {
  id: string
  tenantId: string
  userId: string
  action: string
  resource: string
  resourceId: string
  details: Record<string, unknown>
  ip: string
  userAgent: string
  timestamp: number
  status: 'success' | 'failure'
}

export const AUDIT_LOGS: Map<string, AuditRecord> = new Map()

export interface NotificationRecord {
  id: string
  type: 'info' | 'success' | 'warning' | 'error'
  title: string
  message: string
  timestamp: number
  read: boolean
}

export const NOTIFICATIONS: Map<string, NotificationRecord[]> = new Map()

export const WEB_WINDOWS: Map<string, { url: string; name: string }> = new Map()

export interface TabRecord {
  id: string
  type: string
  title: string
  path?: string
}

export const TABS: Map<string, TabRecord> = new Map()
