/**
 * Project CRUD channels — list, create, files, rename, delete, move-file,
 * timeline. Persistence is delegated to the shared `state.ts` helpers so
 * the file format matches the legacy single-file implementation exactly.
 */
import { existsSync, statSync, unlinkSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import {
  FILES_DIR,
  MIME_TYPES,
  loadProjects,
  registerHandle,
  saveProjects,
} from '../common/index'

type ChatMessage = {
  role: 'user' | 'assistant'
  text: string
  tools?: unknown[]
  attachments?: unknown[]
  createdAt: number
}

const CHAT_HISTORY = new Map<string, ChatMessage[]>()

export function registerProjectHandlers(): void {
  registerHandle('project:resolveChat', (_event: unknown, args: unknown) => {
    const request = (args || {}) as { filePath?: string | null; tempChatId?: string }
    return {
      projectId: 'default',
      chatId:
        request.tempChatId ||
        (request.filePath
          ? `chat-${Buffer.from(request.filePath).toString('base64url').slice(0, 32)}`
          : `unsaved-${Date.now()}`),
    }
  })

  registerHandle('project:appendChat', (_event: unknown, args: unknown) => {
    const request = args as {
      projectId: string
      chatId: string
      role: 'user' | 'assistant'
      text: string
      tools?: unknown[]
      attachments?: unknown[]
    }
    const key = `${request.projectId}:${request.chatId}`
    const messages = CHAT_HISTORY.get(key) || []
    messages.push({
      role: request.role,
      text: request.text,
      tools: request.tools,
      attachments: request.attachments,
      createdAt: Date.now(),
    })
    CHAT_HISTORY.set(key, messages)
  })

  registerHandle('project:loadChat', (_event: unknown, args: unknown) => {
    const request = args as { projectId: string; chatId: string; limit?: number }
    const messages = CHAT_HISTORY.get(`${request.projectId}:${request.chatId}`) || []
    return messages.slice(-(request.limit || 200))
  })

  registerHandle('project:rebindChat', (_event: unknown, args: unknown) => {
    const request = args as { projectId: string; tempChatId: string; newChatId?: string }
    const chatId = request.newChatId || request.tempChatId
    const oldKey = `${request.projectId}:${request.tempChatId}`
    const newKey = `${request.projectId}:${chatId}`
    if (oldKey !== newKey && CHAT_HISTORY.has(oldKey)) {
      CHAT_HISTORY.set(newKey, CHAT_HISTORY.get(oldKey) || [])
      CHAT_HISTORY.delete(oldKey)
    }
    return { projectId: request.projectId, chatId }
  })

  registerHandle('project:list', () => loadProjects())

  registerHandle('project:create', (_event: unknown, args: unknown) => {
    const { name } = args as { name: string }
    const projects = loadProjects()
    /* Idempotent on name: when the caller doesn't specify a unique
     * name (e.g. test scripts that always create "Test Project"),
     * return the existing match instead of filling the sidebar with
     * duplicates. The first project ever created gets the well-known
     * default id "proj-default" so the seed path stays in sync. */
    const cleanName = (name || '新项目').trim()
    const existing = projects.find((p) => p.name === cleanName)
    if (existing) return existing
    const isFirst = projects.length === 0
    const project = {
      id: isFirst ? 'proj-default' : `proj-${Date.now()}`,
      name: cleanName,
      createdAt: Date.now(),
      updatedAt: Date.now(),
      files: [],
    }
    projects.push(project)
    saveProjects(projects)
    return project
  })

  registerHandle('project:files', (_event: unknown, args: unknown) => {
    const { projectId } = args as { projectId: string }
    const projects = loadProjects()
    const project = projects.find((p) => p.id === projectId)
    if (!project) return []

    return project.files
      .map((fileId) => {
        const filePath = join(FILES_DIR, fileId)
        if (existsSync(filePath)) {
          const stats = statSync(filePath)
          return {
            id: fileId,
            name: fileId.split('-').slice(1).join('-'),
            path: filePath,
            size: stats.size,
            mimeType: MIME_TYPES[extname(fileId)] || 'application/octet-stream',
            projectId,
            createdAt: stats.birthtimeMs,
          }
        }
        return null
      })
      .filter(Boolean)
  })

  registerHandle('project:rename', (_event: unknown, args: unknown) => {
    const { id, name } = args as { id: string; name: string }
    const projects = loadProjects()
    const project = projects.find((p) => p.id === id)
    if (project) {
      project.name = name
      project.updatedAt = Date.now()
      saveProjects(projects)
      return { ok: true }
    }
    return { ok: false, error: 'Project not found' }
  })

  registerHandle('project:delete', (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    const projects = loadProjects()
    const index = projects.findIndex((p) => p.id === id)
    if (index >= 0) {
      const project = projects[index]
      for (const fileId of project.files) {
        const filePath = join(FILES_DIR, fileId)
        if (existsSync(filePath)) unlinkSync(filePath)
      }
      projects.splice(index, 1)
      saveProjects(projects)
      return { ok: true }
    }
    return { ok: false, error: 'Project not found' }
  })

  registerHandle('project:moveFile', (_event: unknown, args: unknown) => {
    const { filePath, projectId } = args as { filePath: string; projectId: string }
    const projects = loadProjects()
    const project = projects.find((p) => p.id === projectId)
    if (project) {
      const fileId = basename(filePath)
      if (!project.files.includes(fileId)) {
        project.files.push(fileId)
        project.updatedAt = Date.now()
        saveProjects(projects)
      }
      return { ok: true }
    }
    return { ok: false, error: 'Project not found' }
  })

  registerHandle('project:timeline', (_event: unknown, args: unknown) => {
    const { projectId } = args as { projectId: string }
    const projects = loadProjects()
    const project = projects.find((p) => p.id === projectId)
    if (!project) return []
    return [
      {
        id: `event-${Date.now()}`,
        type: 'created',
        message: '项目已创建',
        timestamp: project.createdAt,
      },
      {
        id: `event-${Date.now() + 1}`,
        type: 'updated',
        message: '项目已更新',
        timestamp: project.updatedAt,
      },
    ]
  })
}
