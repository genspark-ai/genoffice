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
} from '../common/index.js'

export function registerProjectHandlers(): void {
  registerHandle('project:list', () => loadProjects())

  registerHandle('project:create', (_event: unknown, args: unknown) => {
    const { name } = args as { name: string }
    const projects = loadProjects()
    const project = {
      id: `proj-${Date.now()}`,
      name: name || '新项目',
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
    const project = projects.find(p => p.id === projectId)
    if (!project) return []

    return project.files.map(fileId => {
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
    }).filter(Boolean)
  })

  registerHandle('project:rename', (_event: unknown, args: unknown) => {
    const { id, name } = args as { id: string; name: string }
    const projects = loadProjects()
    const project = projects.find(p => p.id === id)
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
    let projects = loadProjects()
    const index = projects.findIndex(p => p.id === id)
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
    const project = projects.find(p => p.id === projectId)
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
    const project = projects.find(p => p.id === projectId)
    if (!project) return []
    return [
      { id: `event-${Date.now()}`, type: 'created', message: '项目已创建', timestamp: project.createdAt },
      { id: `event-${Date.now() + 1}`, type: 'updated', message: '项目已更新', timestamp: project.updatedAt },
    ]
  })
}
