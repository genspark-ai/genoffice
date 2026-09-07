/**
 * Collab supporting channels — version history, comments, templates.
 */
import { DOC_COMMENTS, DOC_VERSIONS, registerHandle, TEMPLATES, initDefaultTemplates } from '../common/index.js'

export function registerHistoryHandlers(): void {
  registerHandle('history:versions', (_event: unknown, args: unknown) => {
    const { docId } = args as { docId: string }
    const docVersions = DOC_VERSIONS.get(docId)
    if (!docVersions) return []
    return docVersions.versions.map(v => ({
      id: v.id,
      timestamp: v.timestamp,
      userId: v.userId,
      message: v.message || '自动保存',
    }))
  })

  registerHandle('history:create-version', (_event: unknown, args: unknown) => {
    const { docId, content, userId, message } = args as {
      docId: string
      content: string
      userId: string
      message?: string
    }

    if (!DOC_VERSIONS.has(docId)) {
      DOC_VERSIONS.set(docId, { docId, versions: [] })
    }

    const docVersions = DOC_VERSIONS.get(docId)!
    const versionId = `v-${Date.now()}`

    docVersions.versions.push({
      id: versionId,
      content,
      timestamp: Date.now(),
      userId,
      message,
    })

    if (docVersions.versions.length > 50) {
      docVersions.versions = docVersions.versions.slice(-50)
    }

    return { ok: true, versionId, timestamp: Date.now() }
  })

  registerHandle('history:get-version', (_event: unknown, args: unknown) => {
    const { docId, versionId } = args as { docId: string; versionId: string }
    const docVersions = DOC_VERSIONS.get(docId)
    if (!docVersions) return null

    const version = docVersions.versions.find(v => v.id === versionId)
    return version || null
  })

  registerHandle('history:restore-version', (_event: unknown, args: unknown) => {
    const { docId, versionId } = args as { docId: string; versionId: string }
    const docVersions = DOC_VERSIONS.get(docId)
    if (!docVersions) return { ok: false, error: 'Document not found' }

    const version = docVersions.versions.find(v => v.id === versionId)
    if (!version) return { ok: false, error: 'Version not found' }

    return { ok: true, content: version.content }
  })
}

export function registerCommentHandlers(): void {
  registerHandle('comments:list', (_event: unknown, args: unknown) => {
    const { docId } = args as { docId: string }
    return DOC_COMMENTS.get(docId) || []
  })

  registerHandle('comments:add', (_event: unknown, args: unknown) => {
    const { docId, userId, userName, content, selection } = args as {
      docId: string
      userId: string
      userName: string
      content: string
      selection?: { start: number; end: number; text: string }
    }

    if (!DOC_COMMENTS.has(docId)) {
      DOC_COMMENTS.set(docId, [])
    }

    const comments = DOC_COMMENTS.get(docId)!
    const commentId = `comment-${Date.now()}`

    comments.push({
      id: commentId,
      userId,
      userName,
      content,
      timestamp: Date.now(),
      resolved: false,
      replies: [],
      selection,
    })

    return { ok: true, commentId }
  })

  registerHandle('comments:reply', (_event: unknown, args: unknown) => {
    const { docId, commentId, userId, userName, content } = args as {
      docId: string
      commentId: string
      userId: string
      userName: string
      content: string
    }

    const comments = DOC_COMMENTS.get(docId)
    if (!comments) return { ok: false, error: 'Document not found' }

    const comment = comments.find(c => c.id === commentId)
    if (!comment) return { ok: false, error: 'Comment not found' }

    const replyId = `reply-${Date.now()}`
    comment.replies.push({
      id: replyId,
      userId,
      userName,
      content,
      timestamp: Date.now(),
    })

    return { ok: true, replyId }
  })

  registerHandle('comments:resolve', (_event: unknown, args: unknown) => {
    const { docId, commentId } = args as { docId: string; commentId: string }
    const comments = DOC_COMMENTS.get(docId)
    if (!comments) return { ok: false, error: 'Document not found' }

    const comment = comments.find(c => c.id === commentId)
    if (!comment) return { ok: false, error: 'Comment not found' }

    comment.resolved = true
    return { ok: true }
  })

  registerHandle('comments:delete', (_event: unknown, args: unknown) => {
    const { docId, commentId } = args as { docId: string; commentId: string }
    const comments = DOC_COMMENTS.get(docId)
    if (!comments) return { ok: false, error: 'Document not found' }

    const index = comments.findIndex(c => c.id === commentId)
    if (index === -1) return { ok: false, error: 'Comment not found' }

    comments.splice(index, 1)
    return { ok: true }
  })
}

export function registerTemplateHandlers(): void {
  initDefaultTemplates()

  registerHandle('templates:list', (_event: unknown, args: unknown) => {
    const { type, category, search } = (args || {}) as { type?: string; category?: string; search?: string }

    let templates = [...TEMPLATES.values()]

    if (type) {
      templates = templates.filter(t => t.type === type)
    }
    if (category) {
      templates = templates.filter(t => t.category === category)
    }
    if (search) {
      const searchLower = search.toLowerCase()
      templates = templates.filter(t =>
        t.name.toLowerCase().includes(searchLower) ||
        t.tags.some(tag => tag.toLowerCase().includes(searchLower)),
      )
    }

    return templates.map(t => ({
      id: t.id,
      name: t.name,
      type: t.type,
      thumbnail: t.thumbnail,
      category: t.category,
      tags: t.tags,
      createdAt: t.createdAt,
      updatedAt: t.updatedAt,
    }))
  })

  registerHandle('templates:get', (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    return TEMPLATES.get(id) || null
  })

  registerHandle('templates:create', (_event: unknown, args: unknown) => {
    const { name, type, content, category, tags } = args as {
      name: string
      type: 'docs' | 'sheets' | 'slides'
      content: string
      category?: string
      tags?: string[]
    }

    const id = `tpl-${Date.now()}`
    TEMPLATES.set(id, {
      id,
      name,
      type,
      content,
      category: category || '自定义',
      tags: tags || [],
      createdAt: Date.now(),
      updatedAt: Date.now(),
    })

    return { ok: true, id }
  })

  registerHandle('templates:delete', (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    if (!TEMPLATES.has(id)) return { ok: false, error: 'Template not found' }
    TEMPLATES.delete(id)
    return { ok: true }
  })
}
