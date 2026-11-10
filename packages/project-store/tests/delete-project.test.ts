import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { existsSync, mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { ProjectStore } from '../src/store.js'

describe('deleteProject keeps chat history reachable', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-delete-'))
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  const chatsDir = (projectId: string) => join(tmpDir, 'projects', projectId, 'chats')

  it('moves the transcript to the default project the file reverts to', () => {
    const filePath = join(tmpDir, 'plan.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    const project = store.createProject('Research')
    store.moveFileToProject(filePath, project.id)

    const ids = store.resolveChatForFile(filePath)
    expect(ids.projectId).toBe(project.id)
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'user', text: 'question' })
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'assistant', text: 'answer' })

    store.deleteProject(project.id)

    const after = store.resolveChatForFile(filePath)
    expect(after.projectId).toBe('default')
    expect(after.chatId).toBe(ids.chatId)
    expect(store.loadChat('default', ids.chatId).map((m) => m.text)).toEqual(['question', 'answer'])
  })

  it('leaves no transcript behind in the trashed project', () => {
    const filePath = join(tmpDir, 'plan.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    const project = store.createProject('Research')
    store.moveFileToProject(filePath, project.id)
    const ids = store.resolveChatForFile(filePath)
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'user', text: 'q' })
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'assistant', text: 'a' })

    store.deleteProject(project.id)

    const trash = join(tmpDir, 'projects', '.trash')
    const trashed = readdirSync(trash).map((d) => join(trash, d, 'chats'))
    for (const dir of trashed) {
      if (existsSync(dir)) {
        expect(readdirSync(dir)).not.toContain(`${ids.chatId}.jsonl`)
      }
    }
    expect(existsSync(join(chatsDir('default'), `${ids.chatId}.jsonl`))).toBe(true)
  })

  it('merges into an existing default transcript, keeping both in order', () => {
    const filePath = join(tmpDir, 'plan.docx')
    writeFileSync(filePath, 'doc', 'utf8')

    // History that already belongs to the default project
    const early = store.resolveChatForFile(filePath)
    store.appendChatMessage(early.projectId, early.chatId, { role: 'user', text: 'default q' })
    store.appendChatMessage(early.projectId, early.chatId, { role: 'assistant', text: 'default a' })

    const project = store.createProject('Research')
    store.moveFileToProject(filePath, project.id)
    const inProject = store.resolveChatForFile(filePath)
    store.appendChatMessage(inProject.projectId, inProject.chatId, {
      role: 'user',
      text: 'project q',
    })
    store.appendChatMessage(inProject.projectId, inProject.chatId, {
      role: 'assistant',
      text: 'project a',
    })

    store.deleteProject(project.id)

    const msgs = store.loadChat('default', early.chatId)
    expect(msgs.map((m) => m.text)).toEqual(['default q', 'default a', 'project q', 'project a'])
    expect(msgs.map((m) => m.seq)).toEqual([0, 1, 2, 3])
  })

  it('appending after the deletion continues the migrated transcript', () => {
    const filePath = join(tmpDir, 'plan.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    const project = store.createProject('Research')
    store.moveFileToProject(filePath, project.id)
    const ids = store.resolveChatForFile(filePath)
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'user', text: 'q' })
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'assistant', text: 'a' })

    store.deleteProject(project.id)

    store.appendChatMessage('default', ids.chatId, { role: 'user', text: 'after delete' })
    store.appendChatMessage('default', ids.chatId, { role: 'assistant', text: 'reply' })
    expect(store.loadChat('default', ids.chatId).map((m) => m.seq)).toEqual([0, 1, 2, 3])
  })

  it('materializes buffered opening messages before migrating them', () => {
    const filePath = join(tmpDir, 'plan.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    const project = store.createProject('Research')
    store.moveFileToProject(filePath, project.id)
    const ids = store.resolveChatForFile(filePath)
    // Only a user message: nothing is on disk yet
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'user', text: 'interrupted' })
    expect(existsSync(join(chatsDir(project.id), `${ids.chatId}.jsonl`))).toBe(false)

    store.deleteProject(project.id)

    expect(store.loadChat('default', ids.chatId).map((m) => m.text)).toEqual(['interrupted'])
  })

  it('migrates every file of the project, each keeping its own transcript', () => {
    const project = store.createProject('Research')
    const first = join(tmpDir, 'first.docx')
    const second = join(tmpDir, 'second.docx')
    writeFileSync(first, 'a', 'utf8')
    writeFileSync(second, 'b', 'utf8')
    store.moveFileToProject(first, project.id)
    store.moveFileToProject(second, project.id)

    const a = store.resolveChatForFile(first)
    store.appendChatMessage(a.projectId, a.chatId, { role: 'user', text: 'a-q' })
    store.appendChatMessage(a.projectId, a.chatId, { role: 'assistant', text: 'a-a' })
    const b = store.resolveChatForFile(second)
    store.appendChatMessage(b.projectId, b.chatId, { role: 'user', text: 'b-q' })
    store.appendChatMessage(b.projectId, b.chatId, { role: 'assistant', text: 'b-a' })

    store.deleteProject(project.id)

    expect(store.loadChat('default', a.chatId).map((m) => m.text)).toEqual(['a-q', 'a-a'])
    expect(store.loadChat('default', b.chatId).map((m) => m.text)).toEqual(['b-q', 'b-a'])
    expect(a.chatId).not.toBe(b.chatId)
  })

  it('a project whose files have no chat still deletes cleanly', () => {
    const project = store.createProject('Empty Chat')
    const filePath = join(tmpDir, 'nochat.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    store.moveFileToProject(filePath, project.id)
    store.resolveProjectForFile(filePath)

    expect(() => store.deleteProject(project.id)).not.toThrow()
    expect(store.resolveProjectForFile(filePath)).toBe('default')
    expect(store.listProjects().find((p) => p.id === project.id)).toBeUndefined()
  })

  it('a project with no files at all still deletes cleanly', () => {
    const project = store.createProject('No Files')
    expect(() => store.deleteProject(project.id)).not.toThrow()
    expect(store.listProjects().find((p) => p.id === project.id)).toBeUndefined()
    expect(store.getProject('default')?.files).toEqual([])
  })

  it('the file is listed by the default project after the deletion', () => {
    const filePath = join(tmpDir, 'plan.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    const project = store.createProject('Research')
    store.moveFileToProject(filePath, project.id)
    store.resolveChatForFile(filePath)

    store.deleteProject(project.id)

    expect(store.listProjectFiles('default')).toEqual([filePath])
    expect(store.listProjectsSummary().find((s) => s.id === 'default')?.fileCount).toBe(1)
  })
})
