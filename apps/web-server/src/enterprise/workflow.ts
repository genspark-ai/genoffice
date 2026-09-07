/**
 * Workflow engine — create/list/get/update/delete/run. Placeholder
 * implementations; the real engine will land in Phase 3 (LUM-551).
 */
import { registerHandle, WORKFLOWS } from '../common/index.js'

export function registerWorkflowHandlers(): void {
  registerHandle('workflow:create', (_event: unknown, args: unknown) => {
    const { name, description, steps, triggers } = (args || {}) as {
      name: string
      description?: string
      steps: Array<{
        id: string
        type: string
        config: Record<string, unknown>
        next?: string
      }>
      triggers?: string[]
    }
    const id = `workflow-${Date.now()}`
    const workflowSteps = (steps || []).map(s => ({
      ...s,
      type: s.type as 'approval' | 'notification' | 'condition' | 'integration',
    }))
    WORKFLOWS.set(id, {
      id,
      tenantId: 'default',
      name,
      description: description || '',
      steps: workflowSteps,
      triggers: triggers || ['manual'],
      status: 'active',
      createdAt: Date.now(),
    })
    return { ok: true, id }
  })

  registerHandle('workflow:list', (_event: unknown, args: unknown) => {
    const { status, limit, offset } = (args || {}) as {
      status?: 'active' | 'paused' | 'archived'
      limit?: number
      offset?: number
    }
    const maxResults = limit || 50
    const startOffset = offset || 0
    let workflows = [...WORKFLOWS.values()]
    if (status) workflows = workflows.filter(w => w.status === status)
    return workflows
      .slice(startOffset, startOffset + maxResults)
      .map(w => ({
        id: w.id,
        name: w.name,
        description: w.description,
        stepCount: w.steps.length,
        triggers: w.triggers,
        status: w.status,
        createdAt: w.createdAt,
      }))
  })

  registerHandle('workflow:get', (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    return WORKFLOWS.get(id) || null
  })

  registerHandle('workflow:update', (_event: unknown, args: unknown) => {
    const { id, name, description, steps, status } = (args || {}) as {
      id: string
      name?: string
      description?: string
      steps?: Array<{
        id: string
        type: string
        config: Record<string, unknown>
        next?: string
      }>
      status?: 'active' | 'paused' | 'archived'
    }
    const workflow = WORKFLOWS.get(id)
    if (!workflow) return { ok: false, error: 'Workflow not found' }
    if (name) workflow.name = name
    if (description !== undefined) workflow.description = description
    if (steps) workflow.steps = steps.map(s => ({
      ...s,
      type: s.type as 'approval' | 'notification' | 'condition' | 'integration',
    }))
    if (status) workflow.status = status
    return { ok: true }
  })

  registerHandle('workflow:delete', (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    if (!WORKFLOWS.has(id)) return { ok: false, error: 'Workflow not found' }
    WORKFLOWS.delete(id)
    return { ok: true }
  })

  registerHandle('workflow:run', async (_event: unknown, args: unknown) => {
    const { id } = (args || {}) as { id: string; data?: Record<string, unknown> }
    const workflow = WORKFLOWS.get(id)
    if (!workflow) return { ok: false, error: 'Workflow not found' }
    if (workflow.status !== 'active') return { ok: false, error: 'Workflow is not active' }

    const executionId = `exec-${Date.now()}`
    return {
      ok: true,
      executionId,
      status: 'running',
      startedAt: Date.now(),
    }
  })
}
