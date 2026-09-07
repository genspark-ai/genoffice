/**
 * SSO/OIDC auth + audit log channels. Placeholder implementations; Phase 3
 * will introduce a real identity provider and persisted audit log.
 */
import { AUDIT_LOGS, registerHandle } from '../common/index.js'

export function registerAuthHandlers(): void {
  registerHandle('auth:sso-login', (_event: unknown, args: unknown) => {
    const { provider, redirectUri } = (args || {}) as {
      provider: string
      redirectUri?: string
    }
    return {
      authUrl: `https://sso.genoffice.ai/authorize?provider=${provider}&redirect_uri=${redirectUri || ''}`,
      state: `state-${Date.now()}`,
    }
  })

  registerHandle('auth:sso-callback', async (_event: unknown, args: unknown) => {
    const { code, state } = (args || {}) as { code: string; state: string }
    return {
      ok: true,
      accessToken: `token-${Date.now()}`,
      refreshToken: `refresh-${Date.now()}`,
      expiresIn: 3600,
      user: {
        id: `user-${Date.now()}`,
        email: 'user@example.com',
        name: 'SSO User',
      },
    }
  })

  registerHandle('auth:logout', (_event: unknown) => ({
    ok: true,
    redirectUrl: '/',
  }))
}

export function registerAuditHandlers(): void {
  registerHandle('audit:log', (_event: unknown, args: unknown) => {
    const { action, resource, resourceId, details, status } = (args || {}) as {
      action: string
      resource: string
      resourceId?: string
      details?: Record<string, unknown>
      status?: 'success' | 'failure'
    }
    const id = `audit-${Date.now()}`
    AUDIT_LOGS.set(id, {
      id,
      tenantId: 'default',
      userId: 'system',
      action,
      resource,
      resourceId: resourceId || '',
      details: details || {},
      ip: '0.0.0.0',
      userAgent: 'GenOffice/1.0',
      timestamp: Date.now(),
      status: status || 'success',
    })
    return { ok: true, id }
  })

  registerHandle('audit:query', (_event: unknown, args: unknown) => {
    const { userId, action, resource, startDate, endDate, limit, offset } = (args || {}) as {
      userId?: string
      action?: string
      resource?: string
      startDate?: number
      endDate?: number
      limit?: number
      offset?: number
    }
    const maxResults = limit || 100
    const startOffset = offset || 0
    let logs = [...AUDIT_LOGS.values()]
    if (userId) logs = logs.filter(l => l.userId === userId)
    if (action) logs = logs.filter(l => l.action.includes(action))
    if (resource) logs = logs.filter(l => l.resource === resource)
    if (startDate) logs = logs.filter(l => l.timestamp >= startDate)
    if (endDate) logs = logs.filter(l => l.timestamp <= endDate)
    return {
      logs: logs
        .sort((a, b) => b.timestamp - a.timestamp)
        .slice(startOffset, startOffset + maxResults),
      total: logs.length,
    }
  })

  registerHandle('audit:export', (_event: unknown, args: unknown) => {
    const { format, startDate, endDate } = (args || {}) as {
      format: 'csv' | 'json' | 'xlsx'
      startDate?: number
      endDate?: number
    }
    let logs = [...AUDIT_LOGS.values()]
    if (startDate) logs = logs.filter(l => l.timestamp >= startDate)
    if (endDate) logs = logs.filter(l => l.timestamp <= endDate)

    const exportId = `export-${Date.now()}`
    return {
      exportId,
      format,
      recordCount: logs.length,
      downloadUrl: `/audit/exports/${exportId}.${format}`,
    }
  })
}
