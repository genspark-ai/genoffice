/**
 * Enterprise user + tenant CRUD — list/get/create/update/delete. The maps
 * `USERS` and `TENANTS` live in `common/state.ts` so other modules (auth,
 * audit) can read them.
 */
import { registerHandle, TENANTS, USERS } from '../common/index.js'

export function registerUserHandlers(): void {
  registerHandle('users:list', (_event: unknown) => {
    return [...USERS.values()].map(u => ({
      id: u.id,
      name: u.name,
      email: u.email,
      role: u.role,
    }))
  })

  registerHandle('users:get', (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    return USERS.get(id) || null
  })

  registerHandle('users:create', (_event: unknown, args: unknown) => {
    const { name, email, role } = args as { name: string; email: string; role?: string }
    const id = `user-${Date.now()}`
    USERS.set(id, {
      id,
      name,
      email,
      role: (role as 'admin' | 'editor' | 'viewer') || 'viewer',
      createdAt: Date.now(),
    })
    return { ok: true, id }
  })

  registerHandle('users:update', (_event: unknown, args: unknown) => {
    const { id, name, email, role } = args as { id: string; name?: string; email?: string; role?: string }
    const user = USERS.get(id)
    if (!user) return { ok: false, error: 'User not found' }
    if (name) user.name = name
    if (email) user.email = email
    if (role) user.role = role as 'admin' | 'editor' | 'viewer'
    return { ok: true }
  })

  registerHandle('users:delete', (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    if (!USERS.has(id)) return { ok: false, error: 'User not found' }
    USERS.delete(id)
    return { ok: true }
  })
}

export function registerTenantHandlers(): void {
  registerHandle('tenant:list', () => {
    return [...TENANTS.values()].map(t => ({
      id: t.id,
      name: t.name,
      domain: t.domain,
      plan: t.plan,
      status: t.status,
    }))
  })

  registerHandle('tenant:create', (_event: unknown, args: unknown) => {
    const { name, domain, plan } = (args || {}) as {
      name: string
      domain: string
      plan?: 'free' | 'pro' | 'enterprise'
    }
    const id = `tenant-${Date.now()}`
    TENANTS.set(id, {
      id,
      name,
      domain,
      plan: plan || 'free',
      settings: {},
      createdAt: Date.now(),
      status: 'trial',
    })
    return { ok: true, id }
  })

  registerHandle('tenant:get', (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    return TENANTS.get(id) || null
  })

  registerHandle('tenant:update', (_event: unknown, args: unknown) => {
    const { id, name, plan, settings, status } = (args || {}) as {
      id: string
      name?: string
      plan?: 'free' | 'pro' | 'enterprise'
      settings?: Record<string, unknown>
      status?: 'active' | 'suspended' | 'trial'
    }
    const tenant = TENANTS.get(id)
    if (!tenant) return { ok: false, error: 'Tenant not found' }
    if (name) tenant.name = name
    if (plan) tenant.plan = plan
    if (settings) tenant.settings = { ...tenant.settings, ...settings }
    if (status) tenant.status = status
    return { ok: true }
  })
}
