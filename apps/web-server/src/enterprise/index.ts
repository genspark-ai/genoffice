/**
 * Enterprise module entry — wires tenant/user CRUD, mail, calendar,
 * workflow, SSO, audit and the enterprise-grade permissions set into
 * the shared registry.
 */
import { registerAuditHandlers, registerAuthHandlers } from './auth-audit'
import { registerCalendarHandlers, registerMailHandlers } from './communications'
import { registerEnterprisePermissionHandlers } from './permissions'
import { registerTenantHandlers, registerUserHandlers } from './users-tenants'
import { registerWorkflowHandlers } from './workflow'

export function registerEnterpriseHandlers(): void {
  registerUserHandlers()
  registerTenantHandlers()
  registerMailHandlers()
  registerCalendarHandlers()
  registerWorkflowHandlers()
  registerAuthHandlers()
  registerAuditHandlers()
  registerEnterprisePermissionHandlers()
}
