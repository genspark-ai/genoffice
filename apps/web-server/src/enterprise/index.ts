/**
 * Enterprise module entry — wires tenant/user CRUD, mail, calendar,
 * workflow, SSO, audit and the enterprise-grade permissions set into
 * the shared registry.
 */
import { registerAuditHandlers, registerAuthHandlers } from './auth-audit.js'
import { registerCalendarHandlers, registerMailHandlers } from './communications.js'
import { registerEnterprisePermissionHandlers } from './permissions.js'
import { registerTenantHandlers, registerUserHandlers } from './users-tenants.js'
import { registerWorkflowHandlers } from './workflow.js'

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
