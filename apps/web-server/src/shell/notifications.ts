/**
 * Notifications channels — send/list/mark-read/clear. In-memory per-user
 * queue, no real push delivery in the web build.
 */
import { NOTIFICATIONS, registerHandle } from '../common/index.js'

export function registerNotificationHandlers(): void {
  registerHandle('notifications:send', (_event: unknown, args: unknown) => {
    const { userId, type, title, message } = args as {
      userId: string
      type: 'info' | 'success' | 'warning' | 'error'
      title: string
      message: string
    }

    if (!NOTIFICATIONS.has(userId)) {
      NOTIFICATIONS.set(userId, [])
    }

    const id = `notif-${Date.now()}`
    NOTIFICATIONS.get(userId)!.push({
      id,
      type,
      title,
      message,
      timestamp: Date.now(),
      read: false,
    })

    return { ok: true, id, unread: NOTIFICATIONS.get(userId)!.filter(n => !n.read).length }
  })

  registerHandle('notifications:list', (_event: unknown, args: unknown) => {
    const { userId, unreadOnly } = args as { userId: string; unreadOnly?: boolean }

    const notifications = NOTIFICATIONS.get(userId) || []
    if (unreadOnly) {
      return notifications.filter(n => !n.read)
    }
    return notifications
  })

  registerHandle('notifications:mark-read', (_event: unknown, args: unknown) => {
    const { userId, notificationId } = args as { userId: string; notificationId: string }

    const notifications = NOTIFICATIONS.get(userId)
    if (!notifications) return { ok: false, error: 'User not found' }

    const notification = notifications.find(n => n.id === notificationId)
    if (!notification) return { ok: false, error: 'Notification not found' }

    notification.read = true
    return { ok: true }
  })

  registerHandle('notifications:clear', (_event: unknown, args: unknown) => {
    const { userId } = args as { userId: string }
    const count = NOTIFICATIONS.get(userId)?.length || 0
    NOTIFICATIONS.delete(userId)
    return { ok: true, cleared: count }
  })
}
