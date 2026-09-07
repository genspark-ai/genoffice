/**
 * Enterprise communication channels — mail and calendar. All state is in
 * `common/state.ts`; these handlers just translate wire payloads to/from
 * the shared maps.
 */
import { CALENDARS, MAILS, registerHandle } from '../common/index.js'

export function registerMailHandlers(): void {
  registerHandle('mail:send', async (_event: unknown, args: unknown) => {
    const { to, subject, body, cc, bcc, attachments } = (args || {}) as {
      to: Array<{ name: string; email: string }>
      subject: string
      body: string
      cc?: Array<{ name: string; email: string }>
      bcc?: Array<{ name: string; email: string }>
      attachments?: Array<{ name: string; size: number }>
    }
    const id = `mail-${Date.now()}`
    MAILS.set(id, {
      id,
      tenantId: 'default',
      from: { name: 'GenOffice', email: 'noreply@genoffice.ai' },
      to: to || [],
      subject,
      body,
      attachments: attachments || [],
      sentAt: Date.now(),
      status: 'pending',
    })
    setTimeout(() => {
      const mail = MAILS.get(id)
      if (mail) mail.status = 'sent'
    }, 1000)
    return { ok: true, id }
  })

  registerHandle('mail:list', (_event: unknown, args: unknown) => {
    const { folder, limit, offset } = (args || {}) as {
      folder?: 'inbox' | 'sent' | 'draft' | 'trash'
      limit?: number
      offset?: number
    }
    const maxResults = limit || 20
    const startOffset = offset || 0
    return [...MAILS.values()]
      .filter(m => folder ? m.status === folder : true)
      .slice(startOffset, startOffset + maxResults)
      .map(m => ({
        id: m.id,
        from: m.from,
        to: m.to,
        subject: m.subject,
        sentAt: m.sentAt,
        status: m.status,
      }))
  })

  registerHandle('mail:get', (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    return MAILS.get(id) || null
  })
}

export function registerCalendarHandlers(): void {
  registerHandle('calendar:create-event', (_event: unknown, args: unknown) => {
    const { title, description, startTime, endTime, attendees, location, reminders, recurrence } = (args || {}) as {
      title: string
      description?: string
      startTime: number
      endTime: number
      attendees?: Array<{ name: string; email: string }>
      location?: string
      reminders?: number[]
      recurrence?: string
    }
    const id = `event-${Date.now()}`
    CALENDARS.set(id, {
      id,
      tenantId: 'default',
      title,
      description: description || '',
      startTime,
      endTime,
      attendees: (attendees || []).map(a => ({ ...a, status: 'pending' as const })),
      location,
      reminders: reminders || [15, 60],
      recurrence,
      status: 'confirmed',
    })
    return { ok: true, id }
  })

  registerHandle('calendar:list-events', (_event: unknown, args: unknown) => {
    const { startDate, endDate, limit, offset } = (args || {}) as {
      startDate?: number
      endDate?: number
      limit?: number
      offset?: number
    }
    const maxResults = limit || 50
    const startOffset = offset || 0
    let events = [...CALENDARS.values()]
    if (startDate) events = events.filter(e => e.startTime >= startDate)
    if (endDate) events = events.filter(e => e.endTime <= endDate)
    return events
      .sort((a, b) => a.startTime - b.startTime)
      .slice(startOffset, startOffset + maxResults)
  })

  registerHandle('calendar:update-event', (_event: unknown, args: unknown) => {
    const { id, title, description, startTime, endTime, attendees, location } = (args || {}) as {
      id: string
      title?: string
      description?: string
      startTime?: number
      endTime?: number
      attendees?: Array<{ name: string; email: string; status: string }>
      location?: string
    }
    const event = CALENDARS.get(id)
    if (!event) return { ok: false, error: 'Event not found' }
    if (title) event.title = title
    if (description !== undefined) event.description = description
    if (startTime) event.startTime = startTime
    if (endTime) event.endTime = endTime
    if (attendees) event.attendees = attendees as typeof event.attendees
    if (location !== undefined) event.location = location
    return { ok: true }
  })

  registerHandle('calendar:delete-event', (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    if (!CALENDARS.has(id)) return { ok: false, error: 'Event not found' }
    CALENDARS.delete(id)
    return { ok: true }
  })
}
