import type { MailBody, MailEnvelope, MailFolder } from '@genoffice/mail-protocols'
import type { MailAccountSummary } from '../shared/ipc'

/**
 * Main-process demo mailbox behind the shell-hosted read handlers (merge
 * lane owns this file). FAKES ONLY — example.com / example.net addresses,
 * never real accounts; the real drivers and sync engine arrive with the
 * sync phase. Mirrors the renderer `devFakes.defaultSeed()` content so the
 * shell-hosted tab shows the same demo mailbox as plain vite dev.
 *
 * Bodies carry the fake-only `html` extension (never on the wire in
 * production): the quarantined reader renders it after sanitizing, with
 * remote images blocked. The canonical `MailBody` shape stays `{text,
 * htmlRef?}`; the extra field rides IPC via structured clone.
 */

export interface MailFakeBody extends MailBody {
  html?: string
}

export interface MailFakeData {
  accounts: MailAccountSummary[]
  folders: MailFolder[]
  envelopes: MailEnvelope[]
  bodies: Record<string, MailFakeBody>
  serverExtra: MailEnvelope[]
}

function envelope(
  uid: string,
  folderId: string,
  from: string[],
  to: string[],
  subject: string,
  date: string,
  snippet: string,
  over?: Partial<MailEnvelope>,
): MailEnvelope {
  return {
    ref: { uid, folderId, threadId: `thread-${uid}` },
    from,
    to,
    cc: [],
    subject,
    date,
    flags: { seen: false, flagged: false, answered: false, draft: false },
    labels: [],
    snippet,
    hasAttach: false,
    ...over,
  }
}

const READER = ['Reader <reader@example.net>']

/** Demo mailbox seed (fresh objects per call; handlers seed once). */
export function buildMailFakeData(): MailFakeData {
  const accounts: MailAccountSummary[] = [
    { accountId: 'demo', label: 'Demo Account', sessionOnly: true },
  ]
  const folders: MailFolder[] = [
    { id: 'inbox', name: 'Inbox', role: 'inbox', unread: 3 },
    { id: 'sent', name: 'Sent', role: 'sent', unread: 0 },
    { id: 'drafts', name: 'Drafts', role: 'drafts', unread: 0 },
    { id: 'archive', name: 'Archive', role: 'archive', unread: 0 },
  ]
  const envelopes: MailEnvelope[] = [
    envelope(
      'm1',
      'inbox',
      ['Acme Billing <billing@example.com>'],
      READER,
      'Your invoice is ready',
      '2026-10-10T09:00:00.000Z',
      'Invoice INV-1042 for October is attached for review.',
      { hasAttach: true },
    ),
    envelope(
      'm2',
      'inbox',
      ['Team Chat <chat@example.com>'],
      READER,
      'Weekly notes with a remote banner',
      '2026-10-11T12:30:00.000Z',
      'This week shipped the new reader with remote images blocked.',
      { flags: { seen: true, flagged: false, answered: false, draft: false } },
    ),
    envelope(
      'm3',
      'inbox',
      ['Support <support@example.com>'],
      READER,
      'Reply-To differs on this one',
      '2026-10-12T08:15:00.000Z',
      'A support note whose Reply-To points elsewhere.',
    ),
    envelope(
      'm4',
      'inbox',
      ['Acme <mail@xn--acme-pay-example.com>'],
      READER,
      'Lookalike sender sample',
      '2026-10-12T18:40:00.000Z',
      'Sent from a localised domain that mimics a known sender.',
    ),
    envelope(
      'm5',
      'inbox',
      ['Flags Demo <flags@example.com>'],
      READER,
      'Flagged and unread sample',
      '2026-10-13T07:05:00.000Z',
      'Starred so flag filters have something to match.',
      { flags: { seen: false, flagged: true, answered: false, draft: false } },
    ),
    envelope(
      'm6',
      'sent',
      READER,
      ['Acme Billing <billing@example.com>'],
      'Re: Your invoice is ready',
      '2026-10-13T09:20:00.000Z',
      'Thanks — paid this morning, receipt appreciated.',
      { flags: { seen: true, flagged: false, answered: false, draft: false } },
    ),
  ]
  const serverExtra: MailEnvelope[] = [
    envelope(
      'srv1',
      'inbox',
      ['Archive Bot <archive@example.com>'],
      READER,
      'Server-only invoice copy',
      '2026-09-01T10:00:00.000Z',
      'An older invoice copy that lives on the server only.',
      { flags: { seen: true, flagged: false, answered: false, draft: false } },
    ),
  ]
  const bodies: Record<string, MailFakeBody> = {
    'inbox/m1': {
      text: 'Hello,\n\nInvoice INV-1042 for October is attached for review.\n\nAcme Billing',
      html: '<p>Hello,</p><p>Invoice <b>INV-1042</b> for October is attached for review.</p><p>Thank you for your business.</p>',
    },
    'inbox/m2': {
      text: 'This week shipped the new reader with remote images blocked.',
      html: '<p>This week shipped the new reader.</p><img src="https://tracker.example.com/pixel.png" alt="Weekly banner"><p>Remote images stay blocked until you allow them.</p>',
    },
    'inbox/m3': {
      text: 'A support note whose Reply-To points elsewhere.',
      html: '<p>A support note whose Reply-To points elsewhere.</p>',
    },
    'inbox/m4': {
      text: 'Sent from a localised domain that mimics a known sender.',
      html: '<p>Sent from a localised domain that mimics a known sender.</p>',
    },
    'inbox/m5': {
      text: 'Starred so flag filters have something to match.',
    },
    'sent/m6': {
      text: 'Thanks — paid this morning, receipt appreciated.',
    },
    'inbox/srv1': {
      text: 'An older invoice copy that lives on the server only.',
    },
  }
  return { accounts, folders, envelopes, bodies, serverExtra }
}
