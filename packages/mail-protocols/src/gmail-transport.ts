/**
 * Production GmailTransport over the googleapis Gmail v1 client.
 *
 * Auth is a bearer access token supplied per call by the main process;
 * no client secret, no refresh token, and no safeStorage access live in
 * this module. Token refresh and vault policy stay in
 * apps/mail/src/main (credential-vault.ts owns the only safeStorage
 * calls). OAuth scope stays gmail.modify (see GMAIL_MODIFY_SCOPE).
 */

import { google } from 'googleapis'
import type { GmailAuthProvider, GmailTransport } from './gmail'

/**
 * Build the production transport. Each API call fetches a fresh
 * access token from the main-process auth provider, so rotation in
 * the vault applies without driver restarts.
 */
export function createGmailTransport(auth: GmailAuthProvider): GmailTransport {
  const clients = async () => {
    const token = await auth.getAccessToken()
    const oauth = new google.auth.OAuth2()
    oauth.setCredentials({ access_token: token })
    return google.gmail({ version: 'v1', auth: oauth })
  }
  return {
    async listLabels() {
      const gmail = await clients()
      const res = await gmail.users.labels.list({ userId: 'me' })
      return (res.data.labels ?? []).map((label) => ({
        id: label.id ?? '',
        name: label.name ?? label.id ?? '',
        messagesUnread: label.messagesUnread ?? 0,
        type: label.type ?? undefined,
      }))
    },

    async listMessages(args) {
      const gmail = await clients()
      const res = await gmail.users.messages.list({
        userId: 'me',
        labelIds: args.labelIds,
        q: args.q,
        maxResults: args.maxResults,
        pageToken: args.pageToken,
      })
      return {
        messages: (res.data.messages ?? []).map((m) => ({
          id: m.id ?? '',
          threadId: m.threadId ?? '',
        })),
        nextPageToken: res.data.nextPageToken ?? undefined,
      }
    },

    async getMessage(id, format) {
      const gmail = await clients()
      const res = await gmail.users.messages.get({
        userId: 'me',
        id,
        format: format === 'full' ? 'full' : 'metadata',
        metadataHeaders: ['From', 'To', 'Cc', 'Subject', 'Date'],
      })
      const payload = res.data.payload
      return {
        id: res.data.id ?? id,
        threadId: res.data.threadId ?? undefined,
        labelIds: res.data.labelIds ?? [],
        snippet: res.data.snippet ?? '',
        historyId: res.data.historyId ?? undefined,
        sizeEstimate: res.data.sizeEstimate ?? undefined,
        payload: payload
          ? {
              mimeType: payload.mimeType ?? undefined,
              filename: payload.filename ?? undefined,
              headers: (payload.headers ?? []).map((h) => ({
                name: h.name ?? '',
                value: h.value ?? '',
              })),
              body: payload.body
                ? { data: payload.body.data ?? undefined, size: payload.body.size ?? undefined }
                : undefined,
              parts: undefined,
            }
          : undefined,
      }
    },

    async sendMessage(rawBase64Url) {
      const gmail = await clients()
      const res = await gmail.users.messages.send({
        userId: 'me',
        requestBody: { raw: rawBase64Url },
      })
      return { id: res.data.id ?? undefined }
    },

    async modifyMessage(id, addLabelIds, removeLabelIds) {
      const gmail = await clients()
      await gmail.users.messages.modify({
        userId: 'me',
        id,
        requestBody: { addLabelIds, removeLabelIds },
      })
    },

    async trashMessage(id) {
      const gmail = await clients()
      await gmail.users.messages.trash({ userId: 'me', id })
    },

    async listHistory(startHistoryId) {
      const gmail = await clients()
      const res = await gmail.users.history.list({ userId: 'me', startHistoryId })
      const records = res.data.history ?? []
      const slim = (
        m?: { id?: string | null; threadId?: string | null; labelIds?: string[] | null } | null,
      ) => ({
        message: m
          ? { id: m.id ?? '', threadId: m.threadId ?? '', labelIds: m.labelIds ?? [] }
          : null,
      })
      return {
        historyId: res.data.historyId ?? startHistoryId,
        added: records.flatMap((h) => h.messagesAdded ?? []).map((e) => slim(e.message)),
        deleted: records.flatMap((h) => h.messagesDeleted ?? []).map((e) => slim(e.message)),
        labelsChanged: records
          .flatMap((h) => [...(h.labelsAdded ?? []), ...(h.labelsRemoved ?? [])])
          .map((e) => slim(e.message)),
      }
    },

    async getCurrentHistoryId() {
      const gmail = await clients()
      const res = await gmail.users.getProfile({ userId: 'me' })
      return res.data.historyId ?? '0'
    },
  }
}
