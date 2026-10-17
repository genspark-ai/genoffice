import { contextBridge, ipcRenderer } from 'electron'
import type { Lang } from '@genoffice/i18n'
import type { AiStreamChunk } from '@genoffice/ai-provider'
import { AI_CHANNELS, MAIL_CHANNELS } from '../shared/ipc'
import type { AiRetrieveRequest, MailApi, UiTheme } from '../shared/ipc'
import type { OfflineStatus } from '../shared/ipc'
import type { MailFlags, MailMessageRef, MailQuery, SendRequest } from '@genoffice/mail-protocols'
import { isMailRef } from '../shared/mail-ref'

// Shell-channel bridge (Lane S owns). Generic passthrough: every method
// forwards to its canonical MAIL_CHANNELS entry with no logic beyond a
// hand-rolled shape guard. This file stays zod-free (sheets precedent);
// main-process handlers re-validate with the zod schemas in
// `../shared/ipc-schemas`. Lane C's compose bridge lives separately in
// `compose-bridge.ts` (Lane C owns).

function assertMailRef(ref: string): void {
  if (!isMailRef(ref)) throw new Error('mail: bad mailRef')
}

function assertRef(ref: MailMessageRef): void {
  if (!ref || typeof ref.uid !== 'string' || typeof ref.folderId !== 'string') {
    throw new Error('mail: bad message ref')
  }
}

function assertQuery(query: MailQuery): void {
  if (!query || typeof query !== 'object') throw new Error('mail: bad query')
}

function assertSend(req: SendRequest): void {
  if (
    !req ||
    !Array.isArray(req.to) ||
    typeof req.subject !== 'string' ||
    typeof req.text !== 'string'
  ) {
    throw new Error('mail: bad send request')
  }
}

function assertAccountId(accountId: string): void {
  if (typeof accountId !== 'string' || accountId.length < 1 || accountId.length > 256) {
    throw new Error('mail: bad account id')
  }
}

function assertOutboxId(outboxId: string): void {
  if (typeof outboxId !== 'string' || outboxId.length < 1 || outboxId.length > 256) {
    throw new Error('mail: bad outbox id')
  }
}

function assertOfflineStatus(status: OfflineStatus): void {
  if (!status || typeof status.online !== 'boolean') throw new Error('mail: bad offline status')
}

// P5 AI retrieval (W owns): hand-rolled guard mirrors mailAiRetrieveSchema;
// main re-validates with zod end-to-end. H4 adds ref-shape checks incl the
// optional exact-match accountId hint (stale hints fail loud in main).
function assertAiRef(ref: MailMessageRef): void {
  if (
    !ref ||
    typeof ref.uid !== 'string' ||
    ref.uid.length < 1 ||
    typeof ref.folderId !== 'string' ||
    ref.folderId.length < 1
  ) {
    throw new Error('mail: bad ai-retrieve ref')
  }
  if (
    ref.accountId !== undefined &&
    (typeof ref.accountId !== 'string' || ref.accountId.length < 1 || ref.accountId.length > 256)
  ) {
    throw new Error('mail: bad ai-retrieve ref')
  }
}

function assertAiRetrieve(req: AiRetrieveRequest): void {
  if (!req || typeof req !== 'object') throw new Error('mail: bad ai-retrieve request')
  const scope = req.scope
  if (
    !scope ||
    !Array.isArray(scope.accountIds) ||
    scope.accountIds.length < 1 ||
    scope.accountIds.length > 8 ||
    scope.accountIds.some((id) => typeof id !== 'string' || id.length < 1 || id.length > 256)
  ) {
    throw new Error('mail: bad ai-retrieve scope')
  }
  if (req.op !== 'triage-headers' && req.op !== 'fetch-bodies' && req.op !== 'cite-resolve') {
    throw new Error('mail: bad ai-retrieve op')
  }
  if (req.op === 'triage-headers' && typeof req.query !== 'string') {
    throw new Error('mail: triage-headers requires a query')
  }
  if ((req.op === 'fetch-bodies' || req.op === 'cite-resolve') && !Array.isArray(req.refs)) {
    throw new Error('mail: fetch-bodies and cite-resolve require refs')
  }
  if (req.refs !== undefined) {
    if (!Array.isArray(req.refs) || req.refs.length > 50) {
      throw new Error('mail: bad ai-retrieve refs')
    }
    for (const ref of req.refs) assertAiRef(ref)
  }
  if (
    req.limit !== undefined &&
    (!Number.isInteger(req.limit) || req.limit < 1 || req.limit > 200)
  ) {
    throw new Error('mail: bad ai-retrieve limit')
  }
}

const api: MailApi = {
  openMessage: (mailRef) => {
    assertMailRef(mailRef)
    return ipcRenderer.invoke(MAIL_CHANNELS.openMessage, { mailRef })
  },
  listAccounts: () => ipcRenderer.invoke(MAIL_CHANNELS.listAccounts),
  addAccount: (label) => ipcRenderer.invoke(MAIL_CHANNELS.addAccount, { label }),
  removeAccount: (accountId) => ipcRenderer.invoke(MAIL_CHANNELS.removeAccount, { accountId }),
  listFolders: (accountId) => ipcRenderer.invoke(MAIL_CHANNELS.listFolders, { accountId }),
  listMessages: (accountId, folderId, limit, cursor) =>
    ipcRenderer.invoke(MAIL_CHANNELS.listMessages, { accountId, folderId, limit, cursor }),
  fetchBody: (ref) => {
    assertRef(ref)
    return ipcRenderer.invoke(MAIL_CHANNELS.fetchBody, { ref })
  },
  searchLocal: (accountIds, query) => {
    assertQuery(query)
    return ipcRenderer.invoke(MAIL_CHANNELS.searchLocal, { accountIds, query })
  },
  searchServer: (accountId, folderId, query) => {
    assertQuery(query)
    return ipcRenderer.invoke(MAIL_CHANNELS.searchServer, { accountId, folderId, query })
  },
  composeSend: (req) => {
    assertSend(req)
    return ipcRenderer.invoke(MAIL_CHANNELS.composeSend, req)
  },
  saveDraft: (accountId, threadKey, text) =>
    ipcRenderer.invoke(MAIL_CHANNELS.saveDraft, { accountId, threadKey, text }),
  outboxStatus: (accountId) => ipcRenderer.invoke(MAIL_CHANNELS.outboxStatus, { accountId }),
  move: (ref, toFolderId, expectedModseq) => {
    assertRef(ref)
    return ipcRenderer.invoke(MAIL_CHANNELS.move, { ref, toFolderId, expectedModseq })
  },
  setFlags: (ref, flags, expectedModseq) => {
    assertRef(ref)
    return ipcRenderer.invoke(MAIL_CHANNELS.flags, { ref, flags, expectedModseq })
  },
  deleteMessage: (ref, expectedModseq) => {
    assertRef(ref)
    return ipcRenderer.invoke(MAIL_CHANNELS.delete, { ref, expectedModseq })
  },
  onSyncEvent: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, event: Parameters<typeof handler>[0]) =>
      handler(event)
    ipcRenderer.on(MAIL_CHANNELS.syncEvents, listener)
    return () => ipcRenderer.removeListener(MAIL_CHANNELS.syncEvents, listener)
  },
  // P4 sync/outbox additions (Y owns): hand-rolled guards mirror the zod
  // schemas in `../shared/ipc-schemas`; main re-validates end-to-end.
  syncNow: (accountId, folderId) => {
    assertAccountId(accountId)
    return ipcRenderer.invoke(MAIL_CHANNELS.syncNow, { accountId, folderId })
  },
  outboxList: (accountId) => {
    assertAccountId(accountId)
    return ipcRenderer.invoke(MAIL_CHANNELS.outboxList, { accountId })
  },
  outboxRetry: (accountId, outboxId) => {
    assertAccountId(accountId)
    assertOutboxId(outboxId)
    return ipcRenderer.invoke(MAIL_CHANNELS.outboxRetry, { accountId, outboxId })
  },
  outboxCancel: (accountId, outboxId) => {
    assertAccountId(accountId)
    assertOutboxId(outboxId)
    return ipcRenderer.invoke(MAIL_CHANNELS.outboxCancel, { accountId, outboxId })
  },
  onOfflineStatus: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, status: OfflineStatus) => {
      assertOfflineStatus(status)
      handler(status)
    }
    ipcRenderer.on(MAIL_CHANNELS.offlineStatus, listener)
    return () => ipcRenderer.removeListener(MAIL_CHANNELS.offlineStatus, listener)
  },
  attachmentFetch: (attachmentPath) =>
    ipcRenderer.invoke(MAIL_CHANNELS.attachmentFetch, { attachmentPath }),
  onAttachmentProgress: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, completed: number, total: number) =>
      handler(completed, total)
    ipcRenderer.on(MAIL_CHANNELS.attachmentProgress, listener)
    return () => ipcRenderer.removeListener(MAIL_CHANNELS.attachmentProgress, listener)
  },
  setDirty: (dirty) => ipcRenderer.send(MAIL_CHANNELS.dirtyChanged, dirty),
  onCloseSaveRequest: (handler) => {
    const listener = () => handler()
    ipcRenderer.on(MAIL_CHANNELS.closeSaveRequest, listener)
    return () => ipcRenderer.removeListener(MAIL_CHANNELS.closeSaveRequest, listener)
  },
  sendCloseSaveResult: (ok) => ipcRenderer.send(MAIL_CHANNELS.closeSaveResult, ok),
  getLanguage: () => ipcRenderer.invoke(MAIL_CHANNELS.getLanguage),
  onLanguageChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, lang: Lang) => handler(lang)
    ipcRenderer.on(MAIL_CHANNELS.languageChanged, listener)
    return () => ipcRenderer.removeListener(MAIL_CHANNELS.languageChanged, listener)
  },
  getTheme: () => ipcRenderer.invoke(MAIL_CHANNELS.getTheme),
  onThemeChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, theme: UiTheme) => handler(theme)
    ipcRenderer.on(MAIL_CHANNELS.themeChanged, listener)
    return () => ipcRenderer.removeListener(MAIL_CHANNELS.themeChanged, listener)
  },
  getAiPanelPrefs: () => ipcRenderer.invoke('app:get-ai-panel-prefs'),
  onAiPanelPrefsChanged: (handler) => {
    const listener = (_event: Electron.IpcRendererEvent, prefs: Parameters<typeof handler>[0]) =>
      handler(prefs)
    ipcRenderer.on('app:ai-panel-prefs-changed', listener)
    return () => ipcRenderer.removeListener('app:ai-panel-prefs-changed', listener)
  },
  onChromePressed: (handler) => {
    const listener = () => handler()
    ipcRenderer.on('app:chrome-pressed', listener)
    return () => ipcRenderer.removeListener('app:chrome-pressed', listener)
  },
  getAiSettings: () => ipcRenderer.invoke(AI_CHANNELS.getSettings),
  aiGskStatus: () => ipcRenderer.invoke(AI_CHANNELS.gskStatus),
  aiStream: (request) => ipcRenderer.invoke(AI_CHANNELS.stream, request),
  aiStreamCancel: (requestId) => ipcRenderer.invoke(AI_CHANNELS.streamCancel, requestId),
  // P5 scoped retrieval (W owns): guard here, zod re-validation in main.
  aiRetrieve: (request) => {
    assertAiRetrieve(request)
    return ipcRenderer.invoke(MAIL_CHANNELS.aiRetrieve, request)
  },
  onAiStream: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, chunk: AiStreamChunk) => handler(chunk)
    ipcRenderer.on(AI_CHANNELS.streamChunk, listener)
    return () => ipcRenderer.removeListener(AI_CHANNELS.streamChunk, listener)
  },
}

contextBridge.exposeInMainWorld('mailApi', api)

export type { MailApi }
export type { MailFlags }
