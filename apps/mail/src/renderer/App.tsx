import { useCallback, useEffect, useMemo, useState } from 'react'
import type { MailEnvelope, MailFolder } from '@genoffice/mail-protocols'
import type { MailAccountSummary, MailApi } from '../shared/ipc'
import { useI18n } from './i18n/locale'
import { FolderTree, type FolderSelection } from './folders/FolderTree'
import { MessageList } from './list/MessageList'
import { SearchBox } from './list/SearchBox'
import { MessageReader, type ReaderSecurity } from './reader/MessageReader'
import { AiPanel, type AiCitationRef, type AiPanelSkillContext } from './ai/AiPanel'
import { createTriageSkill } from './ai/triage-skill'
import { createSummarizeSkill } from './ai/summarize-skill'
import { createComposeAssistSkill } from './ai/compose-assist-skill'
import { createMailSearchSkill } from './ai/search-skill'
import { ComposeHost } from './compose/ComposeHost'
import { OutboxBanner } from './components/OutboxBanner'
import { OutboxPanel } from './components/OutboxPanel'
import { OfflineIndicator } from './components/OfflineIndicator'
import { AccountSyncStatus } from './components/AccountSyncStatus'
import { applyLocalSearch, FakeMailApi, resolveMailApi } from './devFakes'
import './layout.css'

/**
 * Mailbox composition (Lane R owns). Folder tree, virtualized list,
 * quarantined reader, search box, AI panel rail, and the <ComposeHost/>
 * stub mount (real compose is Lane C's; this file only renders the
 * mount point with the locked props). All reads go through `MailApi`
 * (IPC fakes in P3; the real sync engine is a later phase). The outbox
 * banner, panel dialog, offline indicator, and per-account sync status
 * mount here over the same api (P4 shell mount, UI lane owns).
 */

export interface MailAppProps {
  api?: MailApi
}

const PAGE_LIMIT = 50

type ListMode = 'folder' | 'search'

/**
 * P5 AI integration glue (QA-owned at merge): the four K-lane mail skills
 * over the panel's live scope readers. Skills resolve the W-lane retrieval
 * bridge structurally at call time (`window.mailApi.aiRetrieve`); the
 * compose draft flows through the panel context when App wires `getDraft`.
 */
export function createMailAiSkills(
  ctx: AiPanelSkillContext,
): ReturnType<typeof createTriageSkill>[] {
  return [
    createTriageSkill({ getScope: () => ({ accountIds: ctx.getScope().accountIds }) }),
    createSummarizeSkill({ getScope: () => ({ accountIds: ctx.getScope().accountIds }) }),
    createComposeAssistSkill({
      getScope: () => ({ accountIds: ctx.getScope().accountIds }),
      getDraft: () => ctx.getDraft(),
    }),
    createMailSearchSkill({ getScope: () => ({ accountIds: ctx.getScope().accountIds }) }),
  ]
}

export default function App({ api }: MailAppProps) {
  const { t } = useI18n()
  const resolvedApi = useMemo(() => api ?? resolveMailApi(), [api])

  const [accounts, setAccounts] = useState<MailAccountSummary[]>([])
  const [folderLists, setFolderLists] = useState<Record<string, MailFolder[]>>({})
  const [selection, setSelection] = useState<FolderSelection | null>(null)
  const [messages, setMessages] = useState<MailEnvelope[]>([])
  const [cursor, setCursor] = useState<string | undefined>(undefined)
  const [hasMore, setHasMore] = useState(false)
  const [loadingMessages, setLoadingMessages] = useState(false)
  const [selectedUid, setSelectedUid] = useState<string | null>(null)
  const [query, setQuery] = useState('')
  const [searching, setSearching] = useState(false)
  const [mode, setMode] = useState<ListMode>('folder')
  const [serverExpanded, setServerExpanded] = useState(false)
  const [notice, setNotice] = useState<string | null>(null)
  const [composeDirty, setComposeDirty] = useState(false)
  const [outboxOpen, setOutboxOpen] = useState(false)
  const [assistText, setAssistText] = useState<string | null>(null)

  const refreshAccounts = useCallback(async () => {
    const next = await resolvedApi.listAccounts()
    setAccounts(next)
    const folders: Record<string, MailFolder[]> = {}
    for (const account of next) {
      folders[account.accountId] = await resolvedApi.listFolders(account.accountId)
    }
    setFolderLists(folders)
    return next
  }, [resolvedApi])

  useEffect(() => {
    let cancelled = false
    void (async () => {
      try {
        const next = await refreshAccounts()
        if (cancelled || next.length === 0) return
        const first = next[0]
        const firstFolders = await resolvedApi.listFolders(first.accountId)
        const inbox = firstFolders.find((folder) => folder.role === 'inbox') ?? firstFolders[0]
        if (cancelled || !inbox) return
        setSelection({ accountId: first.accountId, folderId: inbox.id })
      } catch {
        if (!cancelled) setNotice(t('mail.errors.ipcRejected'))
      }
    })()
    return () => {
      cancelled = true
    }
    // t is not referentially stable; translate at render time only.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refreshAccounts, resolvedApi])

  const loadFolderPage = useCallback(
    async (target: FolderSelection, nextCursor?: string, append = false) => {
      setLoadingMessages(true)
      try {
        const page = await resolvedApi.listMessages(
          target.accountId,
          target.folderId,
          PAGE_LIMIT,
          nextCursor,
        )
        setMessages((prev) => (append ? [...prev, ...page.envelopes] : page.envelopes))
        setCursor(page.cursor)
        setHasMore(page.cursor !== undefined)
      } catch {
        setNotice(t('mail.errors.ipcRejected'))
      } finally {
        setLoadingMessages(false)
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [resolvedApi],
  )

  useEffect(() => {
    if (!selection) return
    setSelectedUid(null)
    setMode('folder')
    setServerExpanded(false)
    setQuery('')
    void loadFolderPage(selection)
  }, [selection, loadFolderPage])

  useEffect(() => {
    resolvedApi.setDirty(composeDirty)
  }, [resolvedApi, composeDirty])

  useEffect(() => {
    const off = resolvedApi.onCloseSaveRequest(() => {
      resolvedApi.sendCloseSaveResult(true)
    })
    const offSync = resolvedApi.onSyncEvent((event) => {
      if (event.kind === 'resync') setNotice(t('mail.errors.resyncNotice'))
      else if (event.kind === 'conflict') setNotice(t('mail.errors.conflictNotice'))
    })
    return () => {
      off()
      offSync()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [resolvedApi])

  const selectFolder = useCallback((target: FolderSelection) => {
    setSelection(target)
  }, [])

  const openMessage = useCallback(
    (envelope: MailEnvelope) => {
      setSelectedUid(envelope.ref.uid)
      void resolvedApi.setFlags(envelope.ref, { seen: true }).catch(() => undefined)
      setMessages((prev) =>
        prev.map((item) =>
          item.ref.uid === envelope.ref.uid && item.ref.folderId === envelope.ref.folderId
            ? { ...item, flags: { ...item.flags, seen: true } }
            : item,
        ),
      )
    },
    [resolvedApi],
  )

  const submitSearch = useCallback(async () => {
    if (!selection) return
    const needle = query.trim()
    if (!needle) {
      setMode('folder')
      setServerExpanded(false)
      await loadFolderPage(selection)
      return
    }
    setSearching(true)
    try {
      const hits = await resolvedApi.searchLocal([selection.accountId], { text: needle })
      setMessages(applyLocalSearch(hits, needle))
      setMode('search')
      setServerExpanded(false)
      setHasMore(true)
      setSelectedUid(null)
    } catch {
      setNotice(t('mail.errors.ipcRejected'))
    } finally {
      setSearching(false)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selection, query, resolvedApi, loadFolderPage])

  const clearSearch = useCallback(async () => {
    setQuery('')
    setMode('folder')
    setServerExpanded(false)
    if (selection) await loadFolderPage(selection)
  }, [selection, loadFolderPage])

  const loadMore = useCallback(async () => {
    if (!selection) return
    if (mode === 'folder') {
      if (cursor === undefined) return
      await loadFolderPage(selection, cursor, true)
      return
    }
    const needle = query.trim()
    if (!needle || serverExpanded) return
    try {
      const remote = await resolvedApi.searchServer(selection.accountId, selection.folderId, {
        text: needle,
      })
      const known = new Set(messages.map((item) => `${item.ref.folderId}/${item.ref.uid}`))
      setMessages((prev) => [
        ...prev,
        ...applyLocalSearch(remote, needle).filter(
          (item) => !known.has(`${item.ref.folderId}/${item.ref.uid}`),
        ),
      ])
      setServerExpanded(true)
      setHasMore(false)
    } catch {
      setNotice(t('mail.errors.ipcRejected'))
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selection, mode, cursor, query, serverExpanded, messages, resolvedApi, loadFolderPage])

  const addAccount = useCallback(async () => {
    try {
      const created = await resolvedApi.addAccount(
        `${t('mail.accounts.title')} ${accounts.length + 1}`,
      )
      setAccounts((prev) => [...prev, created])
      setFolderLists((prev) => ({ ...prev, [created.accountId]: [] }))
    } catch {
      setNotice(t('mail.errors.ipcRejected'))
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [resolvedApi, accounts.length])

  const removeAccount = useCallback(
    async (accountId: string) => {
      try {
        await resolvedApi.removeAccount(accountId)
        const next = await refreshAccounts()
        if (selection?.accountId === accountId) {
          const first = next[0]
          setSelection(first ? { accountId: first.accountId, folderId: 'inbox' } : null)
          if (!first) setMessages([])
        }
      } catch {
        setNotice(t('mail.errors.ipcRejected'))
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [resolvedApi, refreshAccounts, selection],
  )

  const selectedEnvelope = messages.find((item) => item.ref.uid === selectedUid) ?? null
  const security: ReaderSecurity =
    resolvedApi instanceof FakeMailApi && selectedEnvelope
      ? resolvedApi.fakeSecurity(selectedEnvelope.ref)
      : {}
  const showNoResults = mode === 'search' && !searching && messages.length === 0

  // AI citation click opens the message in the reader (same path as the
  // list selection); a folder-scoped match wins, a uid match suffices.
  const openCitation = useCallback(
    (citation: AiCitationRef) => {
      const found =
        messages.find(
          (item) => item.ref.uid === citation.uid && item.ref.folderId === citation.folderId,
        ) ?? messages.find((item) => item.ref.uid === citation.uid)
      if (found) openMessage(found)
    },
    [messages, openMessage],
  )
  const selectedAccount = accounts.find((account) => account.accountId === selection?.accountId)

  return (
    <div className="mail-app">
      <div className="mail-sidebar">
        <FolderTree
          accounts={accounts}
          foldersByAccount={folderLists}
          selected={selection}
          onSelect={selectFolder}
          onAddAccount={() => void addAccount()}
          onRemoveAccount={(accountId) => void removeAccount(accountId)}
        />
        <AiPanel
          accountIds={selection ? [selection.accountId] : []}
          accounts={accounts}
          onOpenCitation={openCitation}
          onInsertAssist={(text) => setAssistText(text)}
          createSkills={createMailAiSkills}
        />
        <OfflineIndicator api={resolvedApi} accountId={selection?.accountId} />
        {selection && selectedAccount && (
          <AccountSyncStatus
            api={resolvedApi}
            accountId={selection.accountId}
            label={selectedAccount.label}
          />
        )}
      </div>
      <div className="mail-listpane">
        <SearchBox
          value={query}
          searching={searching}
          onChange={setQuery}
          onSubmit={() => void submitSearch()}
          onClear={() => void clearSearch()}
        />
        <OutboxBanner
          api={resolvedApi}
          accountId={selection?.accountId ?? ''}
          notice={notice}
          onOpenPanel={() => setOutboxOpen(true)}
        />
        {outboxOpen && selection && (
          <div
            className="mail-outbox-dialog"
            role="dialog"
            aria-modal="true"
            aria-label={t('mail.outbox.title')}
          >
            <button
              type="button"
              className="mail-outbox-close"
              onClick={() => setOutboxOpen(false)}
            >
              {t('mail.outbox.close')}
            </button>
            <OutboxPanel
              api={resolvedApi}
              accountId={selection.accountId}
              folderId={selection.folderId}
            />
          </div>
        )}
        {showNoResults ? (
          <div className="mail-list-status">{t('mail.search.noResults')}</div>
        ) : (
          <MessageList
            messages={messages}
            selectedUid={selectedUid}
            loading={loadingMessages}
            hasMore={hasMore}
            onSelect={openMessage}
            onLoadMore={() => void loadMore()}
          />
        )}
      </div>
      <div className="mail-readerpane">
        <MessageReader api={resolvedApi} envelope={selectedEnvelope} security={security} />
        <div className="mail-compose-mount">
          <ComposeHost
            accountId={selection?.accountId ?? ''}
            threadKey={selectedEnvelope?.ref.threadId}
            onDirtyChange={setComposeDirty}
            assistText={assistText}
            onAssistConsumed={() => setAssistText(null)}
            onSend={(req) =>
              void resolvedApi
                .composeSend(req)
                .then(() => setNotice(t('mail.compose.outboxQueued')))
                .catch(() => setNotice(t('mail.errors.ipcRejected')))
            }
          />
        </div>
      </div>
    </div>
  )
}
