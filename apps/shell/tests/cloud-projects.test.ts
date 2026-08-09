import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest'
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { gskListPastProjects, type GskPastProjectsPage } from '@genoffice/ai-search'
import {
  cloudErrorReason,
  cloudStoreOwner,
  clearCloudProjectsStore,
  readCloudProjectsStore,
  syncCloudProjects,
} from '../src/main/cloud-projects'

vi.mock('@genoffice/ai-search', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@genoffice/ai-search')>()
  return { ...actual, gskListPastProjects: vi.fn() }
})

const listMock = vi.mocked(gskListPastProjects)

const PROJECTS = [
  {
    projectId: 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee',
    title: 'Deck A',
    kind: 'slides',
    ctimeMs: 1_700_000_000_000,
    projectUrl: '/agents?id=aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee',
  },
]

describe('cloud projects store account binding', () => {
  let dir: string
  let storePath: string
  const envBefore = { key: process.env.GSK_API_KEY, disable: process.env.AI_SEARCH_DISABLE_GSK }

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'cloud-store-'))
    storePath = join(dir, 'cloud-projects.json')
    delete process.env.AI_SEARCH_DISABLE_GSK
    process.env.GSK_API_KEY = 'test-key-account-a'
  })

  afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
    if (envBefore.key === undefined) delete process.env.GSK_API_KEY
    else process.env.GSK_API_KEY = envBefore.key
    if (envBefore.disable === undefined) delete process.env.AI_SEARCH_DISABLE_GSK
    else process.env.AI_SEARCH_DISABLE_GSK = envBefore.disable
  })

  const writeStore = (owner: string) => {
    writeFileSync(
      storePath,
      JSON.stringify({ available: true, projects: PROJECTS, syncedAt: 123, owner }),
    )
  }

  it('derives the owner tag from the key without containing it', () => {
    const tag = cloudStoreOwner()
    expect(tag).toMatch(/^[0-9a-f]{16}$/)
    expect(tag).not.toContain('test-key')
    process.env.GSK_API_KEY = 'test-key-account-b'
    expect(cloudStoreOwner()).not.toBe(tag)
  })

  it('serves the store back to the same account', () => {
    writeStore(cloudStoreOwner())
    const snap = readCloudProjectsStore(storePath)
    expect(snap?.projects.map((p) => p.title)).toEqual(['Deck A'])
    expect(snap?.syncedAt).toBe(123)
  })

  it("rejects and deletes another account's store", () => {
    writeStore(cloudStoreOwner())
    process.env.GSK_API_KEY = 'test-key-account-b'
    expect(readCloudProjectsStore(storePath)).toBeNull()
    expect(existsSync(storePath)).toBe(false)
  })

  it('rejects a legacy store without an owner tag', () => {
    writeFileSync(storePath, JSON.stringify({ available: true, projects: PROJECTS, syncedAt: 1 }))
    expect(readCloudProjectsStore(storePath)).toBeNull()
  })

  it('clearCloudProjectsStore removes the file and tolerates a missing one', () => {
    writeStore(cloudStoreOwner())
    clearCloudProjectsStore(storePath)
    expect(existsSync(storePath)).toBe(false)
    clearCloudProjectsStore(storePath)
  })
})

describe('cloud projects sync account isolation', () => {
  let dir: string
  let storePath: string
  const envBefore = { key: process.env.GSK_API_KEY, disable: process.env.AI_SEARCH_DISABLE_GSK }

  const pageFor = (title: string): GskPastProjectsPage => ({
    projects: [
      {
        projectId: `id-${title}`,
        type: 'slides_agent_git',
        title,
        ctime: '2026-08-01T00:00:00',
        projectUrl: `/agents?id=id-${title}`,
      },
    ],
    total: 1,
    hasMore: false,
  })

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'cloud-sync-'))
    storePath = join(dir, 'cloud-projects.json')
    delete process.env.AI_SEARCH_DISABLE_GSK
    process.env.GSK_API_KEY = 'test-key-account-a'
    listMock.mockReset()
  })

  afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
    if (envBefore.key === undefined) delete process.env.GSK_API_KEY
    else process.env.GSK_API_KEY = envBefore.key
    if (envBefore.disable === undefined) delete process.env.AI_SEARCH_DISABLE_GSK
    else process.env.AI_SEARCH_DISABLE_GSK = envBefore.disable
  })

  it('writes the store bound to the account that synced', async () => {
    listMock.mockResolvedValue(pageFor('Deck A'))
    const snap = await syncCloudProjects(storePath)
    expect(snap.projects.map((p) => p.title)).toEqual(['Deck A'])
    expect(readCloudProjectsStore(storePath)?.projects[0]?.title).toBe('Deck A')
  })

  it('aborts without touching the store when the account switches mid-sync', async () => {
    listMock.mockImplementation(async () => {
      // the page comes back after the user has switched accounts
      process.env.GSK_API_KEY = 'test-key-account-b'
      return pageFor('Deck B')
    })
    await expect(syncCloudProjects(storePath)).rejects.toThrow(/account changed/)
    expect(existsSync(storePath)).toBe(false)
  })

  it('does not share an in-flight sync across accounts', async () => {
    const gates: Array<(page: GskPastProjectsPage) => void> = []
    listMock.mockImplementation(
      () => new Promise<GskPastProjectsPage>((resolve) => gates.push(resolve)),
    )

    const first = syncCloudProjects(storePath)
    expect(syncCloudProjects(storePath)).toBe(first) // same account shares the run

    process.env.GSK_API_KEY = 'test-key-account-b'
    const second = syncCloudProjects(storePath)
    expect(second).not.toBe(first)
    expect(listMock).toHaveBeenCalledTimes(2)

    gates[0]!(pageFor('Deck A'))
    gates[1]!(pageFor('Deck B'))

    // the run started for account A sees the key changed and aborts
    await expect(first).rejects.toThrow(/account changed/)
    const snapB = await second
    expect(snapB.projects.map((p) => p.title)).toEqual(['Deck B'])
    expect(readCloudProjectsStore(storePath)?.projects[0]?.title).toBe('Deck B')
  })
})

// A failed sync used to reject, and the home screen turned every cause into
// "load failed, try again later" with a Retry button. For an out-of-credits
// account that is actively wrong: retrying can never succeed.
describe('cloudErrorReason', () => {
  it('recognises the out-of-credits message the gsk CLI actually returns', () => {
    // verbatim from `gsk projects` on a free account with a zero balance
    expect(cloudErrorReason(new Error('gsk returned an error: Insufficient credits'))).toBe(
      'credits',
    )
  })

  it('recognises an expired sign-in', () => {
    expect(cloudErrorReason(new Error('Request failed: 401 Unauthorized'))).toBe('signedOut')
    expect(cloudErrorReason(new Error('gsk: not logged in'))).toBe('signedOut')
  })

  it('recognises the shapes a blocked or proxy-less network produces', () => {
    for (const message of [
      'connect ECONNREFUSED 127.0.0.1:7890',
      'fetch failed: ETIMEDOUT',
      'getaddrinfo ENOTFOUND www.genspark.ai',
    ]) {
      expect(cloudErrorReason(new Error(message)), message).toBe('network')
    }
  })

  it('falls back to unknown rather than guessing', () => {
    expect(cloudErrorReason(new Error('something else entirely'))).toBe('unknown')
    expect(cloudErrorReason('not even an Error')).toBe('unknown')
  })
})

describe('syncCloudProjects failure handling', () => {
  let failDir: string

  beforeEach(() => {
    failDir = mkdtempSync(join(tmpdir(), 'cloud-fail-'))
  })

  afterEach(() => {
    rmSync(failDir, { recursive: true, force: true })
  })

  it('resolves with the reason and the cached list instead of rejecting', async () => {
    const storePath = join(failDir, 'cloud.json')
    listMock.mockRejectedValueOnce(new Error('gsk returned an error: Insufficient credits'))

    const snapshot = await syncCloudProjects(storePath)

    expect(snapshot.error).toBe('credits')
    // the home screen keeps rendering whatever it already had
    expect(Array.isArray(snapshot.projects)).toBe(true)
  })
})
