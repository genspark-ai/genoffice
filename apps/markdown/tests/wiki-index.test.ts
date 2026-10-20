/**
 * The vault link index behind the backlinks panel and the graph view: what
 * counts as a link when a whole vault is scanned, how links invert into
 * backlinks, and that the mtime memo does not serve stale edges.
 */
import { mkdirSync, mkdtempSync, renameSync, writeFileSync, utimesSync } from 'node:fs'
import { readdir } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'

import {
  MAX_GRAPH_NODES,
  extractWikiTargets,
  nodeVaultFs,
  wikiBacklinks,
  wikiGraph,
  type VaultFs,
} from '../src/main/wiki-index'

const dirs: string[] = []
afterEach(() => {
  dirs.length = 0
})

function vault(files: Record<string, string>): string {
  const root = mkdtempSync(join(tmpdir(), 'wiki-index-'))
  dirs.push(root)
  for (const [name, text] of Object.entries(files)) {
    const path = join(root, name)
    const sep = name.lastIndexOf('/')
    if (sep > 0) {
      mkdirSync(join(root, name.slice(0, sep)), { recursive: true })
    }
    writeFileSync(path, text)
  }
  return root
}

/**
 * nodeVaultFs() caches each directory listing for a couple of seconds (the
 * resolver walks the same tree on every keystroke). A test that writes a file
 * and asks again straight away would be reading the cache, not the index — so
 * the tests that exercise a file-set change use this uncached walk instead of
 * waiting out a TTL.
 */
function uncachedVaultFs(): VaultFs {
  const base = nodeVaultFs()
  return {
    ...base,
    list: async (dir) => {
      const files: string[] = []
      const dirs: string[] = []
      for (const entry of await readdir(dir, { withFileTypes: true })) {
        if (entry.isDirectory()) dirs.push(entry.name)
        else if (entry.isFile()) files.push(entry.name)
      }
      return { files, dirs }
    },
  }
}

describe('extractWikiTargets', () => {
  it('reads links and embeds, dropping aliases, deduplicated in order', () => {
    expect(
      extractWikiTargets('![[Cover.png]] then [[Alpha]] and [[Alpha|again]] and [[Beta#Sec]]'),
    ).toEqual(['Cover.png', 'Alpha', 'Beta#Sec'])
  })

  it('ignores the syntax inside fenced blocks and inline code', () => {
    expect(extractWikiTargets('`[[NotALink]]`\n\n```\n[[AlsoNot]]\n```\n\n[[Real]]')).toEqual([
      'Real',
    ])
  })

  it('ignores a doubly-escaped bracket pair and an empty one', () => {
    // `\[\[` escapes both brackets, so what follows is plain text — the same
    // rule the editor's own tokenizer applies
    expect(extractWikiTargets('Literal \\[\\[NotALink]] end and [[]] done')).toEqual([])
  })
})

describe('wikiBacklinks', () => {
  it('lists the notes whose links land on a file, with the target as written', async () => {
    const root = vault({
      'Index.md': 'See [[Alpha]] and [[Alpha|again]] and [[sub/Beta]]',
      'Alpha.md': 'A [[#Section]] self link',
      'sub/Beta.md': '[[Beta]] by bare name',
      'Lonely.md': 'no links',
    })
    const fs = nodeVaultFs()
    const alpha = await wikiBacklinks(root, join(root, 'Alpha.md'), fs)
    expect(alpha).toEqual([{ source: join(root, 'Index.md'), target: 'Alpha' }])
    const beta = await wikiBacklinks(root, join(root, 'sub/Beta.md'), fs)
    expect(beta).toEqual([{ source: join(root, 'Index.md'), target: 'sub/Beta' }])
    // self links point a note at itself — excluded as noise
    expect(alpha.every((l) => l.source !== join(root, 'Alpha.md'))).toBe(true)
    expect(await wikiBacklinks(root, join(root, 'Lonely.md'), fs)).toEqual([])
  })

  it('a target that climbs out of the vault is no backlink (resolver rule)', async () => {
    const root = vault({
      'sub/A.md': '[[../../outside]]',
      'B.md': '',
    })
    const fs = nodeVaultFs()
    expect(await wikiBacklinks(root, join(root, 'sub/A.md'), fs)).toEqual([])
    // and nothing explodes asking for a file the index never had
    expect((await wikiGraph(root, fs)).edges).toEqual([])
  })

  it('answers nothing for an empty path (an untitled document)', async () => {
    const root = vault({ 'A.md': '[[B]]', 'B.md': '' })
    expect(await wikiBacklinks(root, '', nodeVaultFs())).toEqual([])
  })

  it('picks up an edit after the file mtime moves', async () => {
    const root = vault({ 'A.md': '[[B]]', 'B.md': 'x' })
    const fs = nodeVaultFs()
    const b = join(root, 'B.md')
    expect(await wikiBacklinks(root, b, fs)).toHaveLength(1)
    writeFileSync(join(root, 'A.md'), 'no longer')
    const later = Date.now() / 1000 + 5
    utimesSync(join(root, 'A.md'), later, later)
    expect(await wikiBacklinks(root, b, fs)).toEqual([])
  })

  it('picks up a link whose target did not exist when A was indexed', async () => {
    // A is parsed while B is missing, so its edge memo records "not found" and
    // is keyed only on A's mtime. Creating B changes nothing about A, so without
    // dropping the memo on a file-set change the link stays invisible until A is
    // touched again — the graph never draws A -> B.
    const root = vault({ 'A.md': '[[B]]' })
    const fs = uncachedVaultFs()
    expect(await wikiBacklinks(root, join(root, 'B.md'), fs)).toEqual([])

    writeFileSync(join(root, 'B.md'), 'now it exists')
    expect(await wikiBacklinks(root, join(root, 'B.md'), fs)).toEqual([
      { source: join(root, 'A.md'), target: 'B' },
    ])
  })

  it('drops the edge when a target is renamed away', async () => {
    const root = vault({ 'A.md': '[[B]]', 'B.md': 'x' })
    const fs = uncachedVaultFs()
    expect(await wikiBacklinks(root, join(root, 'B.md'), fs)).toHaveLength(1)

    renameSync(join(root, 'B.md'), join(root, 'C.md'))
    expect(await wikiBacklinks(root, join(root, 'B.md'), fs)).toEqual([])
  })
})

describe('wikiGraph', () => {
  it('a vault inside the node cap is not reported as truncated', async () => {
    const graph = await wikiGraph(vault({ 'A.md': 'x', 'B.md': 'y' }), nodeVaultFs())
    expect(graph.truncated).toBe(false)
    expect(graph.totalNotes).toBe(2)
  })

  it('says so when the node cap cut the vault, and reports the real count', async () => {
    // walk order, not a neighbourhood: the cap takes the first N and the edges
    // only those among them, so a 400-of-3000 draw must not read as the vault
    const files: Record<string, string> = {}
    for (let i = 0; i < MAX_GRAPH_NODES + 5; i += 1) files[`N${i}.md`] = `[[N${i + 1}]]`
    const graph = await wikiGraph(vault(files), nodeVaultFs())
    expect(graph.nodes).toHaveLength(MAX_GRAPH_NODES)
    expect(graph.truncated).toBe(true)
    expect(graph.totalNotes).toBe(MAX_GRAPH_NODES + 5)
    // every edge is between nodes that made the cut, never a dangling target
    const paths = new Set(graph.nodes.map((n) => n.path))
    for (const e of graph.edges) {
      expect(paths.has(e.from)).toBe(true)
      expect(paths.has(e.to)).toBe(true)
    }
  })

  it('every note is a node — orphans included — and links are the edges', async () => {
    const root = vault({
      'A.md': '[[B]]',
      'B.md': '[[A]] and [[B]] self',
      'C.md': 'orphan',
    })
    const graph = await wikiGraph(root, nodeVaultFs())
    expect(graph.nodes.map((n) => n.name).sort()).toEqual(['A', 'B', 'C'])
    expect(graph.edges).toEqual([
      { from: join(root, 'A.md'), to: join(root, 'B.md') },
      { from: join(root, 'B.md'), to: join(root, 'A.md') },
    ])
  })

  it('drops edges that resolve outside the indexed notes or nowhere', async () => {
    const root = vault({
      'A.md': '[[B]] and [[Nowhere]] and ![[pic.png]]',
      'B.md': '',
    })
    const graph = await wikiGraph(root, nodeVaultFs())
    expect(graph.edges).toEqual([{ from: join(root, 'A.md'), to: join(root, 'B.md') }])
  })
})

describe('concurrent builds of one vault', () => {
  it('runs three builds one after another, not side by side', async () => {
    // The panel and the graph open together; both walk the tree and write one
    // shared IndexState. Without the per-root chain they interleave across their
    // awaits and the second walk reads a half-written file set. The delay is
    // what makes interleaving possible at all — a real readdir is fast enough
    // that two builds would rarely actually overlap.
    const files: Record<string, string> = {}
    for (let i = 0; i < 20; i += 1) files[`N${i}.md`] = `[[N${i + 1}]]`
    const root = vault(files)
    const inner = uncachedVaultFs()
    const slow = {
      ...inner,
      list: async (d: string) => {
        await new Promise((r) => setTimeout(r, 1))
        return inner.list(d)
      },
      read: async (p: string) => {
        await new Promise((r) => setTimeout(r, 1))
        return inner.read(p)
      },
    }

    let concurrent = 0
    let peak = 0
    const counting: VaultFs = {
      ...slow,
      list: async (d: string) => {
        peak = Math.max(peak, ++concurrent)
        try {
          return await slow.list(d)
        } finally {
          concurrent -= 1
        }
      },
    }

    const [a, b, c] = await Promise.all([
      wikiGraph(root, counting),
      wikiGraph(root, counting),
      wikiGraph(root, counting),
    ])
    expect(peak).toBe(1)
    for (const g of [a, b, c]) expect(g.nodes).toHaveLength(20)
  })
})
