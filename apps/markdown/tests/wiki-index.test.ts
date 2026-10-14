/**
 * The vault link index behind the backlinks panel and the graph view: what
 * counts as a link when a whole vault is scanned, how links invert into
 * backlinks, and that the mtime memo does not serve stale edges.
 */
import { mkdirSync, mkdtempSync, writeFileSync, utimesSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'

import { extractWikiTargets, nodeVaultFs, wikiBacklinks, wikiGraph } from '../src/main/wiki-index'

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
  it('lists the notes whose links land on a file, with the target as written', () => {
    const root = vault({
      'Index.md': 'See [[Alpha]] and [[Alpha|again]] and [[sub/Beta]]',
      'Alpha.md': 'A [[#Section]] self link',
      'sub/Beta.md': '[[Beta]] by bare name',
      'Lonely.md': 'no links',
    })
    const fs = nodeVaultFs()
    const alpha = wikiBacklinks(root, join(root, 'Alpha.md'), fs)
    expect(alpha).toEqual([{ source: join(root, 'Index.md'), target: 'Alpha' }])
    const beta = wikiBacklinks(root, join(root, 'sub/Beta.md'), fs)
    expect(beta).toEqual([{ source: join(root, 'Index.md'), target: 'sub/Beta' }])
    // self links point a note at itself — excluded as noise
    expect(alpha.every((l) => l.source !== join(root, 'Alpha.md'))).toBe(true)
    expect(wikiBacklinks(root, join(root, 'Lonely.md'), fs)).toEqual([])
  })

  it('a target that climbs out of the vault is no backlink (resolver rule)', () => {
    const root = vault({
      'sub/A.md': '[[../../outside]]',
      'B.md': '',
    })
    const fs = nodeVaultFs()
    expect(wikiBacklinks(root, join(root, 'sub/A.md'), fs)).toEqual([])
    // and nothing explodes asking for a file the index never had
    expect(wikiGraph(root, fs).edges).toEqual([])
  })

  it('answers nothing for an empty path (an untitled document)', () => {
    const root = vault({ 'A.md': '[[B]]', 'B.md': '' })
    expect(wikiBacklinks(root, '', nodeVaultFs())).toEqual([])
  })

  it('picks up an edit after the file mtime moves', () => {
    const root = vault({ 'A.md': '[[B]]', 'B.md': 'x' })
    const fs = nodeVaultFs()
    const b = join(root, 'B.md')
    expect(wikiBacklinks(root, b, fs)).toHaveLength(1)
    writeFileSync(join(root, 'A.md'), 'no longer')
    const later = Date.now() / 1000 + 5
    utimesSync(join(root, 'A.md'), later, later)
    expect(wikiBacklinks(root, b, fs)).toEqual([])
  })
})

describe('wikiGraph', () => {
  it('every note is a node — orphans included — and links are the edges', () => {
    const root = vault({
      'A.md': '[[B]]',
      'B.md': '[[A]] and [[B]] self',
      'C.md': 'orphan',
    })
    const graph = wikiGraph(root, nodeVaultFs())
    expect(graph.nodes.map((n) => n.name).sort()).toEqual(['A', 'B', 'C'])
    expect(graph.edges).toEqual([
      { from: join(root, 'A.md'), to: join(root, 'B.md') },
      { from: join(root, 'B.md'), to: join(root, 'A.md') },
    ])
  })

  it('drops edges that resolve outside the indexed notes or nowhere', () => {
    const root = vault({
      'A.md': '[[B]] and [[Nowhere]] and ![[pic.png]]',
      'B.md': '',
    })
    const graph = wikiGraph(root, nodeVaultFs())
    expect(graph.edges).toEqual([{ from: join(root, 'A.md'), to: join(root, 'B.md') }])
  })
})
