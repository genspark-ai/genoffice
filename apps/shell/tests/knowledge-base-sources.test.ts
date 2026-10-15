import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  assembleCorpus,
  indexableCorpusPaths,
  listKnowledgeFiles,
  searchKnowledgeBase,
  sourcesFromFeatures,
  type KbSource,
} from '../src/main/knowledge-base'
import { listKnowledgeEntries, pathForEntry } from '../src/main/feature-knowledge-host'
import type { FeatureRecord } from '@genoffice/agent-core'

/**
 * The knowledge base's sources (§9): the starred corpus plus what the knowledge
 * features register. The reader is the mod host's own (entry cap, traversal
 * guard), a disabled or missing source is never searched but stays visible, and
 * the corpus carries the source each file came from so answers can be honest
 * about coverage.
 */

function feature(id: string, knowledge: FeatureRecord['knowledge']): FeatureRecord {
  return {
    id,
    name: id + ' name',
    description: '',
    kind: 'knowledge',
    ...(knowledge ? { knowledge } : {}),
    createdAt: 'x',
    updatedAt: 'x',
  }
}

function readerFor(_sources: readonly KbSource[]) {
  return {
    readEntries: (source: KbSource) => {
      if (!source.path) return []
      try {
        return listKnowledgeEntries({
          source: source.kind === 'file' ? 'file' : 'folder',
          path: source.path,
        })
      } catch {
        return []
      }
    },
    pathOfEntry: (source: KbSource, entry: string) =>
      pathForEntry(
        { source: source.kind === 'file' ? 'file' : 'folder', path: source.path! },
        entry,
      ),
  }
}

describe('sourcesFromFeatures', () => {
  it('puts the starred corpus first and one source per knowledge feature', () => {
    const sources = sourcesFromFeatures(
      [
        feature('a', { source: 'folder', path: '/nope/missing', label: 'Standards' }),
        feature('b', { source: 'file', path: '/nope/also-missing.md' }),
        feature('c', undefined),
      ],
      new Set(['b']),
    )
    expect(sources.map((source) => source.id)).toEqual(['starred', 'a', 'b'])
    expect(sources[0]).toMatchObject({ kind: 'starred', enabled: true })
    expect(sources[1]).toMatchObject({ kind: 'folder', label: 'Standards', enabled: true })
    expect(sources[2]).toMatchObject({ kind: 'file', enabled: false })
  })

  it('carries the declared index mode, defaulting to full-text', () => {
    const sources = sourcesFromFeatures(
      [
        feature('a', { source: 'folder', path: '/tmp/a', index: 'names' }),
        feature('b', { source: 'folder', path: '/tmp/b' }),
      ],
      new Set(),
    )
    expect(sources.find((source) => source.id === 'a')!.index).toBe('names')
    expect(sources.find((source) => source.id === 'b')!.index).toBe('full-text')
  })

  it('flags a path that does not exist on this machine instead of dropping it', () => {
    const sources = sourcesFromFeatures(
      [feature('a', { source: 'folder', path: '/nope/missing' })],
      new Set(),
    )
    expect(sources.find((source) => source.id === 'a')!.missing).toBe(true)
  })
})

describe('assembleCorpus', () => {
  it('merges the starred files with a folder source, per-source counts included', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-kb-'))
    writeFileSync(join(dir, 'starred.md'), 'starred body', 'utf8')
    const sourceDir = join(dir, 'standards')
    mkdirSync(sourceDir)
    writeFileSync(join(sourceDir, 'style.md'), 'headings are sentence case', 'utf8')
    writeFileSync(join(sourceDir, 'tone.md'), 'short sentences', 'utf8')

    const sources = sourcesFromFeatures(
      [feature('std', { source: 'folder', path: sourceDir, label: 'House style' })],
      new Set(),
    )
    const corpus = assembleCorpus(
      [join(dir, 'starred.md')],
      sources,
      readerFor(sources).readEntries,
      readerFor(sources).pathOfEntry,
    )
    expect(corpus.files.map((file) => file.sourceId)).toEqual(['starred', 'std', 'std'])
    expect(corpus.files[1]).toMatchObject({ sourceLabel: 'House style', name: 'style.md' })
    const std = corpus.sources.find((source) => source.id === 'std')!
    expect(std.documents).toBe(2)
    expect(std.bytes).toBe('headings are sentence case'.length + 'short sentences'.length)
    expect(corpus.sources[0]!.documents).toBe(1)
  })

  it('a disabled source contributes nothing but stays listed with its label', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-kb-'))
    const sourceDir = join(dir, 'standards')
    mkdirSync(sourceDir)
    writeFileSync(join(sourceDir, 'style.md'), 'x', 'utf8')
    const sources = sourcesFromFeatures(
      [feature('std', { source: 'folder', path: sourceDir })],
      new Set(['std']),
    )
    const corpus = assembleCorpus(
      [],
      sources,
      readerFor(sources).readEntries,
      readerFor(sources).pathOfEntry,
    )
    expect(corpus.files).toEqual([])
    expect(corpus.sources.find((source) => source.id === 'std')).toMatchObject({
      enabled: false,
      documents: 0,
    })
  })

  it('a missing source contributes nothing, stays flagged, and breaks nothing', () => {
    const sources = sourcesFromFeatures(
      [feature('gone', { source: 'folder', path: '/nope/missing', label: 'Imported' })],
      new Set(),
    )
    const corpus = assembleCorpus(
      [],
      sources,
      readerFor(sources).readEntries,
      readerFor(sources).pathOfEntry,
    )
    expect(corpus.files).toEqual([])
    expect(corpus.sources.find((source) => source.id === 'gone')).toMatchObject({
      missing: true,
      documents: 0,
    })
  })

  it('a file source contributes its one file, and an unreadable source degrades to empty', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-kb-'))
    writeFileSync(join(dir, 'single.md'), 'one file', 'utf8')
    const sources = sourcesFromFeatures(
      [
        feature('one', { source: 'file', path: join(dir, 'single.md') }),
        feature('bad', { source: 'folder', path: join(dir, 'not-a-folder') }),
      ],
      new Set(),
    )
    const corpus = assembleCorpus(
      [],
      sources,
      readerFor(sources).readEntries,
      readerFor(sources).pathOfEntry,
    )
    expect(corpus.files.map((file) => file.sourceId)).toEqual(['one'])
    expect(corpus.files[0]!.name).toBe('single.md')
  })

  it('the starred list keeps dropping paths that outlive their file', () => {
    const files = listKnowledgeFiles(['/nope/gone.md'])
    expect(files).toEqual([])
  })
})

describe('indexableCorpusPaths', () => {
  it('a names-only source stays out of the index; a disabled one too', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-kb-'))
    writeFileSync(join(dir, 'names-file.md'), 'x', 'utf8')
    writeFileSync(join(dir, 'full-file.md'), 'x', 'utf8')
    const sources = sourcesFromFeatures(
      [
        feature('names-src', { source: 'folder', path: join(dir, 'names-dir'), index: 'names' }),
        feature('full-src', { source: 'folder', path: join(dir, 'full-dir') }),
      ],
      new Set(['off-src']),
    )
    // hand-built corpus with one file per source, plus a disabled source's file
    const corpus = {
      sources,
      files: [
        { ...corpusFileOf('names-src', join(dir, 'names-file.md')) },
        { ...corpusFileOf('full-src', join(dir, 'full-file.md')) },
        { ...corpusFileOf('off-src', join(dir, 'off-file.md')) },
      ],
    }
    expect(indexableCorpusPaths(corpus)).toEqual([join(dir, 'full-file.md')])
  })

  function corpusFileOf(sourceId: string, path: string) {
    return {
      path,
      folder: 'd',
      folderPath: path.slice(0, path.lastIndexOf('/')),
      name: path.split('/').pop()!,
      ext: 'md',
      sizeBytes: 1,
      sourceId,
      sourceLabel: sourceId,
    }
  }
})

describe('searchKnowledgeBase', () => {
  it('filters the index pool to the enabled corpus, whatever source it came from', () => {
    const index = [
      { path: '/corpus/a.md', name: 'a.md', ext: 'md', snippetText: 'alpha beta' },
      { path: '/elsewhere/b.md', name: 'b.md', ext: 'md', snippetText: 'alpha' },
      { path: '/corpus/c.md', name: 'c.md', ext: 'md', snippetText: 'alpha gamma' },
    ]
    const hits = searchKnowledgeBase(() => index, 'alpha', ['/corpus/a.md', '/corpus/c.md'], 8)
    expect(hits.map((hit) => hit.path)).toEqual(['/corpus/a.md', '/corpus/c.md'])
  })
})
