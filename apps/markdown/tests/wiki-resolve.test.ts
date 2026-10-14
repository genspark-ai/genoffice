/**
 * Shortest-path resolution against a real directory tree: the ordering rules
 * are the whole behaviour, so the fixture is a real tree on disk rather than a
 * stubbed fs.
 */
import { mkdtempSync, mkdirSync, writeFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { nodeFs, resolveWikiLinks, splitTarget } from '../src/main/wiki-resolve'

let root = ''

const write = (rel: string) => {
  const p = join(root, rel)
  mkdirSync(join(p, '..'), { recursive: true })
  writeFileSync(p, `# ${rel}\n`)
  return p
}

beforeAll(() => {
  root = mkdtempSync(join(tmpdir(), 'wikiresolve-'))
  // a note beside its own directory, and the same name again one level deeper
  write('Alpha.md')
  write('sub/Alpha.md')
  write('Beta.txt')
  write('Beta.md')
  write('deep/one/two/Gamma.md')
  write('nested/only/Hidden.md')
  // the same name in two subtrees at different depths: `d1` sorts before `s`,
  // so a depth-first walk reaches d1/d2 first and answers wrong
  write('d1/d2/Contested.md')
  write('s/Contested.md')
  // directories a note must never be found through
  write('.git/Config.md')
  write('node_modules/pkg/Pkg.md')
  // an explicit-extension target that is not markdown
  write('assets/Figure 1.png')
})

afterAll(() => rmSync(root, { recursive: true, force: true }))

const resolve = (targets: string[], fromPath: string) =>
  resolveWikiLinks(targets, { fromPath, root }, nodeFs())

describe('splitTarget', () => {
  it('separates the file from a heading or block fragment', () => {
    expect(splitTarget('Note')).toEqual({ name: 'Note', fragment: '' })
    expect(splitTarget('Note#Some Heading')).toEqual({ name: 'Note', fragment: 'Some Heading' })
    expect(splitTarget('Note#^a1b2c3')).toEqual({ name: 'Note', fragment: '^a1b2c3' })
    expect(splitTarget('#Heading')).toEqual({ name: '', fragment: 'Heading' })
  })
})

describe('resolveWikiLinks', () => {
  it('a note in the document directory beats one nested deeper', () => {
    const map = resolve(['Alpha'], join(root, 'sub', 'note.md'))
    expect(map.get('Alpha')).toEqual({
      path: join(root, 'sub', 'Alpha.md'),
      fragment: '',
      found: true,
    })
  })

  it('a shallower note beats a deeper one of the same name', () => {
    // from a directory that holds no Alpha, the root-level one wins over sub/
    const map = resolve(['Alpha'], join(root, 'note.md'))
    expect(map.get('Alpha')?.path).toBe(join(root, 'Alpha.md'))
  })

  it('finds a note several levels down when nothing shallower matches', () => {
    const map = resolve(['Gamma'], join(root, 'note.md'))
    expect(map.get('Gamma')).toEqual({
      path: join(root, 'deep', 'one', 'two', 'Gamma.md'),
      fragment: '',
      found: true,
    })
  })

  it('walks breadth-first: a shallower match in a later subtree still wins', () => {
    // both live under the root and neither is in the document's directory, so
    // only the traversal order decides. `d1` is walked before `s`, and the
    // depth-3 copy sits inside it — a depth-first walk would answer d1/d2.
    const map = resolve(['Contested'], join(root, 'note.md'))
    expect(map.get('Contested')?.path).toBe(join(root, 's', 'Contested.md'))
  })

  it('a same-file link addresses this document and keeps its fragment', () => {
    const self = join(root, 'note.md')
    const map = resolve(['#Some Heading'], self)
    expect(map.get('#Some Heading')).toEqual({ path: self, fragment: 'Some Heading', found: true })
  })

  it('keeps the fragment while resolving the file it points into', () => {
    const map = resolve(['Alpha#^a1b2c3'], join(root, 'note.md'))
    expect(map.get('Alpha#^a1b2c3')).toEqual({
      path: join(root, 'Alpha.md'),
      fragment: '^a1b2c3',
      found: true,
    })
  })

  it('prefers .md over .txt for a bare name', () => {
    const map = resolve(['Beta'], join(root, 'note.md'))
    expect(map.get('Beta')?.path).toBe(join(root, 'Beta.md'))
  })

  it('an explicit extension is taken literally, spaces and all', () => {
    const map = resolve(['Figure 1.png'], join(root, 'note.md'))
    expect(map.get('Figure 1.png')).toEqual({
      path: join(root, 'assets', 'Figure 1.png'),
      fragment: '',
      found: true,
    })
  })

  it('reports an unmatched target instead of guessing', () => {
    const map = resolve(['Nowhere'], join(root, 'note.md'))
    expect(map.get('Nowhere')).toEqual({ path: '', fragment: '', found: false })
  })

  it('never reaches a note through .git or node_modules', () => {
    const map = resolve(['Config', 'Pkg'], join(root, 'note.md'))
    expect(map.get('Config')?.found).toBe(false)
    expect(map.get('Pkg')?.found).toBe(false)
  })

  it('resolves a batch in one pass and reports every link, found or not', () => {
    const map = resolve(['Alpha', 'Nowhere', 'Gamma', '#Self'], join(root, 'note.md'))
    expect([...map.keys()].sort()).toEqual(['#Self', 'Alpha', 'Gamma', 'Nowhere'])
    expect(map.get('Alpha')?.found).toBe(true)
    expect(map.get('Nowhere')?.found).toBe(false)
    expect(map.get('Gamma')?.found).toBe(true)
    expect(map.get('#Self')?.found).toBe(true)
  })

  it('an unreadable root yields unresolved links, not a throw', () => {
    const map = resolveWikiLinks(
      ['Alpha'],
      { fromPath: '', root: join(root, 'no-such-dir') },
      nodeFs(),
    )
    expect(map.get('Alpha')?.found).toBe(false)
  })

  it('an untitled document still resolves against the root', () => {
    const map = resolve(['Alpha'], '')
    expect(map.get('Alpha')).toEqual({ path: join(root, 'Alpha.md'), fragment: '', found: true })
  })
})

describe('staying inside the vault', () => {
  // the reader's disk is not the resolver's business: a shared note must not be
  // able to name a file outside the root by walking up with `..`
  const outside = () => {
    const dir = mkdtempSync(join(tmpdir(), 'wikioutside-'))
    writeFileSync(join(dir, 'outside.txt'), '# outside\n')
    writeFileSync(join(dir, 'budget.xlsx'), '# budget\n')
    return dir
  }

  it('refuses a target that climbs out with ..', () => {
    const dir = outside()
    try {
      const doc = join(dir, 'vault', 'sub', 'doc.md')
      mkdirSync(join(dir, 'vault', 'sub'), { recursive: true })
      writeFileSync(doc, '# doc\n')
      const vault = join(dir, 'vault')
      for (const target of ['../../outside', '../../budget.xlsx']) {
        const map = resolveWikiLinks([target], { fromPath: doc, root: vault }, nodeFs())
        expect(map.get(target), `${target} must not resolve`).toEqual({
          path: '',
          fragment: '',
          found: false,
        })
      }
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('refuses a target that climbs out with a backslash or an absolute path', () => {
    const doc = join(root, 'sub', 'note.md')
    for (const target of ['..\\..\\Alpha', '/etc/passwd', 'sub/../../Alpha', '....//Alpha']) {
      const map = resolve([target], doc)
      expect(map.get(target)?.found, `${target} must not resolve`).toBe(false)
    }
  })

  it('still resolves a plain relative name inside the document directory', () => {
    const map = resolve(['sub/Alpha'], join(root, 'note.md'))
    expect(map.get('sub/Alpha')).toEqual({
      path: join(root, 'sub', 'Alpha.md'),
      fragment: '',
      found: true,
    })
  })
})
