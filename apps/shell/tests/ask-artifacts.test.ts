import { describe, expect, it } from 'vitest'
import {
  artifactFromCall,
  artifactTarget,
  retryInstruction,
  settleArtifact,
} from '../src/renderer/src/ask-artifacts'

const run = (id: string, args: string[]) => ({ id, name: 'run_genoffice', input: { args } })
const create = (id: string, title: string, format: string) => ({
  id,
  name: 'create_document',
  input: { title, format, content: '# x' },
})

describe('ask task card', () => {
  it('names the artifact from --out or the create_document title', () => {
    expect(
      artifactTarget(
        run('1', ['create', '--type', 'docx', '--from', 'r.md', '--out', '/o/Report.docx']),
      ),
    ).toBe('/o/Report.docx')
    expect(
      artifactFromCall(run('1', ['convert', 'a.docx', '--to', 'pdf', '--out', '/o/a.pdf'])),
    ).toEqual({
      callId: '1',
      name: 'a.pdf',
      status: 'building',
    })
    expect(artifactFromCall(create('2', 'Weekly notes', 'md'))).toEqual({
      callId: '2',
      name: 'Weekly notes.md',
      status: 'building',
    })
  })

  it('skips intermediates outside the output directory when it is known', () => {
    const base = run('1', ['create', '--type', 'xlsx', '--from', 't.json', '--out', 'base.xlsx'])
    const final = run('2', [
      'sheet',
      'apply',
      'base.xlsx',
      '--ops',
      'o.json',
      '--out',
      '/o/Summary.xlsx',
    ])
    expect(artifactFromCall(base, '/o')).toBeNull()
    expect(artifactFromCall(final, '/o')).toMatchObject({ name: 'Summary.xlsx' })
    expect(artifactFromCall(base, null)).toMatchObject({ name: 'base.xlsx' })
    expect(artifactFromCall(create('3', 'Notes', 'md'), '/o')).toMatchObject({ name: 'Notes.md' })
  })

  it('ignores reads, guides and folder outputs', () => {
    expect(artifactFromCall(run('1', ['docs', 'read', 'a.docx']))).toBeNull()
    expect(artifactFromCall(run('1', ['render', 'a.docx', '--out', 'shots']))).toBeNull()
    expect(artifactFromCall({ id: '1', name: 'load_guide', input: { domain: 'docs' } })).toBeNull()
    expect(
      artifactFromCall({ id: '1', name: 'read_knowledge_file', input: { index: 0 } }),
    ).toBeNull()
  })

  it('settles to done with the reported path, or to failed with the CLI reason', () => {
    const call = run('7', ['create', '--type', 'xlsx', '--from', 't.json', '--out', 'Summary.xlsx'])
    const started = [artifactFromCall(call)!]
    const done = settleArtifact(started, call, {
      output:
        'exit 0\nsummary: created Summary.xlsx\noutput_path: /Users/me/Documents/Summary.xlsx\ncite: filenav:///x',
    })
    expect(done[0]).toMatchObject({
      status: 'done',
      path: '/Users/me/Documents/Summary.xlsx',
      name: 'Summary.xlsx',
    })
    const failed = settleArtifact(started, call, {
      output: 'exit 2\nsummary: output exists: Summary.xlsx\n{"status":"error"}',
      isError: true,
    })
    expect(failed[0]).toMatchObject({ status: 'failed', error: 'output exists: Summary.xlsx' })
    expect(retryInstruction(failed[0]!)).toContain('output exists')
  })

  it('reads the indented path line create_document prints', () => {
    const call = create('4', 'Weekly notes', 'docx')
    const out = settleArtifact([artifactFromCall(call)!], call, {
      output:
        'saved Weekly notes.docx\n  path: /Users/me/Documents/Weekly notes.docx\n  cite: filenav:///x',
    })
    expect(out[0]).toMatchObject({ status: 'done', path: '/Users/me/Documents/Weekly notes.docx' })
  })

  it('leaves other artifacts alone and tolerates a result for an unknown call', () => {
    const a = artifactFromCall(run('1', ['convert', 'a.md', '--to', 'html', '--out', 'a.html']))!
    const b = artifactFromCall(run('2', ['convert', 'b.md', '--to', 'html', '--out', 'b.html']))!
    const out = settleArtifact([a, b], run('2', []), { output: 'exit 0\noutput_path: /o/b.html' })
    expect(out[0]!.status).toBe('building')
    expect(out[1]).toMatchObject({ status: 'done', path: '/o/b.html' })
    expect(settleArtifact([a], run('9', []), { output: 'x' })).toEqual([a])
  })
})
