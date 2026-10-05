import { describe, expect, it } from 'vitest'

import { collectLaunchPaths } from '../src/main/launch-paths'

describe('collectLaunchPaths', () => {
  it('collects supported argv files in order and removes duplicates', () => {
    expect(
      collectLaunchPaths(
        ['GenOffice.exe', 'first.docx', 'notes.rtf', 'second.xlsx', 'first.docx'],
        undefined,
        () => true,
      ),
    ).toEqual(['first.docx', 'second.xlsx'])
  })

  it('accepts a .tsv from argv or the second-instance payload', () => {
    expect(
      collectLaunchPaths(['GenOffice.app', '/data/variants.tsv'], undefined, () => true),
    ).toEqual(['/data/variants.tsv'])
  })

  /**
   * A `.txt` or `.json` is now a document the text app opens, so dropping one
   * from argv (or from a second-instance payload, or from a Finder open-file)
   * left the shell at Home with no explanation — the file had been accepted as
   * a path and then thrown away before it ever reached the open route.
   */
  it('accepts the text app extensions from argv', () => {
    expect(collectLaunchPaths(['GenOffice.app', '/notes.txt'], undefined, () => true)).toEqual([
      '/notes.txt',
    ])
    expect(collectLaunchPaths(['GenOffice.app', '/data.json'], undefined, () => true)).toEqual([
      '/data.json',
    ])
    expect(collectLaunchPaths(['GenOffice.app', '/NOTES.TXT'], undefined, () => true)).toEqual([
      '/NOTES.TXT',
    ])
    expect(collectLaunchPaths(['GenOffice.app', '/data.json.md'], undefined, () => true)).toEqual([
      '/data.json.md',
    ])
  })

  it('accepts the text app extensions from the second-instance payload', () => {
    expect(
      collectLaunchPaths(['GenOffice.exe'], { launchPaths: ['/notes.txt'] }, () => true),
    ).toEqual(['/notes.txt'])
  })

  it('still drops a path nothing opens', () => {
    // a file that exists but is not a document the shell can route
    expect(collectLaunchPaths(['GenOffice.app', '/photo.png'], undefined, () => true)).toEqual([])
    expect(collectLaunchPaths(['GenOffice.app', '/notes.txt'], undefined, () => false)).toEqual([])
  })

  it('collects argv and second-instance payload files without duplicates', () => {
    expect(
      collectLaunchPaths(
        ['GenOffice.exe', 'first.docx', 'second.pptx'],
        {
          launchPaths: ['first.docx', 'third.pdf', 42, ''],
          launchPath: 'legacy.md',
        },
        () => true,
      ),
    ).toEqual(['first.docx', 'second.pptx', 'third.pdf', 'legacy.md'])
  })

  it('accepts a legacy launchPath payload', () => {
    expect(
      collectLaunchPaths(['GenOffice.exe'], { launchPath: 'legacy.docx' }, () => true),
    ).toEqual(['legacy.docx'])
  })

  it('falls back to the first existing unsupported argv file', () => {
    expect(
      collectLaunchPaths(
        ['GenOffice.exe', 'missing.doc', 'legacy.rtf'],
        undefined,
        (path) => path === 'legacy.rtf',
      ),
    ).toEqual(['legacy.rtf'])
  })
})
