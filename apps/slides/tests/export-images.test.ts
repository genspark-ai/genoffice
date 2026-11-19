import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { buildExportImagePaths } from '../src/main/export-images'

/**
 * The slides export-images handler writes <baseName>-NN.<ext> per rendered page.
 * The zero-padding width follows the total page count (3 digits for ≥100 pages)
 * so file browsers sort the pages naturally.
 */
describe('buildExportImagePaths', () => {
  it('pads to 2 digits below 100 pages', () => {
    expect(buildExportImagePaths('/tmp/out', 'deck', 9, 'png')).toEqual([
      join('/tmp/out', 'deck-01.png'),
      join('/tmp/out', 'deck-02.png'),
      join('/tmp/out', 'deck-03.png'),
      join('/tmp/out', 'deck-04.png'),
      join('/tmp/out', 'deck-05.png'),
      join('/tmp/out', 'deck-06.png'),
      join('/tmp/out', 'deck-07.png'),
      join('/tmp/out', 'deck-08.png'),
      join('/tmp/out', 'deck-09.png'),
    ])
  })

  it('stays at 2 digits for 10 and 99 pages', () => {
    const ten = buildExportImagePaths('/tmp/out', 'deck', 10, 'png')
    expect(ten[0]).toBe(join('/tmp/out', 'deck-01.png'))
    expect(ten[9]).toBe(join('/tmp/out', 'deck-10.png'))
    expect(ten).toHaveLength(10)

    const ninetyNine = buildExportImagePaths('/tmp/out', 'deck', 99, 'png')
    expect(ninetyNine[98]).toBe(join('/tmp/out', 'deck-99.png'))
  })

  it('switches to 3 digits at 100 pages', () => {
    const paths = buildExportImagePaths('/tmp/out', 'deck', 100, 'png')
    expect(paths[0]).toBe(join('/tmp/out', 'deck-001.png'))
    expect(paths[9]).toBe(join('/tmp/out', 'deck-010.png'))
    expect(paths[99]).toBe(join('/tmp/out', 'deck-100.png'))
  })

  it('carries the chosen extension (png and jpg)', () => {
    expect(buildExportImagePaths('/tmp/out', 'deck', 1, 'png')).toEqual([
      join('/tmp/out', 'deck-01.png'),
    ])
    expect(buildExportImagePaths('/tmp/out', 'deck', 2, 'jpg')).toEqual([
      join('/tmp/out', 'deck-01.jpg'),
      join('/tmp/out', 'deck-02.jpg'),
    ])
  })

  it('joins nested directories and yields no paths for an empty deck', () => {
    expect(buildExportImagePaths('/tmp/out', 'a/b', 1, 'png')).toEqual([
      join('/tmp/out', 'a/b-01.png'),
    ])
    expect(buildExportImagePaths('/tmp/out', 'deck', 0, 'png')).toEqual([])
  })
})
