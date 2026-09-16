/**
 * pi-lumos-description regression test.
 *
 * The LumosAI bundled SKILL.md uses every YAML multi-line form (folded `>`,
 * literal `|`, single- and double-quoted scalars) so the YAML-aware description
 * extractor has to handle all of them. This test pins every form so the next
 * refactor does not silently regress.
 */
import { describe, expect, it } from 'vitest'
import { extractLumosDescription } from '../src/shell/pi-resources'

describe('extractLumosDescription', () => {
  it('reads a plain scalar on the same line', () => {
    const body = '---\nname: docs-skill\ndescription: Read and write docx blocks\n---\n# body'
    expect(extractLumosDescription(body, 'FALLBACK')).toBe('Read and write docx blocks')
  })

  it('reads a double-quoted scalar', () => {
    const body = [
      '---',
      'name: translate-pdf',
      'description: "Translate PDF files to Chinese while preserving format"',
      'author: LumosAI',
      '---',
    ].join('\n')
    expect(extractLumosDescription(body, 'FALLBACK')).toBe(
      'Translate PDF files to Chinese while preserving format',
    )
  })

  it('reads a folded block scalar (description: >)', () => {
    const body = [
      '---',
      'name: translate',
      'aliases:',
      '  - translation',
      'description: >',
      '  统一文件翻译入口技能。根据文件扩展名自动选择正确的子技能（translate-pdf、translate-xls、translate-ppt）进行翻译。',
      '  支持 PDF、Excel 和 PowerPoint 文件，翻译时保留原始格式、版面和样式。',
      'command_dispatch: prompt',
      '---',
    ].join('\n')
    const result = extractLumosDescription(body, 'FALLBACK')
    expect(result.startsWith('统一文件翻译入口技能')).toBe(true)
    expect(result.length).toBeGreaterThan(40)
  })

  it('reads a literal block scalar (description: |)', () => {
    const body = [
      '---',
      'name: foo',
      'description: |',
      '  line one',
      '  line two',
      'author: bar',
      '---',
    ].join('\n')
    expect(extractLumosDescription(body, 'FALLBACK')).toBe('line one line two')
  })

  it('falls back when the frontmatter has no description', () => {
    const body = '---\nname: foo\n---\n# body'
    expect(extractLumosDescription(body, 'FALLBACK_NAME')).toBe('FALLBACK_NAME')
  })

  it('falls back when the frontmatter is missing the closing fence', () => {
    const body = 'description: a stray description'
    expect(extractLumosDescription(body, 'FALLBACK')).toBe('FALLBACK')
  })
})

describe('extractLumosDescription (escape sequences)', () => {
  it('unescapes \\" inside a double-quoted scalar to a literal quote', () => {
    const body = [
      '---',
      'name: lumos-self-update',
      'description: "Check updates. Use when action=\\"check\\" returns true."',
      '---',
    ].join('\n')
    expect(extractLumosDescription(body, 'FALLBACK')).toBe(
      'Check updates. Use when action="check" returns true.',
    )
  })

  it('unescapes \\\\ inside a double-quoted scalar to a single backslash', () => {
    const body = [
      '---',
      'name: foo',
      'description: "Path is C:\\\\Users\\\\me"',
      '---',
    ].join('\n')
    expect(extractLumosDescription(body, 'FALLBACK')).toBe('Path is C:\\Users\\me')
  })

  it('collapses doubled apostrophes inside a single-quoted scalar', () => {
    const body = [
      '---',
      'name: foo',
      "description: 'Don''t translate this name'",
      '---',
    ].join('\n')
    expect(extractLumosDescription(body, 'FALLBACK')).toBe("Don't translate this name")
  })
})
