import { describe, expect, it } from 'vitest'
import { buildBlankDocx, parseDocx, readSectionSettings } from '@genoffice/docx-engine'
import { withDesignDefault } from '../src/renderer/file-actions'
import { defaultParagraphStyleId } from '../src/renderer/style-gallery'

describe('Set as Default on a new blank document', () => {
  it('leaves the blank document alone when no look was kept', async () => {
    const blank = await buildBlankDocx({})
    expect(await withDesignDefault(blank, null)).toBe(blank)
    expect(await withDesignDefault(blank, {})).toBe(blank)
  })

  it('writes the kept fonts and paragraph spacing into the document itself', async () => {
    const blank = await buildBlankDocx({})
    const styled = await withDesignDefault(blank, {
      fonts: { major: 'Georgia', minor: 'Verdana', eastAsia: 'SimSun' },
      spacing: { beforeTwips: 120, afterTwips: 160, line: 2 },
    })
    const parsed = await parseDocx(styled)
    expect(parsed.themeFonts).toMatchObject({ major: 'Georgia', minor: 'Verdana' })
    const normal = parsed.styles.get(defaultParagraphStyleId(parsed.styles) ?? 'Normal')
    expect(normal?.display?.spaceBeforeTwips).toBe(120)
    expect(normal?.display?.spaceAfterTwips).toBe(160)
    expect(normal?.display?.lineSpacing).toBe(2)
    // still a blank document
    expect(parsed.blocks).toHaveLength((await parseDocx(blank)).blocks.length)
  })

  it('writes the kept page setup into the document, and skips one the paper cannot hold', async () => {
    const blank = await buildBlankDocx({})
    const page = {
      section: {
        marginTop: 2000,
        marginRight: 1500,
        marginBottom: 1800,
        marginLeft: 2200,
        headerDist: 600,
        footerDist: 500,
        gutter: 300,
        vAlign: 'center' as const,
        columns: 2,
      },
      mirror: true,
      titlePg: true,
      evenOdd: false,
    }
    const parsed = await parseDocx(await withDesignDefault(blank, { page }))
    expect(readSectionSettings(parsed)).toMatchObject(page.section)
    expect(parsed.mirrorMargins).toBe(true)
    expect(parsed.titlePg).toBe(true)
    expect(parsed.evenAndOddHeaders ?? false).toBe(false)

    const before = readSectionSettings(await parseDocx(blank))
    const tooWide = { ...page, section: { ...page.section, marginLeft: before.pageWidth } }
    const kept = readSectionSettings(
      await parseDocx(await withDesignDefault(blank, { page: tooWide })),
    )
    expect(kept.marginLeft).toBe(before.marginLeft)
    expect(kept.columns).toBe(before.columns)
  })

  it('writes the kept font into the default paragraph style, next to kept spacing', async () => {
    const blank = await buildBlankDocx({})
    const parsed = await parseDocx(
      await withDesignDefault(blank, {
        font: { latin: 'Georgia', sizeHalfPoints: 28, italic: true, color: '1F4E79' },
        spacing: { afterTwips: 160, line: 2 },
      }),
    )
    const normal = parsed.styles.get(defaultParagraphStyleId(parsed.styles) ?? 'Normal')
    expect(normal?.display).toMatchObject({
      fontAscii: 'Georgia',
      sizeHalfPoints: 28,
      italic: true,
      color: '1F4E79',
      spaceAfterTwips: 160,
    })
    expect(normal?.display?.bold).toBeFalsy()
  })
})
