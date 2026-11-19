import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const framed = (attrs: string): string =>
  `<w:p><w:pPr><w:framePr ${attrs}/><w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`

async function frameBoxOf(bodyXml: string) {
  const doc = await parseDocx(await buildDocx({ bodyXml }))
  return doc.blocks[0].format?.frameBox
}

describe('w:framePr frame box (display-only)', () => {
  it('reads the size and floats a text-anchored wrapping frame', async () => {
    expect(
      await frameBoxOf(
        framed(
          'w:w="1440" w:h="1440" w:hSpace="180" w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:y="1"',
        ),
      ),
    ).toEqual({ wTwips: 1440, hTwips: 1440, hSpaceTwips: 180, floatSide: 'left' })
  })

  it('keeps exact heights and right-aligned frames on their side', async () => {
    expect(
      await frameBoxOf(
        framed(
          'w:w="2000" w:h="600" w:hRule="exact" w:wrap="around" w:vAnchor="text" w:xAlign="right"',
        ),
      ),
    ).toEqual({ wTwips: 2000, hTwips: 600, hRule: 'exact', floatSide: 'right' })
  })

  it('does not float page-anchored or non-wrapping frames, and skips drop caps', async () => {
    expect(
      await frameBoxOf(
        framed('w:w="4000" w:x="1200" w:y="800" w:hAnchor="page" w:vAnchor="page" w:wrap="none"'),
      ),
    ).toEqual({ wTwips: 4000 })
    expect(await frameBoxOf(framed('w:w="1440" w:wrap="notBeside" w:vAnchor="text"'))).toEqual({
      wTwips: 1440,
    })
    expect(
      await frameBoxOf(
        framed('w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"'),
      ),
    ).toBeUndefined()
    expect(await frameBoxOf(framed('w:wrap="around" w:vAnchor="text"'))).toBeUndefined()
  })
})
