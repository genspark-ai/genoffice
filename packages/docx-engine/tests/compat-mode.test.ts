/** settings.xml w:compatSetting compatibilityMode: attribute order is generator-specific */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const SETTINGS_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const URI = 'w:uri="http://schemas.microsoft.com/office/word"'

async function compatOf(compatSetting: string) {
  const doc = await parseDocx(
    await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      extraParts: [
        {
          path: 'word/settings.xml',
          xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${SETTINGS_NS}><w:compat>${compatSetting}</w:compat></w:settings>`,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
        },
      ],
    }),
  )
  return doc.compatibilityMode
}

describe('parseCompatibilityMode', () => {
  it('reads w:name before w:val', async () => {
    expect(await compatOf(`<w:compatSetting w:name="compatibilityMode" ${URI} w:val="15"/>`)).toBe(
      15,
    )
  })

  it('reads w:val before w:name', async () => {
    expect(await compatOf(`<w:compatSetting w:val="15" w:name="compatibilityMode" ${URI}/>`)).toBe(
      15,
    )
    expect(await compatOf(`<w:compatSetting ${URI} w:val="14" w:name="compatibilityMode"/>`)).toBe(
      14,
    )
  })

  it('ignores other settings and defaults to 0', async () => {
    expect(
      await compatOf(
        `<w:compatSetting w:val="1" w:name="overrideTableStyleFontSizeAndJustification" ${URI}/>` +
          `<w:compatSetting w:val="0" w:name="enableOpenTypeFeatures" ${URI}/>`,
      ),
    ).toBe(0)
  })
})
