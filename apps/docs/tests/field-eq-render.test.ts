import { describe, expect, it } from 'vitest'
import type { FieldDisplay } from '@genoffice/docx-engine'
import { renderFieldSpec } from '../src/renderer/editor/protected-render'

describe('EQ field result rendering', () => {
  it('the math host carries the run marks and the OMML for buildProtectedDom', () => {
    const field: FieldDisplay = {
      kind: 'text',
      left: 'x = 1/2',
      runs: [
        { text: 'x = ' },
        {
          text: '1/2',
          bold: true,
          color: 'FF0000',
          sizeHalfPoints: 28,
          math: { omml: '<m:oMath><m:f/></m:oMath>' },
        },
      ],
    }
    const spec = renderFieldSpec(field) as unknown[]
    const host = spec[3] as [string, Record<string, string>, string]
    expect(host[0]).toBe('span')
    expect(host[1].class).toBe('doc-inline-math doc-field-math')
    expect(host[1]['data-omml']).toBe('<m:oMath><m:f/></m:oMath>')
    expect(host[1].style).toContain('font-weight:700')
    expect(host[1].style).toContain('font-size:14pt')
    expect(host[1].style).toContain('#FF0000')
    expect(host[2]).toBe('1/2')
  })
})
