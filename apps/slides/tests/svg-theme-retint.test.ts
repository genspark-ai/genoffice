import { describe, expect, it, vi } from 'vitest'
import type { OpenedPptx } from '@genoffice/pptx-engine'
import { retintThemedSvg } from '../src/main/session-state'

vi.mock('electron', () => ({
  BrowserWindow: { getFocusedWindow: () => null },
  webContents: { getAllWebContents: () => [] },
}))
vi.mock('../src/main/fonts', () => ({
  createSystemFontMetrics: () => ({}),
}))

const THEME_XML =
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="t">' +
  '<a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
  '</a:clrScheme></a:themeElements></a:theme>'

function openedWith(themeXml: string, masterXml?: string): OpenedPptx {
  return {
    deck: { slides: [{ path: 'ppt/slides/slide1.xml' }], size: { cx: 1, cy: 1 } },
    archive: {
      resolveSlideChain: () => ({
        themePath: 'ppt/theme/theme1.xml',
        masterPath: masterXml ? 'ppt/slideMasters/slideMaster1.xml' : undefined,
      }),
      readText: (p: string) => (p === 'ppt/theme/theme1.xml' ? themeXml : (masterXml ?? null)),
    },
  } as unknown as OpenedPptx
}

describe('retintThemedSvg (MsftOfcThm_* classes follow the live theme like PowerPoint)', () => {
  it('rewrites a Background1 fill class to the theme bg1 color', () => {
    const svg =
      '<svg><style>.MsftOfcThm_Background1_Fill_v2 {\n fill:#DEF3FF; \n}</style>' +
      '<path class="MsftOfcThm_Background1_Fill_v2" fill="#9EF5EF"/></svg>'
    const out = retintThemedSvg(svg, openedWith(THEME_XML))
    expect(out).toContain('fill:#FFFFFF')
    expect(out).not.toContain('#DEF3FF')
    // the presentation attribute stays (CSS class wins in the cascade anyway)
    expect(out).toContain('fill="#9EF5EF"')
  })

  it('rewrites accent stroke classes and leaves unknown classes alone', () => {
    const svg =
      '<style>.MsftOfcThm_Accent2_Stroke_v2 { stroke:#000000; }' +
      '.MsftOfcThm_Mystery_Fill { fill:#123456; }</style>'
    const out = retintThemedSvg(svg, openedWith(THEME_XML))
    expect(out).toContain('stroke:#ED7D31')
    expect(out).toContain('fill:#123456')
  })

  it('honors the master clrMap (dark master flips bg1 to dk1)', () => {
    const master =
      '<p:sldMaster xmlns:p="x"><p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>'
    const svg = '<style>.MsftOfcThm_Background1_Fill { fill:#ABCDEF; }</style>'
    const out = retintThemedSvg(svg, openedWith(THEME_XML, master))
    expect(out).toContain('fill:#111111')
  })

  it('returns the svg unchanged when the theme is unreadable', () => {
    const opened = {
      deck: { slides: [{ path: 'p' }], size: { cx: 1, cy: 1 } },
      archive: { resolveSlideChain: () => ({}), readText: () => null },
    } as unknown as OpenedPptx
    const svg = '<style>.MsftOfcThm_Background1_Fill { fill:#ABCDEF; }</style>'
    expect(retintThemedSvg(svg, opened)).toBe(svg)
  })
})
