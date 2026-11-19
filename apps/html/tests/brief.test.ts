import { describe, expect, it } from 'vitest'
import {
  briefMetaTag,
  briefSummary,
  injectBrief,
  parseBrief,
  type Brief,
} from '../src/renderer/document/brief'

const brief: Brief = {
  core_hook: "Q3 grew 40% while costs fell — here's how",
  style: {
    name: 'Boardroom navy',
    tone: 'executive',
    palette: {
      primary: '#1e3a5f',
      accent: '#c9a227',
      bg: '#ffffff',
      surface: '#f4f6fa',
      text: '#1f2328',
    },
    typography: { heading: 'Inter', body: 'Inter', scale: 'regular' },
    tokens: { radius: '6px', spacing: '8px', shadow: 'none' },
    layout: 'single column 760px',
    density: 'regular',
    docx_friendly: true,
  },
  alternatives: [
    {
      name: 'Dark tech',
      tone: 'technical',
      palette: { primary: '#7dd3fc', accent: '#f472b6', bg: '#0b1020', text: '#e5e7eb' },
      typography: { heading: 'Space Grotesk', body: 'Inter' },
    },
  ],
  sections: [
    { title: 'Highlights', type: 'summary', brief: 'three KPIs', layout: 'stats_row' },
    { title: 'Revenue', type: 'data', brief: 'table by region', image_queries: [] },
  ],
  meta: { title: 'Q3 Review', language: 'en', audience: 'board' },
  user_edited: ['core_hook'],
  version: 1,
}

describe('brief persistence', () => {
  it('round-trips through the meta tag with quotes and angle brackets escaped', () => {
    const tag = briefMetaTag(brief)
    expect(tag.startsWith(`<meta name="genoffice:brief" content='`)).toBe(true)
    expect(tag).not.toContain("here's") // the apostrophe inside the single-quoted attribute is escaped
    const back = parseBrief(`<html><head>${tag}</head></html>`)
    expect(back?.core_hook).toBe(brief.core_hook)
    expect(back?.sections).toHaveLength(2)
    expect(back?.style.tokens?.radius).toBe('6px')
    expect(back?.style.palette.surface).toBe('#f4f6fa')
    expect(back?.user_edited).toBeUndefined()
  })

  it('pins only the chosen direction, never the alternatives', () => {
    expect(briefMetaTag(brief)).not.toContain('Dark tech')
    expect(parseBrief(`<head>${briefMetaTag(brief)}</head>`)?.alternatives).toBeUndefined()
  })

  it('injects after <meta charset>, else after <head>, else creates a head, and refreshes in place', () => {
    const a = injectBrief('<html><head><meta charset="utf-8"><title>x</title></head></html>', brief)
    expect(a).toMatch(/<meta charset="utf-8">\n<meta name="genoffice:brief"/)
    const b = injectBrief('<html><head><title>x</title></head></html>', brief)
    expect(b).toMatch(/<head>\n<meta name="genoffice:brief"/)
    const c = injectBrief('<html><body></body></html>', brief)
    expect(c).toMatch(/<html>\n<head><meta name="genoffice:brief"/)
    const d = injectBrief('<p>frag</p>', brief)
    expect(d.startsWith('<meta name="genoffice:brief"')).toBe(true)
    const refreshed = injectBrief(a, { ...brief, core_hook: 'changed' })
    expect((refreshed.match(/genoffice:brief/g) ?? []).length).toBe(1)
    expect(parseBrief(refreshed)?.core_hook).toBe('changed')
  })

  it('ignores malformed meta and summarizes compactly', () => {
    expect(parseBrief(`<meta name="genoffice:brief" content='{not json'>`)).toBeNull()
    expect(parseBrief(`<meta name="genoffice:brief" content='{"core_hook":1}'>`)).toBeNull()
    const s = briefSummary(brief)
    expect(s).toContain('core hook: Q3 grew 40%')
    expect(s).toContain('style: Boardroom navy — executive')
    expect(s).toContain('surface=#f4f6fa')
    expect(s).toContain('tokens: radius=6px spacing=8px shadow=none')
    expect(s).toContain('sections: Highlights · Revenue')
    expect(s).toContain('docx-friendly')
    expect(s).not.toContain('Dark tech')
  })
})
